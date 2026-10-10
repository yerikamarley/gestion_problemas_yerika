"""Lecturas simultáneas con historial íntegro y cancelación coordinada."""

from threading import Barrier, Event, Lock, get_ident
from concurrent.futures import wait as esperar
from unittest import TestCase
from unittest.mock import patch

import app_logic
from services import servicenow
from services.servicenow import ServiceNowClient, ServiceNowError
from tests.test_servicenow import caso, field, response


class FuentesParalelasTest(TestCase):
    def setUp(self):
        self.principal = get_ident()
        self.settings = {"SERVICENOW_AUTH_TYPE": "client_credentials", "SERVICENOW_INSTANCE": "test.service-now.com",
                         "SERVICENOW_CLIENT_ID": "client-test", "SERVICENOW_CLIENT_SECRET": "secret-test"}
        self.hilos_config = []

        def config(key, default=""):
            self.hilos_config.append(get_ident())
            return self.settings.get(key, default)

        configuracion = patch("services.servicenow.config_value", side_effect=config)
        configuracion.start()
        self.addCleanup(configuracion.stop)
        servicenow._tokens_oauth.clear()
        self.addCleanup(servicenow._tokens_oauth.clear)
        token = response([])
        token.json.return_value = {"access_token": "token-test", "expires_in": 1800, "token_type": "Bearer"}
        autenticacion = patch("requests.sessions.Session.post", return_value=token)
        self.post = autenticacion.start()
        self.addCleanup(autenticacion.stop)

    def test_tres_fuentes_simultaneas_sesiones_independientes_y_ui_en_hilo_principal(self):
        entrada = Barrier(3)
        sesiones = {}
        hilos_ui = []
        mutex = Lock()
        case_id = caso()["sys_id"]["value"]

        def consultar(session, url, **kw):
            fuente = url.rsplit("/", 1)[-1]
            with mutex:
                sesiones[fuente] = session
            entrada.wait(timeout=3)  # Solo avanza si las tres lecturas ya empezaron.
            fila = {"sys_id": "b" * 32, "sys_created_on": "2026-10-01 11:00:00"}
            if fuente == "task_sla":
                fila.update(task=case_id, sla="SLA prueba", stage="in_progress", has_breached="false")
            elif fuente == "sys_journal_field":
                fila.update(element_id=case_id, element="work_notes", value="Nota completa")
            else:
                fila.update(instance=case_id, type="received", body_text="Correo completo", origemail="test@example.test")
            return response([fila], 1)

        with ServiceNowClient(progress_callback=lambda *_: hilos_ui.append(get_ident())) as cliente, \
             patch("requests.sessions.Session.get", autospec=True, side_effect=consultar):
            documentos = cliente._consolidar_pagina([caso()])
            self.assertEqual(3, len({id(session) for session in sesiones.values()}))
            self.assertTrue(all(session is not cliente._session for session in sesiones.values()))
            self.assertEqual(3, cliente._executor._max_workers)
        self.assertEqual(1, self.post.call_count)
        self.assertEqual({self.principal}, set(self.hilos_config))
        self.assertEqual({self.principal}, set(hilos_ui))
        self.assertEqual(["Nota completa", "Correo completo"],
                         [r["contenido"] for r in documentos[0]["historial_conversacion"]])
        self.assertEqual(1, len(documentos[0]["slas"]))
        self.assertEqual("codigo", documentos[0]["campos_raw"]["u_campo_extra"]["value"])

    def test_sesiones_y_conexiones_http_se_reutilizan_entre_lotes(self):
        sesiones = {}
        mutex = Lock()

        def consultar(session, url, **kw):
            fuente = url.rsplit("/", 1)[-1]
            with mutex:
                sesiones.setdefault(fuente, []).append(session)
            return response([], 0)

        with ServiceNowClient() as cliente, patch("requests.sessions.Session.get", autospec=True, side_effect=consultar):
            docs1 = cliente._consolidar_pagina([caso()])
            docs2 = cliente._consolidar_pagina([caso()])
        self.assertEqual(1, self.post.call_count)
        self.assertEqual(3, len(sesiones))
        for llamadas in sesiones.values():
            self.assertEqual(2, len(llamadas))
            self.assertIs(llamadas[0], llamadas[1])
        self.assertEqual(docs1, docs2)

    def test_fallo_de_correo_no_entrega_historial_parcial_y_libera_todos_los_lectores(self):
        case_id = caso()["sys_id"]["value"]

        def consultar(session, url, **kw):
            if url.endswith("/sys_email"):
                return response([], status=403)
            return response([{"sys_id": "b" * 32, "element_id": case_id, "element": "work_notes", "value": "Nota"}], 1)

        cierres = []
        with patch("requests.sessions.Session.get", autospec=True, side_effect=consultar), \
             patch("requests.sessions.Session.close", autospec=True, side_effect=lambda session: cierres.append(session)), \
             ServiceNowClient() as cliente:
            with self.assertRaisesRegex(ServiceNowError, "sys_email.*403"):
                cliente._consolidar_pagina([caso()])
        self.assertEqual(4, len(cierres))
        self.assertEqual(4, len({id(session) for session in cierres}))

    def test_prioridad_manual_cancela_siguientes_paginas_y_callbacks_permanecen_en_hilo_principal(self):
        iniciadas = Barrier(3)
        listas = Event()
        liberar = Event()
        llamadas = []
        mutex = Lock()
        hilos_ui = []

        def consultar(session, url, **kw):
            with mutex:
                llamadas.append(url)
            iniciadas.wait(timeout=3)
            listas.set()
            if not liberar.wait(timeout=3):
                raise AssertionError("Las consultas pendientes no se liberaron al cancelar.")
            return response([{"sys_id": "b" * 32, "body_text": "Texto completo"}], 101)

        def progreso(*_):
            hilos_ui.append(get_ident())
            self.assertTrue(listas.wait(timeout=3))
            raise app_logic._TurnoManualSolicitado

        def esperar_consultas(futuros, **kwargs):
            if not kwargs:  # La limpieza comienza después de señalar la cancelación.
                liberar.set()
            return esperar(futuros, **kwargs)

        with ServiceNowClient(progress_callback=progreso) as cliente, \
             patch("requests.sessions.Session.get", autospec=True, side_effect=consultar), \
             patch("services.servicenow.wait", side_effect=esperar_consultas):
            with self.assertRaises(app_logic._TurnoManualSolicitado):
                cliente._consolidar_pagina([caso()])
        self.assertEqual(3, len(llamadas))
        self.assertEqual({self.principal}, set(hilos_ui))

    def test_correos_html_y_reintento_400_conservan_mensajes_en_paralelo(self):
        raws = []
        emails = []
        for i in range(4):
            raw = caso()
            raw.update(sys_id=field(f"{i + 1:032x}"), number=field(f"CS{i + 1000000}"))
            raws.append(raw)
            emails.append({"sys_id": f"{i + 1000:032x}", "instance": raw["sys_id"]["value"],
                           "type": "received", "body_text": ""})

        def consultar(session, url, **kw):
            if not url.endswith("/sys_email"):
                return response([], 0)
            query = kw["params"]["sysparm_query"]
            if kw["params"]["sysparm_fields"] == "sys_id,body":
                ids = query.split("sys_idIN")[1].split("^")[0].split(",")
                return response([{"sys_id": row["sys_id"], "body": f"<p>Mensaje completo {i}</p>"}
                                 for i, row in enumerate(emails) if row["sys_id"] in ids], len(ids))
            ids = query.split("instanceIN")[1].split("^")[0].split(",")
            if len(ids) > 2:
                return response([], status=400)
            return response([row for row in emails if row["instance"] in ids], len(ids))

        with ServiceNowClient() as cliente, patch("requests.sessions.Session.get", autospec=True, side_effect=consultar):
            docs = cliente._consolidar_pagina(raws)
        self.assertEqual(4, len(docs))
        self.assertEqual([f"Mensaje completo {i}" for i in range(4)],
                         [doc["historial_conversacion"][0]["contenido"] for doc in docs])
        self.assertEqual(1, self.post.call_count)
