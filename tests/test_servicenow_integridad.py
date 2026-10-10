import json
import re
from datetime import date
from unittest import TestCase
from unittest.mock import MagicMock, patch

import app_logic
from repositories import database
from repositories import servicenow as repository
from services.servicenow import ServiceNowClient, ServiceNowError
from services.servicenow_diagnostico import diagnosticar_fallo
from tests.test_servicenow import caso, documento, field, response


class IntegridadConsultaTest(TestCase):
    def test_limite_del_servidor_de_24_no_salta_casos_del_mes(self):
        config = {"SERVICENOW_INSTANCE": "test.service-now.com", "SERVICENOW_USER": "test",
                  "SERVICENOW_PASSWORD": "test"}
        rows = []
        for i in range(664):
            row = caso()
            row.update(sys_id=field(f"{i + 1:032x}"), number=field(f"CS{i + 1000000}"))
            rows.append(row)

        def consultar(url, **kw):
            if not url.endswith("/sn_customerservice_case"):
                return response([], 0)
            params = kw["params"]
            cursor = re.search(r"\^sys_id>([0-9a-f]{32})", params["sysparm_query"])
            filtrados = [r for r in rows if not cursor or r["sys_id"]["value"] > cursor[1]]
            offset = params["sysparm_offset"]
            return response(filtrados[offset:offset + 24], len(filtrados))

        with patch("services.servicenow.config_value", side_effect=lambda k, d="": config.get(k, d)), \
             ServiceNowClient() as cliente, patch.object(cliente._session, "get", side_effect=consultar) as get:
            docs = [d for pagina in cliente.listar_mes("2026-10") for d in pagina]
        self.assertEqual(664, len(docs))
        self.assertEqual(664, len({d["case_id"] for d in docs}))
        self.assertEqual(664, cliente.total_registros)
        casos = [c for c in get.call_args_list if c.args[0].endswith("/sn_customerservice_case")]
        self.assertEqual(28, len(casos))
        self.assertTrue(all(c.kwargs["params"]["sysparm_offset"] == 0 for c in casos))
        self.assertIn("^sys_id>", casos[1].kwargs["params"]["sysparm_query"])

    def test_total_664_con_solo_24_visibles_se_guarda_pero_no_se_anuncia_completo(self):
        with patch.object(app_logic, "exigir_permiso_actor"), \
             patch.object(app_logic, "get_conn", return_value=MagicMock()), \
             patch.object(app_logic, "_esperar_turno_manual", return_value=True), \
             patch.object(app_logic, "_registrar_estado_ingesta") as estado, \
             patch.object(app_logic, "ServiceNowClient") as constructor, \
             patch.object(app_logic, "_guardar_pagina_servicenow", return_value=(24, 0)) as guardar:
            cliente = constructor.return_value.__enter__.return_value
            cliente.total_registros = 664
            cliente.listar_rango.return_value = iter([[documento()] * 24])
            with self.assertRaisesRegex(ServiceNowError, "664.*24"):
                app_logic.sincronizar_servicenow(actor_email="admin", fecha_inicio=date.today(), fecha_fin=date.today())
        guardar.assert_called_once()
        self.assertEqual("error", estado.call_args.kwargs["estado"])
        self.assertEqual(24, estado.call_args.kwargs["registros"])
        self.assertFalse(any(c.kwargs["estado"] == "correcto" for c in estado.call_args_list))


class GuardadoPaginaTest(TestCase):
    def test_cien_casos_y_sus_conversaciones_se_guardan_en_dos_operaciones(self):
        docs = []
        filas = []
        for i in range(100):
            doc = documento()
            doc["case_id"] = f"{i + 1:032x}"
            historia = [{"id": f"{i + 1000:032x}", "contenido": "Contenido completo " * 100 + str(i)}]
            doc.update(historial_conversacion=historia, traza_unificada=historia)
            docs.append(doc)
            filas.append({"numero": f"CS{i + 1000000}", "descripcion": f"Caso {i}", "historial_conversacion": historia[0]["contenido"]})
        with patch.object(repository, "db_execute") as ejecutar:
            repository.guardar_documentos_pagina(MagicMock(), docs, "cases", filas, "admin")
        self.assertEqual(2, ejecutar.call_count)
        self.assertEqual(filas, json.loads(ejecutar.call_args_list[0].args[2][0]))
        snapshots = json.loads(ejecutar.call_args_list[1].args[2][0])
        self.assertEqual(100, len(snapshots))
        for i, snapshot in enumerate(snapshots):
            self.assertEqual(filas[i]["numero"], snapshot["numero"])
            self.assertEqual(docs[i]["historial_conversacion"], snapshot["documento"]["historial_conversacion"])
            self.assertNotIn("traza_unificada", snapshot["documento"])
        self.assertTrue(all("WHERE TRUE ON CONFLICT" in " ".join(c.args[1].split()) for c in ejecutar.call_args_list))

    def test_campos_omitidos_y_campos_vacios_no_se_mezclan_entre_filas(self):
        filas = [{"numero": "CS1000001", "estado": "Nuevo"},
                 {"numero": "CS1000002", "estado": "Nuevo", "historial_conversacion": "", "producto": None}]
        with patch.object(repository, "db_execute") as ejecutar:
            repository.guardar_documentos_pagina(MagicMock(), [documento(), documento()], "cases", filas, "admin")
        self.assertEqual(3, ejecutar.call_count)
        self.assertNotIn("historial_conversacion", ejecutar.call_args_list[0].args[1])
        self.assertIn("historial_conversacion = EXCLUDED.historial_conversacion", ejecutar.call_args_list[1].args[1])
        self.assertEqual(filas[1], json.loads(ejecutar.call_args_list[1].args[2][0])[0])
        self.assertIn("COALESCE(servicenow_records.documento", ejecutar.call_args_list[2].args[1])

    def test_conexion_no_prepara_consultas_que_el_pooler_no_puede_conservar(self):
        with patch.object(database, "SUPABASE_DATABASE_URL", "postgresql://test"), patch("psycopg.connect") as conectar:
            database.get_conn()
        self.assertIsNone(conectar.call_args.kwargs["prepare_threshold"])


class DiagnosticoServidorTest(TestCase):
    def test_error_muestra_etapa_codigo_y_cantidad_sin_exponer_datos_en_pantalla_o_logs(self):
        class FalloDB(Exception):
            sqlstate = "26000"

        with self.assertLogs("services.servicenow_diagnostico", level="ERROR") as log:
            error = diagnosticar_fallo(FalloDB("postgresql://user:secret@host contenido privado"),
                                      etapa="guardar la página en la base", guardados=24, total=664)
        mensaje = str(error)
        self.assertIn("24 de 664", mensaje)
        self.assertIn("26000", mensaje)
        self.assertIn("guardar la página", mensaje)
        self.assertIn("Referencia:", mensaje)
        for texto in [mensaje, *log.output]:
            self.assertNotIn("secret", texto)
            self.assertNotIn("contenido privado", texto)
        self.assertIn("FalloDB", log.output[0])

    def test_fallo_al_guardar_se_revela_aunque_tambien_falle_guardar_estado_o_liberar_turno(self):
        class FalloDB(Exception):
            sqlstate = "55P03"

        conn = MagicMock()
        conn.rollback.side_effect = RuntimeError("secret de conexión rota")
        with patch.object(app_logic, "exigir_permiso_actor"), \
             patch.object(app_logic, "get_conn", return_value=conn), \
             patch.object(app_logic, "_esperar_turno_manual", return_value=True), \
             patch.object(app_logic, "_registrar_estado_ingesta", side_effect=[None, RuntimeError("secret de estado")]), \
             patch.object(app_logic, "ServiceNowClient") as constructor, \
             patch.object(app_logic, "_guardar_pagina_servicenow", side_effect=FalloDB("secret de dato")), \
             self.assertLogs("services.servicenow_diagnostico", level="ERROR") as log:
            cliente = constructor.return_value.__enter__.return_value
            cliente.total_registros = 664
            cliente.listar_rango.return_value = iter([[documento()]])
            with self.assertRaisesRegex(ServiceNowError, "55P03") as error:
                app_logic.sincronizar_servicenow(actor_email="admin", fecha_inicio=date.today(), fecha_fin=date.today())
        self.assertIn("0 de 664", str(error.exception))
        self.assertIn("guardar la página en la base", str(error.exception))
        self.assertTrue(all("secret" not in texto for texto in [str(error.exception), *log.output]))
        conn.close.assert_called_once()
