from datetime import date, timedelta
import unittest
from unittest.mock import MagicMock, patch

import pandas as pd

import app_logic
import app_ui
from services.servicenow import ServiceNowClient, ServiceNowError, limites_rango_utc, SOURCE_LABELS
from tests.test_servicenow import caso, documento, field, response


class ConsultaServiceNowTest(unittest.TestCase):
    def setUp(self):
        config = {"SERVICENOW_INSTANCE": "test.service-now.com", "SERVICENOW_USER": "api-test",
                  "SERVICENOW_PASSWORD": "password-test"}
        parche = patch("services.servicenow.config_value", side_effect=lambda k, d="": config.get(k, d))
        parche.start()
        self.addCleanup(parche.stop)

    def test_rango_inclusivo_colombiano_cruza_mes_anio_y_dia_bisiesto(self):
        for inicio, fin, esperado in [
            (date(2026, 10, 10), date(2026, 10, 10), ("2026-10-10 05:00:00", "2026-10-11 05:00:00")),
            (date(2026, 10, 31), date(2026, 11, 1), ("2026-10-31 05:00:00", "2026-11-02 05:00:00")),
            ("2026-12-31", "2027-01-01", ("2026-12-31 05:00:00", "2027-01-02 05:00:00")),
            ("2024-02-29", "2024-02-29", ("2024-02-29 05:00:00", "2024-03-01 05:00:00")),
        ]:
            with self.subTest(inicio=inicio):
                self.assertEqual(esperado, limites_rango_utc(inicio, fin))
        for inicio, fin in [(None, date.today()), ("2026-10-10", None), ("2026-10-10", "2026-10-09"), ("2026-02-30", "2026-03-01")]:
            with self.subTest(inicio=inicio), self.assertRaises(ValueError):
                limites_rango_utc(inicio, fin)

    def test_rango_pagina_todos_los_casos_e_informa_antes_de_cada_consulta(self):
        mensajes = []
        raw = caso()
        with ServiceNowClient(progress_callback=lambda _, texto: mensajes.append(texto)) as client:
            def consultar(url, **kwargs):
                tabla = url.rsplit("/", 1)[1]
                self.assertIn(SOURCE_LABELS[tabla], mensajes[-1])
                return response([raw], 1) if tabla == "sn_customerservice_case" else response([], 0)
            with patch.object(client._session, "get", side_effect=consultar) as get:
                docs = [d for page in client.listar_rango(date(2026, 10, 9), date(2026, 10, 10)) for d in page]
        self.assertEqual(1, len(docs))
        query = get.call_args_list[0].kwargs["params"]["sysparm_query"]
        self.assertIn("sys_created_on>=2026-10-09 05:00:00^sys_created_on<2026-10-11 05:00:00", query)
        self.assertEqual(4, get.call_count)
        self.assertTrue(any("correos del historial" in texto for texto in mensajes))
        self.assertFalse(any("password-test" in texto for texto in mensajes))

    def test_correos_texto_completos_y_html_solo_si_hace_falta(self):
        texto = "Detalle importante del cliente " * 3000 + " FIN TEXTO"
        html = "Respuesta de soporte " * 3000 + " &amp; FIN HTML"
        correos = [
            {"sys_id": "c" * 32, "instance": "a" * 32, "type": "received", "body_text": texto,
             "sys_created_on": "2026-10-09 23:20:38"},
            {"sys_id": "d" * 32, "instance": "a" * 32, "type": "sent", "body_text": "",
             "sys_created_on": "2026-10-09 23:20:48"},
        ]
        def consultar(url, **kwargs):
            campos = kwargs["params"].get("sysparm_fields", "").split(",")
            if not url.endswith("/sys_email"):
                return response([], 0)
            self.assertEqual("false", kwargs["params"]["sysparm_display_value"])
            if campos == ["sys_id", "body"]:
                self.assertIn("d" * 32, kwargs["params"]["sysparm_query"])
                self.assertNotIn("c" * 32, kwargs["params"]["sysparm_query"])
                return response([{"sys_id": "d" * 32, "body": f"<p>{html}</p>"}], 1)
            self.assertIn("body_text", campos)
            self.assertNotIn("body", campos)
            return response(correos, 2)
        with ServiceNowClient() as client, patch.object(client._session, "get", side_effect=consultar) as get:
            doc = client.consolidar(caso())
        self.assertEqual(4, get.call_count)
        self.assertEqual(texto, doc["historial_conversacion"][0]["contenido"])
        self.assertTrue(doc["historial_conversacion"][1]["contenido"].endswith("& FIN HTML"))
        self.assertGreater(len(doc["historial_conversacion"][1]["contenido"]), 60000)

    def test_busqueda_de_cambios_no_descarga_cuerpos_de_mensajes(self):
        from datetime import datetime, timezone
        calls = []
        def paginas(tabla, query, **kwargs):
            calls.append((tabla, kwargs))
            return iter([])
        with ServiceNowClient() as client, patch.object(client, "_paginas", side_effect=paginas):
            list(client.listar_casos_actualizados(datetime(2026, 10, 10, 5, tzinfo=timezone.utc),
                                                datetime(2026, 10, 10, 6, tzinfo=timezone.utc)))
        self.assertEqual("sys_id,instance", calls[-1][1]["fields"])
        self.assertEqual("sys_id,element_id", calls[-2][1]["fields"])
        self.assertEqual("sys_id,task", calls[-3][1]["fields"])

    def test_timeout_no_se_reintenta_en_lectura_ni_expone_respuesta_privada(self):
        import requests
        with ServiceNowClient() as client:
            retries = client._session.get_adapter("https://").max_retries
            self.assertEqual(0, retries.read)
            self.assertEqual(1, retries.total)
            with patch.object(client._session, "get", side_effect=requests.ReadTimeout("password-test")) as get:
                with self.assertRaisesRegex(ServiceNowError, "tardó demasiado") as error:
                    list(client._paginas("sys_email", "instance=" + "a" * 32))
            get.assert_called_once()
            self.assertEqual((10, 30), get.call_args.kwargs["timeout"])
            self.assertNotIn("password-test", str(error.exception))


class SincronizacionManualTest(unittest.TestCase):
    def test_actualiza_rango_con_progreso_y_no_mueve_cursor_automatico(self):
        conn = MagicMock()
        estados = [MagicMock(), MagicMock()]
        mensajes = []
        with patch.object(app_logic, "exigir_permiso_actor"), \
             patch.object(app_logic, "get_conn", side_effect=[conn, *estados]), \
             patch.object(app_logic, "db_execute", return_value=MagicMock(fetchone=lambda: (True,))) as execute, \
             patch.object(app_logic, "guardar_estado_sincronizacion_manual") as estado, \
             patch.object(app_logic, "guardar_estado_sincronizacion") as cursor, \
             patch.object(app_logic, "ServiceNowClient") as constructor, \
             patch.object(app_logic, "_guardar_pagina_servicenow", return_value=(1, 1)) as guardar:
            remoto = constructor.return_value.__enter__.return_value
            remoto.total_registros = 2
            def paginas(*args):
                informar = constructor.call_args.kwargs["progress_callback"]
                informar(0, "Consultando casos...")
                yield [documento()]
                informar(0, "Consultando correos del historial...")
                yield [documento()]
            remoto.listar_rango.side_effect = paginas
            result = app_logic.sincronizar_servicenow(
                actor_email="admin", fecha_inicio=date(2026, 10, 9), fecha_fin=date(2026, 10, 10),
                progress_callback=lambda valor, texto: mensajes.append((valor, texto)))
            remoto.listar_rango.assert_called_once_with(date(2026, 10, 9), date(2026, 10, 10), "casos")
            remoto.listar_mes.assert_not_called()
            self.assertEqual(2, guardar.call_count)
            self.assertEqual(2, result["actualizados"])
            self.assertEqual(["sincronizando", "correcto"], [c.kwargs["estado"] for c in estado.call_args_list])
            cursor.assert_not_called()
        self.assertEqual([v for v, _ in mensajes], sorted(v for v, _ in mensajes))
        self.assertEqual(1, mensajes[-1][0])
        self.assertIn("Casos guardados: 2", mensajes[-1][1])
        self.assertTrue(any("correos del historial" in m for _, m in mensajes))
        self.assertIn("pg_try_advisory_xact_lock", execute.call_args.args[1])
        conn.close.assert_called_once()
        conn.rollback.assert_called_once()
        conn.commit.assert_not_called()
        for conn_estado in estados:
            conn_estado.commit.assert_called_once()
            conn_estado.close.assert_called_once()

    def test_si_otra_sincronizacion_esta_activa_no_espera_ni_consulta_api(self):
        conn = MagicMock()
        with patch.object(app_logic, "exigir_permiso_actor"), \
             patch.object(app_logic, "get_conn", return_value=conn), \
             patch.object(app_logic, "db_execute", return_value=MagicMock(fetchone=lambda: (False,))) as execute, \
             patch.object(app_logic, "ServiceNowClient") as cliente:
            result = app_logic.sincronizar_servicenow(actor_email="admin", fecha_inicio=date.today(), fecha_fin=date.today())
        self.assertTrue(result["ocupado"])
        cliente.assert_not_called()
        self.assertIn("pg_try_advisory_xact_lock", execute.call_args.args[1])
        conn.close.assert_called_once()

    def test_fallo_despues_de_una_pagina_conserva_datos_y_cursor(self):
        def paginas(*args):
            yield [documento()]
            raise ServiceNowError("ServiceNow tardó demasiado al consultar correos.")
        with patch.object(app_logic, "exigir_permiso_actor"), \
             patch.object(app_logic, "get_conn", return_value=MagicMock()), \
             patch.object(app_logic, "ServiceNowClient") as constructor, \
             patch.object(app_logic, "guardar_estado_sincronizacion_manual") as estado, \
             patch.object(app_logic, "guardar_estado_sincronizacion") as cursor, \
             patch.object(app_logic, "_guardar_pagina_servicenow", return_value=(1, 0)) as guardar:
            constructor.return_value.__enter__.return_value.listar_rango.side_effect = paginas
            with self.assertRaises(ServiceNowError):
                app_logic.sincronizar_servicenow(actor_email="admin", fecha_inicio=date.today(), fecha_fin=date.today())
        guardar.assert_called_once()
        self.assertEqual("error", estado.call_args.kwargs["estado"])
        self.assertEqual(1, estado.call_args.kwargs["registros"])
        cursor.assert_not_called()

    def test_pulso_mantiene_transaccion_activa_sin_confirmar_ni_modificar_datos(self):
        conn = MagicMock()
        with patch.object(app_logic, "monotonic", side_effect=[0, 0, 29, 30, 30, 31, 61, 61]), \
             patch.object(app_logic, "db_execute") as ejecutar:
            pulsar = app_logic._crear_pulso_sincronizacion(conn)
            for _ in range(5):
                pulsar()
        self.assertEqual(2, ejecutar.call_count)
        self.assertTrue(all(c.args == (conn, "SELECT 1") for c in ejecutar.call_args_list))
        conn.commit.assert_not_called()
        conn.rollback.assert_not_called()

    def test_interfaz_sin_numero_sincroniza_hoy_ayer_mes_y_rango(self):
        from streamlit.testing.v1 import AppTest
        app = AppTest.from_string('''
from unittest.mock import patch
import streamlit as st
import app_ui as ui
st.session_state["user"] = "admin"
def sincronizar(**kwargs):
    st.session_state["consulta"] = {k: v for k, v in kwargs.items() if k != "progress_callback"}
    st.session_state["llamadas"] = st.session_state.get("llamadas", 0) + 1
    kwargs["progress_callback"](1, "Sincronizacion finalizada.")
    return {"cargados": 3, "actualizados": 2}
with patch.object(ui, "sincronizar_servicenow", side_effect=sincronizar), \\
     patch.object(ui, "sincronizar_casos_automaticamente", side_effect=AssertionError("Consulta manual desde cursor automatico")):
    ui.vista_sincronizar_servicenow("casos")
''').run()
        hoy = pd.Timestamp.now(tz="America/Bogota").date()
        self.assertFalse(app.exception)
        self.assertEqual(["Hoy", "Ayer y hoy", "Mes", "Rango de fechas"], app.radio[0].options)
        self.assertEqual(0, len(app.text_input))
        app.button(key="servicenow_sync_casos_ejecutar").click().run()
        self.assertFalse(app.exception)
        self.assertEqual(hoy, app.session_state["consulta"]["fecha_inicio"])
        self.assertEqual(hoy, app.session_state["consulta"]["fecha_fin"])
        app.radio[0].set_value("Ayer y hoy").run()
        app.button(key="servicenow_sync_casos_ejecutar").click().run()
        self.assertEqual(hoy - timedelta(days=1), app.session_state["consulta"]["fecha_inicio"])
        app.radio[0].set_value("Mes").run()
        app.number_input(key="anio_consulta_sync_casos").set_value(2024).run()
        app.selectbox(key="mes_consulta_sync_casos").select(2).run()
        app.button(key="servicenow_sync_casos_ejecutar").click().run()
        self.assertEqual(date(2024, 2, 1), app.session_state["consulta"]["fecha_inicio"])
        self.assertEqual(date(2024, 2, 29), app.session_state["consulta"]["fecha_fin"])
        app.radio[0].set_value("Rango de fechas").run()
        app.date_input[0].set_value((date(2026, 8, 31), date(2026, 9, 1))).run()
        app.button(key="servicenow_sync_casos_ejecutar").click().run()
        self.assertEqual(date(2026, 8, 31), app.session_state["consulta"]["fecha_inicio"])
        self.assertEqual(date(2026, 9, 1), app.session_state["consulta"]["fecha_fin"])
        self.assertEqual(4, app.session_state["llamadas"])
        app.date_input[0].set_value((date(2026, 8, 31),)).run()
        self.assertEqual(0, len(app.button))
        self.assertEqual(4, app.session_state["llamadas"])
        self.assertFalse(app.exception)

    def test_interfaz_no_deja_progreso_pendiente_si_esta_ocupada(self):
        from streamlit.testing.v1 import AppTest
        app = AppTest.from_string('''
from unittest.mock import patch
import streamlit as st
import app_ui as ui
def crear(*args):
    def informar(valor, texto):
        st.session_state["progreso"] = (valor, texto)
    return informar
with patch.object(ui, "sincronizar_servicenow", return_value={"ocupado": True, "cargados": 0}), \\
     patch.object(ui, "crear_barra_progreso_carga", side_effect=crear), \\
     patch.object(ui, "limpiar_cache_datos", side_effect=AssertionError("No se modificaron datos")):
    ui.vista_sincronizar_servicenow("casos")
''').run()
        app.button(key="servicenow_sync_casos_ejecutar").click().run()
        self.assertFalse(app.exception)
        self.assertEqual(1, app.session_state["progreso"][0])
        self.assertIn("otra sincronización", app.session_state["progreso"][1])
        self.assertIn("Espera a que termine", app.info[0].value)

    def test_navegar_no_repite_preparacion_de_tablas_e_indices(self):
        from streamlit.testing.v1 import AppTest
        app_ui.inicializar_base_datos_cache.clear()
        self.addCleanup(app_ui.inicializar_base_datos_cache.clear)
        app = AppTest.from_string('''
from unittest.mock import patch
import streamlit as st
import app_ui as ui
st.button("Volver a cargar")
def preparar():
    st.session_state["preparaciones"] = st.session_state.get("preparaciones", 0) + 1
with patch.object(ui, "aplicar_tema_visual"), \\
     patch.object(ui, "init_db", side_effect=preparar), \\
     patch.object(ui, "iniciar_sincronizacion_casos"), \\
     patch.object(ui, "hay_usuarios", return_value=True), \\
     patch.object(ui, "login", return_value=False):
    ui.run_app()
''').run()
        self.assertFalse(app.exception)
        self.assertEqual(1, app.session_state["preparaciones"])
        app.button[0].click().run()
        self.assertFalse(app.exception)
        self.assertEqual(1, app.session_state["preparaciones"])


if __name__ == "__main__":
    unittest.main()
