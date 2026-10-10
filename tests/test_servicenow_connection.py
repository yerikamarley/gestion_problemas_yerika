import os
import tempfile
import unittest
from datetime import datetime, timedelta, timezone
from pathlib import Path
from unittest.mock import MagicMock, patch

from dotenv import dotenv_values
from streamlit.testing.v1 import AppTest

import app_logic
import app_ui
from core import servicenow_config
from services import servicenow
from services.servicenow import ServiceNowClient, ServiceNowError, ServiceNowAuthenticationError, ServiceNowMFARequiredError, mapear_documento
from services.servicenow_worker import CaseSyncWorker
from tests.test_servicenow import caso, documento, field, response


class ConnectionPersistenceTest(unittest.TestCase):
    def test_guarda_secretos_solo_en_env_y_conserva_otras_variables(self):
        with tempfile.TemporaryDirectory() as directory, \
             patch.object(servicenow_config, "PROJECT_ROOT", Path(directory)), \
             patch.dict(os.environ, {}, clear=True):
            path = Path(directory) / ".env"
            path.write_text("OTHER_SETTING=conservar\n", encoding="utf-8")
            values = {"SERVICENOW_USER": "prueba", "SERVICENOW_PASSWORD": "test-'password",
                      "SERVICENOW_SYNC_ACTOR": "admin@example.test", "SERVICENOW_SYNC_ENABLED": "true",
                      "SERVICENOW_INSTANCE": "test.service-now.com"}
            servicenow_config.guardar_conexion_servidor(values)
            stored = dotenv_values(path)
            self.assertEqual("conservar", stored["OTHER_SETTING"])
            self.assertEqual(values["SERVICENOW_PASSWORD"], stored["SERVICENOW_PASSWORD"])
            self.assertEqual(values["SERVICENOW_PASSWORD"], os.environ["SERVICENOW_PASSWORD"])
            self.assertEqual(values["SERVICENOW_INSTANCE"], stored["SERVICENOW_INSTANCE"])
            self.assertFalse(list(Path(directory).glob(".servicenow-*")))

    def test_fallo_escritura_no_reemplaza_conexion_anterior(self):
        with tempfile.TemporaryDirectory() as directory, \
             patch.object(servicenow_config, "PROJECT_ROOT", Path(directory)), \
             patch.dict(os.environ, {"SERVICENOW_USER": "anterior"}), \
             patch.object(servicenow_config, "set_key", side_effect=OSError("fallo simulado")):
            path = Path(directory) / ".env"
            path.write_text("SERVICENOW_USER=anterior\n", encoding="utf-8")
            with self.assertRaises(OSError):
                servicenow_config.guardar_conexion_servidor({
                    "SERVICENOW_USER": "nuevo", "SERVICENOW_PASSWORD": "test",
                    "SERVICENOW_SYNC_ACTOR": "admin", "SERVICENOW_SYNC_ENABLED": "true",
                })
            self.assertEqual("SERVICENOW_USER=anterior\n", path.read_text(encoding="utf-8"))
            self.assertEqual("anterior", os.environ["SERVICENOW_USER"])
            self.assertFalse(list(Path(directory).glob(".servicenow-*")))


class ConnectionAuthorizationTest(unittest.TestCase):
    def test_no_admin_no_consulta_ni_guarda_credenciales(self):
        with patch.object(app_logic, "exigir_permiso_actor", side_effect=app_logic.AutorizacionError), \
             patch.object(app_logic, "ServiceNowClient") as client, \
             patch.object(app_logic, "guardar_conexion_servidor") as save:
            with self.assertRaises(app_logic.AutorizacionError):
                app_logic.conectar_servicenow("prueba", "test", actor_email="viewer")
        client.assert_not_called()
        save.assert_not_called()

    def test_acceso_rechazado_no_sobrescribe_la_conexion(self):
        with patch.object(app_logic, "exigir_permiso_actor"), \
             patch.object(app_logic, "ServiceNowClient") as client, \
             patch.object(app_logic, "guardar_conexion_servidor") as save:
            client.return_value.__enter__.return_value.validar_acceso.side_effect = ServiceNowError("HTTP 403")
            with self.assertRaises(ServiceNowError):
                app_logic.conectar_servicenow("prueba", "test", actor_email="admin")
        save.assert_not_called()

    def test_admin_revocado_entre_validacion_y_guardado(self):
        conn = MagicMock()
        with patch.object(app_logic, "exigir_permiso_actor"), \
             patch.object(app_logic, "ServiceNowClient"), \
             patch.object(app_logic, "get_conn", return_value=conn), \
             patch.object(app_logic, "_bloquear_y_validar_admin", side_effect=app_logic.AutorizacionError), \
             patch.object(app_logic, "guardar_conexion_servidor") as save:
            with self.assertRaises(app_logic.AutorizacionError):
                app_logic.conectar_servicenow("prueba", "test", actor_email="admin")
        save.assert_not_called()
        conn.rollback.assert_called_once()
        conn.close.assert_called_once()

    def test_conexion_exitosa_no_retorna_secretos(self):
        conn = MagicMock()
        with patch.object(app_logic, "exigir_permiso_actor"), \
             patch.object(app_logic, "ServiceNowClient") as client, \
             patch.object(app_logic, "get_conn", return_value=conn), \
             patch.object(app_logic, "_bloquear_y_validar_admin"), \
             patch.object(app_logic, "guardar_conexion_servidor") as save:
            result = app_logic.conectar_servicenow("prueba", "test-secret", actor_email="ADMIN@EXAMPLE.TEST", codigo_mfa="123456")
        self.assertEqual({"conectado": True, "automatica": True}, result)
        client.assert_called_once_with(user="prueba", password="test-secret", mfa_code="123456")
        self.assertEqual("admin@example.test", save.call_args.args[0]["SERVICENOW_SYNC_ACTOR"])
        self.assertNotIn("123456", str(save.call_args))
        client.return_value.__enter__.return_value.recordar_sesion.assert_called_once()
        conn.close.assert_called_once()

    def test_instancia_indicada_se_normaliza_y_guarda_tras_validar(self):
        with patch.object(app_logic, "exigir_permiso_actor"), \
             patch.object(app_logic, "ServiceNowClient") as client, \
             patch.object(app_logic, "get_conn", return_value=MagicMock()), \
             patch.object(app_logic, "_bloquear_y_validar_admin"), \
             patch.object(app_logic, "guardar_conexion_servidor") as save:
            app_logic.conectar_servicenow("prueba", "test", actor_email="admin",
                                         instancia="https://test.service-now.com/navpage.do")
        client.assert_called_once_with(user="prueba", password="test", mfa_code="", instance="test.service-now.com")
        client.return_value.__enter__.return_value.validar_acceso.assert_called_once()
        self.assertEqual("test.service-now.com", save.call_args.args[0]["SERVICENOW_INSTANCE"])


class MFASessionTest(unittest.TestCase):
    def setUp(self):
        servicenow._sesiones_api.clear()
        self.addCleanup(servicenow._sesiones_api.clear)
        settings = {"SERVICENOW_INSTANCE": "test.service-now.com", "SERVICENOW_USER": "prueba", "SERVICENOW_PASSWORD": "test"}
        patcher = patch("services.servicenow.config_value", side_effect=lambda name, default="": settings.get(name, default))
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_detecta_requisito_mfa_y_no_expone_respuesta(self):
        rejected = response([], status=401)
        rejected.headers = {"X-MFA_TOKEN": "invalid"}
        with ServiceNowClient() as client, patch.object(client._session, "get", return_value=rejected):
            with self.assertRaises(ServiceNowMFARequiredError):
                client.validar_acceso()

    def test_cookies_se_comparten_solo_en_servidor_sin_guardar_codigo(self):
        with ServiceNowClient(mfa_code="123456") as client:
            self.assertEqual("123456", client._session.headers["X-MFA_TOKEN"])
            client._session.cookies.set("JSESSIONID", "sesion-test", domain="test.service-now.com", path="/")
            client.recordar_sesion()
            self.assertNotIn("X-MFA_TOKEN", client._session.headers)
        with ServiceNowClient() as following:
            self.assertEqual("sesion-test", following._session.cookies.get("JSESSIONID"))
            self.assertNotIn("X-MFA_TOKEN", following._session.headers)


class IncrementalSyncTest(unittest.TestCase):
    def test_correo_de_caso_antiguo_se_recoge_y_no_se_duplica_por_fuentes(self):
        desde = datetime(2026, 10, 10, 5, tzinfo=timezone.utc)
        hasta = desde + timedelta(minutes=2)
        settings = {"SERVICENOW_INSTANCE": "test.service-now.com", "SERVICENOW_USER": "prueba", "SERVICENOW_PASSWORD": "test"}
        calls = []
        antiguo = caso()
        antiguo["sys_id"] = field("b" * 32)

        def pages(table, query):
            calls.append((table, query))
            if table == "sn_customerservice_case":
                return iter([[antiguo]]) if "sys_idIN" in query else iter([[caso()]])
            reference = {"task_sla": "task", "sys_journal_field": "element_id", "sys_email": "instance"}[table]
            return iter([[{reference: field("b" * 32)}]])

        with patch("services.servicenow.config_value", side_effect=lambda n, d="": settings.get(n, d)), \
             ServiceNowClient() as client, patch.object(client, "_paginas", side_effect=pages), \
             patch.object(client, "_consolidar_pagina", side_effect=lambda rows: [{"case_id": raw["sys_id"]["value"]} for raw in rows]):
            docs = [d for page in client.listar_casos_actualizados(desde, hasta) for d in page]
        self.assertEqual(["a" * 32, "b" * 32], [d["case_id"] for d in docs])
        self.assertIn("sys_updated_on>=2026-10-10 05:00:00", calls[0][1])
        self.assertTrue(any(t == "sys_email" and "target_table=sn_customerservice_case" in q for t, q in calls))
        self.assertEqual(1, sum("sys_idIN" in q for _, q in calls))

    def test_cursor_avanza_solo_si_todo_se_guarda(self):
        ultima = datetime(2026, 10, 10, 6, tzinfo=timezone.utc)
        now = ultima + timedelta(minutes=2)
        for falla in (False, True):
            with self.subTest(falla=falla):
                control = MagicMock()
                with patch.object(app_logic, "exigir_permiso_actor"), \
                     patch.object(app_logic, "get_conn", return_value=control), \
                     patch.object(app_logic, "db_execute", return_value=MagicMock(fetchone=lambda: (True,))), \
                     patch.object(app_logic, "leer_estado_sincronizacion", return_value={"ultima_correcta": ultima}), \
                     patch.object(app_logic, "guardar_estado_sincronizacion") as state, \
                     patch.object(app_logic, "datetime", wraps=datetime) as clock, \
                     patch.object(app_logic, "ServiceNowClient") as client, \
                     patch.object(app_logic, "_guardar_pagina_servicenow", side_effect=RuntimeError("secret@host") if falla else None, return_value=(1, 1)):
                    clock.now.return_value = now
                    remote = client.return_value.__enter__.return_value
                    remote.listar_casos_actualizados.return_value = iter([[documento()]])
                    if falla:
                        with self.assertRaises(RuntimeError):
                            app_logic.sincronizar_casos_automaticamente(actor_email="admin")
                        self.assertTrue(all(c.kwargs.get("correcta") is None for c in state.call_args_list))
                        self.assertEqual("error", state.call_args.kwargs["estado"])
                        self.assertNotIn("secret@host", state.call_args.kwargs["mensaje"])
                    else:
                        self.assertEqual({"cargados": 1, "actualizados": 1}, app_logic.sincronizar_casos_automaticamente(actor_email="admin"))
                        self.assertEqual(now, state.call_args.kwargs["correcta"])
                    remote.listar_casos_actualizados.assert_called_once_with(ultima - timedelta(minutes=2), now)
                control.close.assert_called_once()

    def test_ciclo_simultaneo_no_hace_llamadas_a_api(self):
        control = MagicMock()
        with patch.object(app_logic, "exigir_permiso_actor"), \
             patch.object(app_logic, "get_conn", return_value=control), \
             patch.object(app_logic, "db_execute", return_value=MagicMock(fetchone=lambda: (False,))), \
             patch.object(app_logic, "ServiceNowClient") as client:
            result = app_logic.sincronizar_casos_automaticamente(actor_email="admin")
        self.assertTrue(result["ocupado"])
        client.assert_not_called()
        control.close.assert_called_once()

    def test_respuesta_sin_notas_no_vacia_columnas_historicas(self):
        row = mapear_documento(documento())
        self.assertNotIn("notas_trabajo", row)
        self.assertNotIn("observaciones_trabajo", row)


class AutomaticWorkerTest(unittest.TestCase):
    def test_autenticacion_rechazada_detiene_intentos_hasta_reconexion_admin(self):
        for error in (ServiceNowMFARequiredError, ServiceNowAuthenticationError):
            with self.subTest(error=error), patch("services.servicenow_worker.Thread"), \
                 patch("services.servicenow_worker.sincronizacion_habilitada", return_value=True), \
                 patch("services.servicenow_worker.config_value", return_value="admin"):
                callback = MagicMock(side_effect=error("renovar"))
                worker = CaseSyncWorker(callback)
                worker.ejecutar_ciclo()
                worker.ejecutar_ciclo()
                self.assertEqual(1, callback.call_count)
                worker.despertar()
                callback.side_effect = None
                worker.ejecutar_ciclo()
                self.assertEqual(2, callback.call_count)
                worker.detener()

    def test_sin_conexion_activa_no_llama_servicenow(self):
        callback = MagicMock()
        with patch("services.servicenow_worker.Thread"), \
             patch("services.servicenow_worker.sincronizacion_habilitada", return_value=False):
            worker = CaseSyncWorker(callback)
            worker.ejecutar_ciclo()
            worker.detener()
        callback.assert_not_called()


class AdminConnectionPresentationTest(unittest.TestCase):
    def app(self, admin=True):
        source = '''
from unittest.mock import patch
import streamlit as st
import app_ui as ui
st.session_state['user'] = 'admin@example.test'
with patch.object(ui, 'exigir_permiso_actor', PERMISSION), patch.object(ui, 'render_estado_actualizacion_casos'):
    ui.vista_cargar_casos()
'''
        source = source.replace("PERMISSION", "return_value={}" if admin else "side_effect=ui.AutorizacionError")
        return AppTest.from_string(source).run(timeout=15)

    def test_no_admin_no_ve_formulario_servicenow(self):
        app = self.app(admin=False)
        self.assertFalse(app.exception)
        self.assertEqual(0, len(app.text_input))
        self.assertIn("Solo un administrador", app.error[0].value)

    def test_admin_activa_oauth_sin_pedir_credenciales_personales(self):
        app = self.app()
        self.assertFalse(app.exception)
        self.assertEqual(0, len(app.get("file_uploader")))
        self.assertFalse(any("Contraseña" in element.label or "Código de verificación" in element.label
                             for element in app.text_input))
        with patch.object(app_ui, "conectar_servicenow_oauth", return_value={"conectado": True}) as connect, \
             patch.object(app_ui, "iniciar_sincronizacion_casos") as worker:
            app.button[0].click().run(timeout=15)
        self.assertFalse(app.exception)
        connect.assert_called_once_with(actor_email="admin@example.test")
        worker.return_value.despertar.assert_called_once()
        self.assertEqual("", app.session_state["servicenow_login_password"])
        self.assertEqual("", app.session_state["servicenow_login_mfa"])
        self.assertTrue(app.success)


if __name__ == "__main__":
    unittest.main()
