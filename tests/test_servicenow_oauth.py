import unittest
from unittest.mock import MagicMock, patch

import requests

import app_logic
from core import servicenow_config
from services import servicenow
from services.servicenow import ServiceNowClient, ServiceNowError, ServiceNowAuthenticationError
from services.servicenow_worker import CaseSyncWorker
from tests.test_servicenow import response


class OAuthClientTest(unittest.TestCase):
    def setUp(self):
        self.settings = {
            "SERVICENOW_AUTH_TYPE": "client_credentials", "SERVICENOW_INSTANCE": "test.service-now.com",
            "SERVICENOW_CLIENT_ID": "client-test", "SERVICENOW_CLIENT_SECRET": "secret-test",
            "SERVICENOW_USER": "legacy-test", "SERVICENOW_PASSWORD": "legacy-password-test",
        }
        servicenow._tokens_oauth.clear()
        self.addCleanup(servicenow._tokens_oauth.clear)
        config = patch("services.servicenow.config_value", side_effect=lambda k, d="": self.settings.get(k, d))
        config.start()
        self.addCleanup(config.stop)

    def token(self, value="token-test", expires=100, status=200):
        reply = response([], status=status)
        reply.json.return_value = {"access_token": value, "expires_in": expires, "token_type": "Bearer"}
        return reply

    def test_un_token_compartido_entre_clientes_y_sin_password_personal(self):
        with patch("requests.sessions.Session.post", return_value=self.token()) as post, \
             patch("requests.sessions.Session.get", return_value=response([], 0)) as get:
            for _ in range(2):
                with ServiceNowClient() as client:
                    self.assertIsNone(client._session.auth)
                    client.validar_acceso()
        self.assertEqual(1, post.call_count)
        self.assertEqual(8, get.call_count)
        self.assertEqual({"grant_type": "client_credentials", "client_id": "client-test", "client_secret": "secret-test"},
                         post.call_args.kwargs["data"])
        self.assertFalse(post.call_args.kwargs["allow_redirects"])
        for call in get.call_args_list:
            self.assertEqual("Bearer token-test", call.kwargs["headers"]["Authorization"])
            self.assertNotIn("secret-test", str(call))

    def test_renueva_antes_de_expirar_sin_refresh_token(self):
        with ServiceNowClient() as client, patch.object(servicenow, "monotonic", return_value=100) as clock, \
             patch.object(client._session, "post", side_effect=[self.token("token-one"), self.token("token-two")]) as post:
            self.assertEqual("token-one", client._obtener_token())
            clock.return_value = 189
            self.assertEqual("token-one", client._obtener_token())
            clock.return_value = 190
            self.assertEqual("token-two", client._obtener_token())
        self.assertEqual(2, post.call_count)
        self.assertTrue(all(c.kwargs["data"]["grant_type"] == "client_credentials" for c in post.call_args_list))

    def test_token_revocado_se_renueva_y_reintenta_solo_una_vez(self):
        for acepta in (True, False):
            with self.subTest(acepta=acepta):
                servicenow._tokens_oauth.clear()
                with ServiceNowClient() as client, \
                     patch.object(client._session, "post", side_effect=[self.token("old"), self.token("new")]) as post, \
                     patch.object(client._session, "get", side_effect=[response([], status=401),
                                      response([], 0, status=200 if acepta else 401)]) as get:
                    if acepta:
                        self.assertEqual([], list(client._paginas("task_sla", "active=true")))
                    else:
                        with self.assertRaises(ServiceNowAuthenticationError):
                            list(client._paginas("task_sla", "active=true"))
                self.assertEqual(2, post.call_count)
                self.assertEqual(2, get.call_count)
                self.assertEqual("Bearer new", get.call_args.kwargs["headers"]["Authorization"])

    def test_permiso_403_no_renueva_token_ni_cambia_a_basic(self):
        with ServiceNowClient() as client, patch.object(client._session, "post", return_value=self.token()) as post, \
             patch.object(client._session, "get", return_value=response([], status=403)) as get:
            with self.assertRaisesRegex(ServiceNowError, "permisos de lectura"):
                client.validar_acceso()
        self.assertEqual(1, post.call_count)
        self.assertEqual(1, get.call_count)

    def test_errores_no_exponen_secretos_ni_respuesta_remota(self):
        rejected = self.token(status=401)
        rejected.json.return_value = {"error": "secret-test", "error_description": "token-test"}
        for result in (rejected, requests.ConnectionError("secret-test token-test https://test.service-now.com")):
            kwargs = {"side_effect": result} if isinstance(result, Exception) else {"return_value": result}
            with self.subTest(result=result), ServiceNowClient() as client, \
                 patch.object(client._session, "post", **kwargs), self.assertRaises(ServiceNowError) as error:
                client.validar_acceso()
            self.assertFalse(any(value in str(error.exception) for value in ("secret-test", "token-test", "test.service-now.com")))
        self.assertFalse(servicenow._tokens_oauth)

    def test_respuesta_invalida_no_se_guarda_en_cache(self):
        for token, expires in (("bad\r\nheader", 10), ("token", "nan"), ("token", -1), ("", 100)):
            with self.subTest(token=token), ServiceNowClient() as client, \
                 patch.object(client._session, "post", return_value=self.token(token, expires)), \
                 self.assertRaises(ServiceNowError):
                client._obtener_token()
            self.assertFalse(servicenow._tokens_oauth)

    def test_rechazo_identifica_causa_sin_exponer_descripcion_remota(self):
        for codigo, descripcion, esperado in (
            ("invalid_client", "secret-test token-test", "ID o el secreto de cliente"),
            ("unsupported_grant_type", "secret-test", "glide.oauth.inbound.client.credential.grant_type.enabled"),
            ("unauthorized_client", "secret-test", "Credenciales del cliente"),
            ("invalid_scope", "token-test", "SERVICENOW_OAUTH_SCOPE"),
            ("server_error", "access_denied", "usuario de aplicación OAuth de portalgestion"),
            ("access_denied", "secret-test", "usuario de aplicación OAuth de portalgestion"),
            ("invalid_request", "secret-test", "parámetro rechazado"),
            ("invalid_grant", "secret-test", "concesión OAuth"),
            ("secret-test", "token-test", "registro OAuth"),
            ({"message": "secret-test"}, "token-test", "registro OAuth"),
        ):
            with self.subTest(codigo=codigo), ServiceNowClient() as client:
                rejected = self.token(status=401)
                rejected.json.return_value = {"error": codigo, "error_description": descripcion}
                with patch.object(client._session, "post", return_value=rejected), \
                     self.assertRaises(ServiceNowAuthenticationError) as error:
                    client.validar_acceso()
                message = str(error.exception)
                self.assertIn(esperado, message)
                self.assertIn("HTTP 401", message)
                self.assertNotIn("secret-test", message)
                self.assertNotIn("token-test", message)
                self.assertFalse(servicenow._tokens_oauth)

    def test_rechazo_no_json_conserva_diagnostico_http_seguro(self):
        for payload in (ValueError("secret-test"), ["secret-test"], None):
            with self.subTest(payload=payload), ServiceNowClient() as client:
                rejected = self.token(status=403)
                if isinstance(payload, Exception):
                    rejected.json.side_effect = payload
                else:
                    rejected.json.return_value = payload
                with patch.object(client._session, "post", return_value=rejected), \
                     self.assertRaises(ServiceNowAuthenticationError) as error:
                    client.validar_acceso()
                self.assertIn("HTTP 403", str(error.exception))
                self.assertNotIn("secret-test", str(error.exception))

    def test_secreto_rotado_no_reutiliza_el_token_anterior(self):
        with patch("requests.sessions.Session.post", side_effect=[self.token("first"), self.token("second")]) as post:
            with ServiceNowClient() as client:
                self.assertEqual("first", client._obtener_token())
            self.settings["SERVICENOW_CLIENT_SECRET"] = "secret-rotated"
            with ServiceNowClient() as client:
                self.assertEqual("second", client._obtener_token())
        self.assertEqual(2, post.call_count)

    def test_configuracion_oauth_no_requiere_usuario_password(self):
        self.settings.pop("SERVICENOW_USER")
        self.settings.pop("SERVICENOW_PASSWORD")
        with patch.object(servicenow_config, "config_value", side_effect=lambda k, d="": self.settings.get(k, d)):
            self.assertTrue(servicenow_config.conexion_configurada())
            self.settings.pop("SERVICENOW_CLIENT_SECRET")
            self.assertFalse(servicenow_config.conexion_configurada())
            with self.assertRaises(ServiceNowError):
                ServiceNowClient()


class OAuthActivationTest(unittest.TestCase):
    def test_solo_admin_puede_validar_y_activar(self):
        with patch.object(app_logic, "exigir_permiso_actor", side_effect=app_logic.AutorizacionError), \
             patch.object(app_logic, "ServiceNowClient") as client, \
             patch.object(app_logic, "activar_sincronizacion_casos") as activate:
            with self.assertRaises(app_logic.AutorizacionError):
                app_logic.conectar_servicenow_oauth(actor_email="viewer")
        client.assert_not_called()
        activate.assert_not_called()

    def test_valida_y_revalida_admin_antes_de_activar_sin_escribir_secretos(self):
        for revocado in (False, True):
            with self.subTest(revocado=revocado):
                conn = MagicMock()
                with patch.object(app_logic, "exigir_permiso_actor"), \
                     patch.object(app_logic, "config_value", return_value="client_credentials"), \
                     patch.object(app_logic, "ServiceNowClient") as client, \
                     patch.object(app_logic, "get_conn", return_value=conn), \
                     patch.object(app_logic, "_bloquear_y_validar_admin", side_effect=app_logic.AutorizacionError if revocado else None), \
                     patch.object(app_logic, "activar_sincronizacion_casos") as activate, \
                     patch.object(app_logic, "guardar_conexion_servidor") as save:
                    if revocado:
                        with self.assertRaises(app_logic.AutorizacionError):
                            app_logic.conectar_servicenow_oauth(actor_email="ADMIN@EXAMPLE.TEST")
                        activate.assert_not_called()
                        conn.rollback.assert_called_once()
                    else:
                        self.assertEqual({"conectado": True, "automatica": True},
                                         app_logic.conectar_servicenow_oauth(actor_email="ADMIN@EXAMPLE.TEST"))
                        activate.assert_called_once_with(conn, "admin@example.test")
                        conn.commit.assert_called_once()
                    save.assert_not_called()
                    client.return_value.__enter__.return_value.validar_acceso.assert_called_once()
                    conn.close.assert_called_once()

    def test_fuente_inaccesible_no_activa_actualizacion(self):
        with patch.object(app_logic, "exigir_permiso_actor"), \
             patch.object(app_logic, "config_value", return_value="client_credentials"), \
             patch.object(app_logic, "ServiceNowClient") as client, \
             patch.object(app_logic, "get_conn") as conn:
            client.return_value.__enter__.return_value.validar_acceso.side_effect = ServiceNowError("HTTP 403")
            with self.assertRaises(ServiceNowError):
                app_logic.conectar_servicenow_oauth(actor_email="admin")
        conn.assert_not_called()

    def test_trabajador_lee_activacion_durable_sin_credenciales_en_env(self):
        config = MagicMock(return_value=(True, "admin@example.test"))
        callback = MagicMock()
        with patch("services.servicenow_worker.Thread"), patch("services.servicenow_worker.config_value", return_value=""):
            for _ in range(2):
                worker = CaseSyncWorker(callback, config)
                worker.ejecutar_ciclo()
                worker.detener()
        self.assertEqual(2, callback.call_count)
        callback.assert_called_with(actor_email="admin@example.test")


if __name__ == "__main__":
    unittest.main()
