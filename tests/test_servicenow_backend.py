import base64
import json
import unittest
from unittest.mock import patch

import backend
from app_logic import AutorizacionError
from services.servicenow import ServiceNowError


class ServiceNowBackendTest(unittest.IsolatedAsyncioTestCase):
    async def request(self, path, method="GET", body=None, credentials="admin:local-password"):
        encoded = base64.b64encode(credentials.encode()).decode() if credentials else ""
        headers = [(b"authorization", f"Basic {encoded}".encode())] if credentials else []
        messages = []
        scope = {"type": "http", "asgi": {"version": "3.0"}, "http_version": "1.1",
                 "method": method, "scheme": "http", "path": path, "raw_path": path.encode(),
                 "query_string": b"", "root_path": "", "headers": headers,
                 "server": ("testserver", 80), "client": ("127.0.0.1", 1000)}

        async def receive():
            return {"type": "http.request", "body": json.dumps(body).encode(), "more_body": False}

        async def send(message):
            messages.append(message)

        await backend.app(scope, receive, send)
        status = next(m["status"] for m in messages if m["type"] == "http.response.start")
        payload = b"".join(m.get("body", b"") for m in messages if m["type"] == "http.response.body")
        return status, json.loads(payload)

    async def test_sin_identidad_no_consulta_ni_sincroniza(self):
        with patch.object(backend, "sincronizar_servicenow") as sync, \
             patch.object(backend, "consultar_traza_servicenow") as read:
            status, _ = await self.request("/api/servicenow/casos/CS0215905", credentials=None)
        self.assertEqual(401, status)
        sync.assert_not_called()
        read.assert_not_called()

    async def test_usuario_local_sin_permiso_recibe_403(self):
        with patch.object(backend, "autenticar_usuario", return_value={"email": "viewer"}), \
             patch.object(backend, "sincronizar_servicenow", side_effect=AutorizacionError):
            status, _ = await self.request("/api/servicenow/casos/sync", "POST", {"mes": "2026-10"})
        self.assertEqual(403, status)

    async def test_sync_autenticado_no_acepta_usuario_ni_filtros_remotos(self):
        with patch.object(backend, "autenticar_usuario", return_value={"email": "admin"}), \
             patch.object(backend, "sincronizar_servicenow", return_value={"cargados": 3}) as sync:
            status, body = await self.request("/api/servicenow/incidentes/sync", "POST", {"mes": "2026-10"})
            self.assertEqual(200, status)
            self.assertEqual(3, body["cargados"])
            sync.assert_called_once_with(actor_email="admin", tipo="incidentes", mes="2026-10")
            status, _ = await self.request("/api/servicenow/casos/sync", "POST", {"mes": "2026-10", "url": "http://localhost"})
            self.assertEqual(400, status)
            self.assertEqual(1, sync.call_count)

    async def test_get_local_y_registro_ausente(self):
        with patch.object(backend, "autenticar_usuario", return_value={"email": "admin"}), \
             patch.object(backend, "consultar_traza_servicenow", return_value=None), \
             patch.object(backend, "sincronizar_servicenow") as sync:
            status, _ = await self.request("/api/servicenow/casos/CS0215905")
        self.assertEqual(404, status)
        sync.assert_not_called()

    async def test_errores_de_conexion_no_exponen_datos_privados(self):
        for error, expected in [(RuntimeError("password@database-host"), 503), (ServiceNowError("HTTP 403"), 502)]:
            with patch.object(backend, "autenticar_usuario", return_value={"email": "admin"}), \
                 patch.object(backend, "sincronizar_servicenow", side_effect=error):
                status, body = await self.request("/api/servicenow/casos/CS0215905/sync", "POST")
            self.assertEqual(expected, status)
            self.assertNotIn("password@database-host", json.dumps(body))


if __name__ == "__main__":
    unittest.main()
