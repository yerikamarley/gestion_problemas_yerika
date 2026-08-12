import unittest
from unittest.mock import Mock, patch

import app_logic
from core.permissions import ROLES_PERMITIDOS, ROLE_VIEWER_LEGACY


def usuario(email="identidad_prueba", role="admin", active=True):
    return {"email": email, "role": role, "active": active}


class AuthenticatedSessionTest(unittest.TestCase):
    def test_usuario_admin_existente_y_activo(self):
        session = {"user": "identidad_prueba", "role": "viewer"}
        resultado = app_logic.refrescar_usuario_autenticado(session, lambda _: usuario())
        self.assertEqual("admin", resultado["role"])
        self.assertEqual("admin", session["role"])

    def test_usuario_viewer_heredado_continua_valido(self):
        session = {"user": "identidad_prueba"}
        resultado = app_logic.refrescar_usuario_autenticado(
            session, lambda _: usuario(role=ROLE_VIEWER_LEGACY)
        )
        self.assertEqual(ROLE_VIEWER_LEGACY, resultado["role"])

    def test_reconoce_cada_rol_definitivo(self):
        for role in ROLES_PERMITIDOS:
            with self.subTest(role=role):
                session = {"user": "identidad_prueba"}
                resultado = app_logic.refrescar_usuario_autenticado(
                    session, lambda _, role=role: usuario(role=role)
                )
                self.assertEqual(role, resultado["role"])

    def test_usuario_inexistente_revoca_sesion(self):
        session = {"user": "identidad_prueba", "role": "admin", "filtro": "anterior"}
        self.assertIsNone(app_logic.refrescar_usuario_autenticado(session, lambda _: None))
        self.assertEqual({}, session)

    def test_usuario_inactivo_revoca_sesion(self):
        session = {"user": "identidad_prueba", "role": "admin"}
        resultado = app_logic.refrescar_usuario_autenticado(
            session, lambda _: usuario(active=False)
        )
        self.assertIsNone(resultado)
        self.assertEqual({}, session)

    def test_rol_invalido_revoca_sesion_sin_permiso_por_defecto(self):
        session = {"user": "identidad_prueba", "role": "admin"}
        resultado = app_logic.refrescar_usuario_autenticado(
            session, lambda _: usuario(role="rol_desconocido")
        )
        self.assertIsNone(resultado)
        self.assertEqual({}, session)

    def test_cambio_de_rol_actualiza_la_sesion(self):
        session = {"user": "identidad_prueba", "role": "soporte"}
        app_logic.refrescar_usuario_autenticado(
            session, lambda _: usuario(role="experiencia")
        )
        self.assertEqual("experiencia", session["role"])

    def test_desactivacion_durante_sesion_activa(self):
        session = {"user": "identidad_prueba", "role": "soporte"}
        consulta = Mock(return_value=usuario(role="soporte", active=False))
        self.assertIsNone(app_logic.refrescar_usuario_autenticado(session, consulta))
        self.assertEqual({}, session)

    def test_eliminacion_durante_sesion_activa(self):
        session = {"user": "identidad_prueba", "role": "soporte"}
        consulta = Mock(return_value=None)
        self.assertIsNone(app_logic.refrescar_usuario_autenticado(session, consulta))
        self.assertEqual({}, session)

    def test_limpieza_segura_elimina_todo_el_estado_anterior(self):
        session = {"user": "identidad_prueba", "role": "admin", "dato": object()}
        app_logic.cerrar_sesion_invalida(session)
        self.assertEqual({}, session)

    def test_ignora_un_rol_enviado_por_la_interfaz(self):
        session = {
            "user": "identidad_prueba",
            "role": "admin",
            "widget_role": "admin",
            "query_role": "admin",
        }
        app_logic.refrescar_usuario_autenticado(
            session, lambda _: usuario(role="gerencias")
        )
        self.assertEqual("gerencias", session["role"])

    def test_obtener_usuario_actual_no_expone_hash(self):
        registro = {
            "email": "identidad_prueba",
            "role": "admin",
            "active": True,
            "password_hash": "valor-simulado",
            "last_login": "fecha-simulada",
        }
        with patch.object(app_logic, "usuario_por_email", return_value=registro):
            resultado = app_logic.obtener_usuario_actual("identidad_prueba")
        self.assertEqual(
            {"email": "identidad_prueba", "role": "admin", "active": True},
            resultado,
        )
        self.assertNotIn("password_hash", resultado)

    def test_flujo_de_admin_inicial_permanece_configurable_y_generico(self):
        with patch.object(app_logic, "usuario_por_email", return_value=usuario()):
            resultado = app_logic.obtener_usuario_actual("identidad_prueba")
        self.assertEqual("admin", app_logic.obtener_rol_actual(resultado))
        self.assertTrue(hasattr(app_logic, "ADMIN_EMAIL"))
        self.assertTrue(hasattr(app_logic, "INITIAL_ADMIN_PASSWORD"))


if __name__ == "__main__":
    unittest.main()
