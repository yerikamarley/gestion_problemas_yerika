import unittest
from unittest.mock import Mock, patch

import app_logic


class FakeCursor:
    def __init__(self, row=None, rows=None):
        self.row = row
        self.rows = rows or []

    def fetchone(self):
        return self.row

    def fetchall(self):
        return self.rows


class UserManagementAuthorizationTest(unittest.TestCase):
    def test_rol_invalido_y_viewer_no_son_asignables(self):
        for role in ("viewer", "rol_invalido"):
            with self.subTest(role=role), self.assertRaises(ValueError):
                app_logic.guardar_usuario(
                    "actor_prueba", "Nombre de prueba", "destino_prueba", None, role
                )

    def test_ultimo_admin_se_valida_bajo_bloqueo_y_transaccion(self):
        connection = Mock()
        sql_seen = []

        def execute(_conn, sql, params=()):
            normalized = " ".join(sql.split())
            sql_seen.append(normalized)
            if "WHERE email = ? FOR UPDATE" in normalized:
                return FakeCursor(("admin", True))
            if "SELECT 1 FROM app_users WHERE email" in normalized:
                return FakeCursor((1,))
            if "COUNT(*) FROM app_users WHERE role = 'admin'" in normalized:
                return FakeCursor((1,))
            return FakeCursor()

        with patch.object(app_logic, "get_conn", return_value=connection), patch.object(
            app_logic, "db_execute", side_effect=execute
        ):
            with self.assertRaisesRegex(ValueError, "último Admin"):
                app_logic.guardar_usuario(
                    "actor_prueba",
                    "Nombre de prueba",
                    "destino_prueba",
                    None,
                    "soporte",
                    active=True,
                )

        self.assertTrue(any("LOCK TABLE app_users" in sql for sql in sql_seen))
        connection.rollback.assert_called_once()
        connection.commit.assert_not_called()

    def test_actor_no_admin_no_puede_administrar(self):
        connection = Mock()

        def execute(_conn, sql, params=()):
            if "FOR UPDATE" in sql:
                return FakeCursor(("soporte", True))
            return FakeCursor()

        with patch.object(app_logic, "get_conn", return_value=connection), patch.object(
            app_logic, "db_execute", side_effect=execute
        ):
            with self.assertRaises(app_logic.AutorizacionError):
                app_logic.guardar_usuario(
                    "actor_prueba", "Nombre de prueba", "destino_prueba", None, "soporte"
                )
        connection.rollback.assert_called_once()


if __name__ == "__main__":
    unittest.main()
