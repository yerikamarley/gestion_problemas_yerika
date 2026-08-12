import unittest
from unittest.mock import MagicMock, patch

import pandas as pd

import app_logic
from core.permissions import (
    ACTION_MANAGE_USERS,
    ACTION_PURGE_INCIDENTS,
    ACTION_WRITE_CASES,
    ROLE_ADMIN,
    VIEW_CASOS,
)


def identity(role=ROLE_ADMIN, active=True):
    return {"email": "identidad_prueba", "role": role, "active": active}


class OperationAuthorizationTest(unittest.TestCase):
    def test_admin_recibe_capacidad(self):
        actor = app_logic.exigir_permiso_actor(
            "identidad_prueba", ACTION_MANAGE_USERS, lambda _: identity()
        )
        self.assertEqual(ROLE_ADMIN, actor["role"])

    def test_no_admin_no_puede_cargar_excel(self):
        with self.assertRaises(app_logic.AutorizacionError):
            app_logic.exigir_permiso_actor(
                "identidad_prueba", ACTION_WRITE_CASES, lambda _: identity("soporte")
            )

    def test_no_admin_no_puede_administrar_ni_limpiar(self):
        for capacidad in (ACTION_MANAGE_USERS, ACTION_PURGE_INCIDENTS):
            with self.subTest(capacidad=capacidad), self.assertRaises(app_logic.AutorizacionError):
                app_logic.exigir_permiso_actor(
                    "identidad_prueba", capacidad, lambda _: identity("gerencias")
                )

    def test_consulta_restringida_es_rechazada(self):
        with self.assertRaises(app_logic.AutorizacionError):
            app_logic.exigir_permiso_actor(
                "identidad_prueba", VIEW_CASOS, lambda _: identity("gerencias")
            )

    def test_rol_no_proviene_del_formulario(self):
        with self.assertRaises(app_logic.AutorizacionError):
            app_logic.exigir_permiso_actor(
                "identidad_prueba", ACTION_MANAGE_USERS, lambda _: identity("soporte")
            )

    def test_usuario_eliminado_inactivo_o_invalido(self):
        invalidos = (None, identity(active=False), identity("rol_invalido"))
        for registro in invalidos:
            with self.subTest(registro=registro), self.assertRaises(app_logic.AutorizacionError):
                app_logic.exigir_permiso_actor(
                    "identidad_prueba", ACTION_MANAGE_USERS, lambda _, r=registro: r
                )

    def test_fallo_de_conexion_no_concede_permiso(self):
        def falla(_):
            raise RuntimeError("fallo simulado")

        with self.assertRaises(RuntimeError):
            app_logic.exigir_permiso_actor("identidad_prueba", ACTION_MANAGE_USERS, falla)

    def test_mutacion_rechaza_antes_de_procesar_datos(self):
        dataframe = MagicMock(spec=pd.DataFrame)
        with patch.object(app_logic, "exigir_permiso_actor", side_effect=app_logic.AutorizacionError):
            with self.assertRaises(app_logic.AutorizacionError):
                app_logic.guardar_casos(dataframe, actor_email="identidad_prueba")
        dataframe.__len__.assert_not_called()


if __name__ == "__main__":
    unittest.main()
