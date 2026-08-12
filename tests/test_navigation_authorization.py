import unittest

import app_ui
from core.permissions import (
    NOMBRES_ROLES,
    ROLE_ADMIN,
    ROLE_AREA_IA,
    ROLE_EXPERIENCIA,
    ROLE_GERENCIAS,
    ROLE_SOPORTE,
    ROLE_VIEWER_LEGACY,
    VIEW_ADMINISTRAR_USUARIOS,
)


class NavigationAuthorizationTest(unittest.TestCase):
    def test_catalogo_unico_contiene_todas_las_vistas(self):
        self.assertEqual(18, len(app_ui.VIEW_CATALOG))
        self.assertEqual(18, len({item[0] for item in app_ui.VIEW_CATALOG}))
        self.assertFalse(hasattr(app_ui, "ADMIN_MENU_OPTIONS"))
        self.assertFalse(hasattr(app_ui, "VIEWER_MENU_OPTIONS"))

    def test_cantidad_exacta_por_rol(self):
        esperadas = {
            ROLE_ADMIN: 18,
            ROLE_SOPORTE: 7,
            ROLE_EXPERIENCIA: 7,
            ROLE_GERENCIAS: 4,
            ROLE_AREA_IA: 7,
            ROLE_VIEWER_LEGACY: 13,
        }
        self.assertEqual(esperadas, {rol: len(app_ui.catalogo_permitido(rol)) for rol in esperadas})

    def test_area_ia_tiene_nombre_visible(self):
        self.assertEqual("Área de IA", NOMBRES_ROLES[ROLE_AREA_IA])

    def test_categorias_vacias_no_aparecen(self):
        categorias = app_ui.categorias_permitidas(ROLE_GERENCIAS)
        self.assertNotIn(app_ui.CATEGORY_ADMIN, categorias)

    def test_id_inexistente_o_manipulado_se_reemplaza(self):
        vista, rechazada = app_ui.resolver_vista_permitida(ROLE_SOPORTE, "id_inexistente")
        self.assertTrue(rechazada)
        self.assertIsNotNone(vista)

    def test_cambio_de_rol_reemplaza_vista_anterior(self):
        vista, rechazada = app_ui.resolver_vista_permitida(
            ROLE_GERENCIAS, VIEW_ADMINISTRAR_USUARIOS
        )
        self.assertTrue(rechazada)
        self.assertNotEqual(VIEW_ADMINISTRAR_USUARIOS, vista)

    def test_usuario_sin_vistas_no_recibe_seleccion(self):
        self.assertEqual((None, True), app_ui.resolver_vista_permitida("rol_invalido", "vista"))


if __name__ == "__main__":
    unittest.main()
