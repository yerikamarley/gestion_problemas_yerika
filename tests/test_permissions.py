import unittest

from core.permissions import (
    PERMISOS_POR_ROL,
    ROLE_ADMIN,
    ROLE_AREA_IA,
    ROLE_EXPERIENCIA,
    ROLE_GERENCIAS,
    ROLE_SOPORTE,
    ROLE_VIEWER_LEGACY,
    ROLES_PERMITIDOS,
    VISTAS,
    VIEW_ADMINISTRAR_USUARIOS,
    VIEW_CARGAR_CASOS,
    VIEW_CARGAR_INCIDENTES,
    VIEW_CARGAR_PROBLEMAS,
    VIEW_DASHBOARD_INCIDENTES,
    VIEW_REINCIDENCIAS_PROBLEMAS,
    normalizar_rol,
    obtener_permisos_rol,
    obtener_vistas_permitidas,
    puede_acceder,
    rol_valido,
)


class PermissionsTest(unittest.TestCase):
    def test_define_cinco_roles_permitidos(self):
        self.assertEqual(
            ("admin", "soporte", "experiencia", "gerencias", "area_ia"),
            ROLES_PERMITIDOS,
        )

    def test_define_vistas_unicas(self):
        self.assertEqual(18, len(VISTAS))
        self.assertEqual(len(VISTAS), len(set(VISTAS)))

    def test_normaliza_roles_sin_inventar_un_default(self):
        self.assertEqual("area_ia", normalizar_rol(" Área IA ".replace("Á", "A")))
        self.assertEqual("gerencias", normalizar_rol(" GERENCIAS "))
        self.assertEqual("", normalizar_rol(None))

    def test_valida_roles_definitivos_y_heredado(self):
        for rol in ROLES_PERMITIDOS:
            self.assertTrue(rol_valido(rol))
        self.assertTrue(rol_valido(ROLE_VIEWER_LEGACY))
        self.assertFalse(rol_valido(ROLE_VIEWER_LEGACY, incluir_heredados=False))
        self.assertFalse(rol_valido("desconocido"))

    def test_admin_accede_a_todas_las_vistas(self):
        self.assertEqual(frozenset(VISTAS), obtener_permisos_rol(ROLE_ADMIN))
        self.assertEqual(VISTAS, obtener_vistas_permitidas(ROLE_ADMIN))

    def test_cantidad_de_vistas_por_rol_definitivo(self):
        esperadas = {
            ROLE_ADMIN: 18,
            ROLE_SOPORTE: 7,
            ROLE_EXPERIENCIA: 7,
            ROLE_GERENCIAS: 4,
            ROLE_AREA_IA: 7,
        }
        self.assertEqual(esperadas, {rol: len(PERMISOS_POR_ROL[rol]) for rol in esperadas})

    def test_cargas_y_administracion_son_exclusivas_de_admin(self):
        exclusivas = (
            VIEW_CARGAR_CASOS,
            VIEW_CARGAR_INCIDENTES,
            VIEW_CARGAR_PROBLEMAS,
            VIEW_ADMINISTRAR_USUARIOS,
        )
        for vista in exclusivas:
            self.assertTrue(puede_acceder(ROLE_ADMIN, vista))
            for rol in (*ROLES_PERMITIDOS[1:], ROLE_VIEWER_LEGACY):
                self.assertFalse(puede_acceder(rol, vista))

    def test_rol_desconocido_no_recibe_permisos(self):
        self.assertEqual(frozenset(), obtener_permisos_rol("desconocido"))
        self.assertEqual((), obtener_vistas_permitidas("desconocido"))
        self.assertFalse(puede_acceder("desconocido", VISTAS[0]))
        self.assertFalse(puede_acceder(ROLE_ADMIN, "vista_inexistente"))

    def test_viewer_conserva_temporalmente_el_acceso_actual(self):
        vistas = obtener_permisos_rol(ROLE_VIEWER_LEGACY)
        self.assertEqual(13, len(vistas))
        self.assertIn(VIEW_REINCIDENCIAS_PROBLEMAS, vistas)
        self.assertNotIn(VIEW_DASHBOARD_INCIDENTES, vistas)
        self.assertNotIn(VIEW_ADMINISTRAR_USUARIOS, vistas)


if __name__ == "__main__":
    unittest.main()
