import unittest

import app_logic
from core.security import hash_password, normalizar_email, verificar_password


class SecurityTest(unittest.TestCase):
    def test_normaliza_email(self):
        self.assertEqual("persona@empresa.com", normalizar_email(" Persona@Empresa.COM "))

    def test_hash_y_verificacion(self):
        password_hash = hash_password("clave-segura")
        self.assertTrue(verificar_password("clave-segura", password_hash))
        self.assertFalse(verificar_password("otra-clave", password_hash))

    def test_hash_invalido_no_genera_excepcion(self):
        self.assertFalse(verificar_password("clave", "formato-invalido"))

    def test_app_logic_conserva_api_compatible(self):
        self.assertIs(app_logic.normalizar_email, normalizar_email)
        self.assertIs(app_logic.hash_password, hash_password)


if __name__ == "__main__":
    unittest.main()

