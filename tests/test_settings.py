import os
import unittest
from unittest.mock import patch

from core.settings import config_value


class SettingsTest(unittest.TestCase):
    def test_variable_de_entorno_tiene_prioridad(self):
        with patch.dict(os.environ, {"APP_TEST_SETTING": "desde-entorno"}):
            self.assertEqual("desde-entorno", config_value("APP_TEST_SETTING"))

    def test_retorna_default_si_no_existe_configuracion(self):
        with patch.dict(os.environ, {}, clear=True):
            self.assertEqual("valor-default", config_value("CONFIG_INEXISTENTE_PRUEBA", "valor-default"))


if __name__ == "__main__":
    unittest.main()

