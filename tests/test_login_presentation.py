import unittest
from unittest.mock import patch
import app_ui
from streamlit.testing.v1 import AppTest

class LoginPresentationTest(unittest.TestCase):
    def portal(self):
        return AppTest.from_string('import app_ui\napp_ui.aplicar_tema_visual()\napp_ui.login()').run(timeout=15)

    def test_native_form_and_invalid_email(self):
        at = self.portal()
        self.assertFalse(at.exception)
        self.assertEqual(at.text_input[1].label, 'Contrase\u00f1a')
        with patch.object(app_ui, 'autenticar_usuario') as auth:
            at.text_input[0].set_value('correo-invalido')
            at.button[0].click().run()
            auth.assert_not_called()
        self.assertIn('correo', at.error[0].value)

    def test_rejected_credentials_keep_form_available(self):
        at = self.portal()
        with patch.object(app_ui, 'autenticar_usuario', return_value=None) as auth:
            at.text_input[0].set_value('prueba@example.com')
            at.text_input[1].set_value('incorrecta')
            at.button[0].click().run()
            auth.assert_called_once_with('prueba@example.com', 'incorrecta')
        self.assertFalse(at.exception)
        self.assertTrue(at.error)
        self.assertEqual(len(at.text_input), 2)

if __name__ == '__main__':
    unittest.main()
