import unittest
from unittest.mock import MagicMock, patch

from diagnose_servicenow_oauth import diagnose


class OAuthDiagnosticTest(unittest.TestCase):
    def run_response(self, status, payload):
        session = MagicMock()
        session.__enter__.return_value = session
        session.post.return_value.status_code = status
        session.post.return_value.headers = {}
        session.post.return_value.json.return_value = payload
        output = []
        with patch("diagnose_servicenow_oauth.requests.Session", return_value=session):
            result = diagnose("test.service-now.com", "client", "private-secret", output.append)
        return result, "\n".join(output), session

    def test_token_and_private_response_never_printed(self):
        result, output, session = self.run_response(200, {"access_token": "private-token"})
        self.assertEqual(0, result)
        self.assertIn("TOKEN EMITIDO", output)
        self.assertNotIn("private-token", output)
        self.assertNotIn("private-secret", output)
        self.assertFalse(session.trust_env)
        self.assertFalse(session.post.call_args.kwargs["allow_redirects"])

    def test_actual_server_error_preserved_and_free_text_omitted(self):
        result, output, _ = self.run_response(401, {
            "error": "server_error", "error_description": "access_denied",
            "extra": "private-secret private-token",
        })
        self.assertEqual(1, result)
        self.assertIn("OAuth: server_error", output)
        self.assertIn("Descripción OAuth: access_denied", output)
        self.assertNotIn("private-secret", output)
        _, unknown_output, _ = self.run_response(401, {
            "error": "private-secret", "error_description": "private-token",
        })
        self.assertNotIn("private-secret", unknown_output)
        self.assertNotIn("private-token", unknown_output)

    def test_encrypted_secret_or_foreign_host_never_sent(self):
        for instance, secret in (("test.service-now.com", "\ufddeencrypted\ufdef"),
                                 ("foreign.example.com", "private-secret")):
            with self.subTest(instance=instance), patch("diagnose_servicenow_oauth.requests.Session") as session:
                self.assertEqual(2, diagnose(instance, "client", secret, lambda _: None))
                session.assert_not_called()
