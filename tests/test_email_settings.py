import json
import unittest

import email_settings as es

SECRET = "ghp_token_di_prova_molto_lungo_1234567890"
NOW = "2026-10-01T08:00:00Z"


def saved(previous=None, **kw):
    fields = {"username": "turni.utic@gmail.com", "from_addr": "", "host": "", "port": 0,
              "starttls": True, "new_password": "abcd efgh ijkl mnop"}
    fields.update(kw)
    return es.updated_settings(previous, secret=SECRET, now=NOW, **fields)


class PasswordEncryptionTests(unittest.TestCase):
    def test_password_is_stored_only_encrypted(self):
        stored = saved()

        dump = json.dumps(stored)
        self.assertNotIn("abcd efgh ijkl mnop", dump)
        self.assertNotIn("abcdefghijklmnop", dump)
        self.assertEqual(es.decrypt_password(stored["password_enc"], SECRET), "abcdefghijklmnop")

    def test_other_key_cannot_read_it(self):
        stored = saved()
        self.assertIsNone(es.decrypt_password(stored["password_enc"], "altro-token"))
        self.assertIsNone(es.decrypt_password("spazzatura", SECRET))

    def test_empty_password_keeps_the_previous_one(self):
        first = saved()
        second = saved(first, new_password="", username="altro@gmail.com")

        self.assertEqual(second["password_enc"], first["password_enc"])
        self.assertEqual(second["username"], "altro@gmail.com")

    def test_sender_defaults_to_the_account_and_gmail_server(self):
        stored = saved()
        cfg, warnings = es.effective_smtp({}, stored, SECRET)

        self.assertEqual(cfg["from"], "turni.utic@gmail.com")
        self.assertEqual((cfg["host"], cfg["port"], cfg["starttls"]), ("smtp.gmail.com", 587, True))
        self.assertEqual(cfg["password"], "abcdefghijklmnop")
        self.assertEqual(warnings, [])


class PasswordFormatTests(unittest.TestCase):
    def test_google_app_password_loses_its_spaces_on_any_server(self):
        stored = saved(host="smtp.altro.it", new_password=" abcd efgh ijkl mnop ")
        self.assertEqual(es.decrypt_password(stored["password_enc"], SECRET), "abcdefghijklmnop")

    def test_other_passwords_are_kept_as_typed(self):
        stored = saved(host="smtp.altro.it", new_password="la mia password 1")
        self.assertEqual(es.decrypt_password(stored["password_enc"], SECRET), "la mia password 1")


class EffectiveConfigTests(unittest.TestCase):
    SECRETS = {"host": "smtp.gmail.com", "port": 587, "username": "vecchio@gmail.com",
               "password": "vecchiapassword", "from": "vecchio@gmail.com", "starttls": True}

    def test_without_panel_settings_secrets_are_used(self):
        cfg, warnings = es.effective_smtp(self.SECRETS, None, SECRET)
        self.assertEqual(cfg, self.SECRETS)
        self.assertEqual(warnings, [])

    def test_panel_settings_win_over_secrets(self):
        cfg, _ = es.effective_smtp(self.SECRETS, saved(), SECRET)

        self.assertEqual(cfg["username"], "turni.utic@gmail.com")
        self.assertEqual(cfg["from"], "turni.utic@gmail.com")
        self.assertEqual(cfg["password"], "abcdefghijklmnop")

    def test_unreadable_password_falls_back_to_secrets_with_a_warning(self):
        cfg, warnings = es.effective_smtp(self.SECRETS, saved(), "token-cambiato")

        self.assertEqual(cfg["password"], "vecchiapassword")
        self.assertEqual(len(warnings), 1)
        self.assertIn("reinseriscila", warnings[0])

    def test_public_view_never_contains_the_password(self):
        view = es.public_view(saved())
        self.assertNotIn("password_enc", view)
        self.assertTrue(view["password_set"])


if __name__ == "__main__":
    unittest.main()
