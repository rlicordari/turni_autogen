import unittest

import yaml

import doctor_contacts as dc

TEXT = """# commento
"Migliorato":
  email: "vecchia@example.org"
  phone: "+39333"
"Rossi":
  email: "rossi@example.org"
"""


class SetEmailTests(unittest.TestCase):
    def test_changes_only_that_doctor_and_keeps_the_phone(self):
        data = yaml.safe_load(dc.set_email(TEXT, "Migliorato", " nuova@example.org "))

        self.assertEqual(data["Migliorato"], {"email": "nuova@example.org", "phone": "+39333"})
        self.assertEqual(data["Rossi"], {"email": "rossi@example.org"})

    def test_new_doctor_is_added(self):
        data = yaml.safe_load(dc.set_email(TEXT, "Calabrò", "calabro@example.org"))
        self.assertEqual(data["Calabrò"], {"email": "calabro@example.org"})

    def test_empty_email_removes_it(self):
        data = yaml.safe_load(dc.set_email(TEXT, "Migliorato", ""))
        self.assertEqual(data["Migliorato"], {"phone": "+39333"})
        data = yaml.safe_load(dc.set_email(TEXT, "Rossi", ""))
        self.assertNotIn("Rossi", data)

    def test_invalid_email_is_refused(self):
        for bad in ("senza-chiocciola", "a@b", "due@@x.it", "uno@x.it, due@y.it"):
            with self.assertRaises(ValueError, msg=bad):
                dc.set_email(TEXT, "Migliorato", bad)

    def test_empty_file(self):
        data = yaml.safe_load(dc.set_email("", "Migliorato", "m@example.org"))
        self.assertEqual(data, {"Migliorato": {"email": "m@example.org"}})


if __name__ == "__main__":
    unittest.main()
