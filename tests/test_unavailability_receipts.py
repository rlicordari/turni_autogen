import unittest

import unavailability_receipts as ur


def row(day, shift, note=""):
    return {"doctor": "Migliorato", "date": day, "shift": shift, "note": note}


class ReceiptTests(unittest.TestCase):
    def setUp(self):
        self.rows = {
            "2026-10": [
                row("2026-10-07", "Ferie", "visita privata"),
                row("2026-10-04", "Tutto il giorno"),
            ]
        }
        self.diffs = {
            "2026-10": {
                "added_count": 1,
                "removed_count": 1,
                "details": {
                    "added": [{"date": "2026-10-07", "shift": "Ferie"}],
                    "removed": [{"date": "2026-10-12", "shift": "Mattina"}],
                },
            }
        }

    def test_receipt_lists_every_saved_day_in_local_time_without_notes(self):
        subject, body = ur.build_receipt(
            "Migliorato", self.rows, self.diffs,
            saved_at_utc="2026-09-30T09:56:00Z", commit_sha="ce2a156b0000",
        )

        self.assertIn("Migliorato", subject)
        self.assertIn("Ottobre 2026", subject)
        self.assertIn("30/09/2026 11:56", body)  # ora italiana (CEST)
        self.assertIn("ce2a156", body)
        full_list = body.split("Elenco completo registrato:")[1]
        self.assertLess(full_list.index("dom 04/10"), full_list.index("mer 07/10"))
        self.assertIn("mer 07/10  Ferie (con nota)", full_list)
        self.assertIn("lun 12/10  Mattina", body)  # rimossa
        self.assertNotIn("visita privata", body)

    def test_admin_receipt_says_who_changed_the_data(self):
        _subject, body = ur.build_receipt(
            "Migliorato", self.rows, self.diffs,
            saved_at_utc="2026-09-30T09:56:00Z", commit_sha=None, by_admin=True,
        )

        self.assertIn("amministratore", body)

    def test_empty_month_is_reported_explicitly(self):
        _subject, body = ur.build_receipt(
            "Migliorato", {"2026-10": []},
            {"2026-10": {"added_count": 0, "removed_count": 1, "details": {"added": [], "removed": [{"date": "2026-10-07", "shift": "Ferie"}]}}},
            saved_at_utc="2026-09-30T09:56:00Z", commit_sha=None,
        )

        self.assertIn("Nessuna indisponibilità registrata", body)


class RecipientTests(unittest.TestCase):
    def test_doctor_gets_mail_and_copies_are_deduplicated(self):
        to, cc = ur.recipients("Doc@Example.org", ["admin@example.org", "doc@example.org", "UTIC@polime.it", "utic@polime.it"])

        self.assertEqual(to, ["Doc@Example.org"])
        self.assertEqual(cc, ["admin@example.org", "UTIC@polime.it"])

    def test_without_doctor_email_copies_become_recipients(self):
        to, cc = ur.recipients("", ["admin@example.org", "non-una-mail"])

        self.assertEqual(to, ["admin@example.org"])
        self.assertEqual(cc, [])

    def test_parse_email_list_accepts_commas_semicolons_and_newlines(self):
        self.assertEqual(
            ur.parse_email_list("a@x.it; b@y.it,\n c@z.it , "),
            ["a@x.it", "b@y.it", "c@z.it"],
        )


if __name__ == "__main__":
    unittest.main()
