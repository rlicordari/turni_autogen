import unittest
from datetime import date, datetime

import pandas as pd

from fixed_assignment_options import parse_fixed_assignment_rows


class FixedAssignmentOptionsTests(unittest.TestCase):
    def setUp(self):
        self.options = {
            "start": date(2026, 11, 1),
            "end": date(2026, 11, 30),
            "doctors": ["Trio", "Manganaro", "Licordari"],
            "shift_labels": {
                "Notte (J)": "J",
                "UTIC pomeriggio (H)": "H",
                "Interni (T)": "T",
            },
        }

    def test_new_editor_rows_with_string_dates_are_all_passed_to_solver(self):
        rows = pd.DataFrame([
            {"Medico": "Trio", "Giorno": "2026-11-07", "Turno/Colonna": "Notte (J)"},
            {"Medico": "Trio", "Giorno": "2026-11-09", "Turno/Colonna": "UTIC pomeriggio (H)"},
            {"Medico": "Manganaro", "Giorno": "2026-11-09", "Turno/Colonna": "Interni (T)"},
            {"Medico": "Licordari", "Giorno": "2026-11-09", "Turno/Colonna": "Notte (J)"},
        ])

        self.assertEqual(parse_fixed_assignment_rows(rows, **self.options), [
            {"doctor": "Trio", "date": "2026-11-07", "column": "J"},
            {"doctor": "Trio", "date": "2026-11-09", "column": "H"},
            {"doctor": "Manganaro", "date": "2026-11-09", "column": "T"},
            {"doctor": "Licordari", "date": "2026-11-09", "column": "J"},
        ])

    def test_existing_dates_and_display_format_are_supported(self):
        rows = pd.DataFrame([
            {"Medico": "Trio", "Giorno": date(2026, 11, 7), "Turno/Colonna": "Notte (J)"},
            {"Medico": "Manganaro", "Giorno": datetime(2026, 11, 9), "Turno/Colonna": "Interni (T)"},
            {"Medico": "Licordari", "Giorno": "09/11/2026", "Turno/Colonna": "Notte (J)"},
        ])
        self.assertEqual(len(parse_fixed_assignment_rows(rows, **self.options)), 3)

    def test_incomplete_row_is_rejected_instead_of_silently_dropped(self):
        rows = pd.DataFrame([
            {"Medico": "Trio", "Giorno": "2026-11-07", "Turno/Colonna": ""},
        ])
        with self.assertRaisesRegex(ValueError, "completa"):
            parse_fixed_assignment_rows(rows, **self.options)

    def test_invalid_shift_is_not_converted_to_night(self):
        rows = pd.DataFrame([
            {"Medico": "Trio", "Giorno": "2026-11-07", "Turno/Colonna": "Altro"},
        ])
        with self.assertRaisesRegex(ValueError, "turno non riconosciuto"):
            parse_fixed_assignment_rows(rows, **self.options)


if __name__ == "__main__":
    unittest.main()
