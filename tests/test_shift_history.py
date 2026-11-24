import unittest

import shift_history as sh


def parsed(days):
    """days: [(iso, {col: [names]})]"""
    import datetime as dt
    names = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    out = []
    for iso, assignments in days:
        d = dt.date.fromisoformat(iso)
        out.append({"date": iso, "dow": names[d.weekday()], "is_holiday": d.weekday() == 6, "assignments": assignments})
    return {"days": out}


class MergeUploadTests(unittest.TestCase):
    def test_file_across_two_months_then_the_next_month_does_not_double_count(self):
        # Dicembre 2026 generato 1/12-7/1, poi arriva il definitivo di gennaio 1-31.
        dec_jan = parsed([("2026-12-31", {"J": ["Cusmà"]}), ("2027-01-01", {"J": ["Crea"]}), ("2027-01-02", {"J": ["Dattilo"]})])
        hist, months = sh.merge_upload({}, dec_jan)
        self.assertEqual(months, ["2026-12", "2027-01"])

        jan = parsed([("2027-01-01", {"J": ["Crea"]}), ("2027-01-02", {"J": ["Trio"]}), ("2027-01-03", {"J": ["Trio"]})])
        hist, _ = sh.merge_upload(hist, jan)

        agg = sh.aggregate_multi_month(hist)
        self.assertEqual(agg["Crea"]["J"]["total"], 1, "1 gennaio contato una volta sola")
        self.assertEqual(agg["Trio"]["J"]["total"], 2)
        self.assertNotIn("Dattilo", agg, "il giorno ricaricato sostituisce quello vecchio")
        self.assertEqual(agg["Cusmà"]["J"]["total"], 1)
        self.assertNotIn("_days", agg)

    def test_night_before_a_period_comes_from_the_definitive_days(self):
        hist, _ = sh.merge_upload({}, parsed([("2026-11-30", {"J": ["Licordari"]})]))
        self.assertEqual(sh.night_doctors_on(hist, "2026-11-30"), ["Licordari"])
        self.assertEqual(sh.night_doctors_on(hist, "2026-11-29"), [])

    def test_valid_doctors_filter_and_months_counted(self):
        hist, _ = sh.merge_upload({}, parsed([("2026-10-04", {"D": ["Trio"], "E": ["Trio"], "C": ["nota a mano"]})]),
                                  valid_doctors={"Trio"})
        agg = sh.aggregate_multi_month(hist)
        self.assertEqual(list(agg), ["Trio"])
        self.assertEqual(agg["Trio"]["_festivi_DE_HI"], 1, "D+E della domenica contano una volta")
        self.assertEqual(agg["Trio"]["_months_counted"], 1)


if __name__ == "__main__":
    unittest.main()
