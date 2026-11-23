import datetime as dt
import unittest

import festivi_store as fs

D = dt.date

YAML = """
festivi_extra:
  - "2026-06-03"  # Madonna della Lettera
  - date: "2026-12-31"
    name: "S. Silvestro"
entries:
  - date: "2026-12-25"
    doctor: "Pugliatti"
    shift: "Mattina"
    note: "Natale"
  - date: "2026-12-24"
    doctor: "Trio"
    shift: "Notte"
"""

DAY_DOCS = ["Allegra", "Crea", "Cusmà", "Grimaldi", "Licordari", "Trio"]
NIGHT_DOCS = ["Allegra", "Crea", "Cusmà", "Licordari", "Trio"]


class ParseTests(unittest.TestCase):
    def test_old_and_new_extra_formats_and_round_trip(self):
        data = fs.parse(YAML)

        self.assertEqual(data["festivi_extra"], [
            {"date": "2026-06-03", "name": ""},
            {"date": "2026-12-31", "name": "S. Silvestro"},
        ])
        self.assertEqual(fs.parse(fs.to_text(data)), data)

    def test_solver_format_maps_shifts_to_columns(self):
        out = fs.to_solver(fs.parse(YAML))

        self.assertEqual(out["festivi_extra"], ["2026-06-03", "2026-12-31"])
        self.assertIn({"doctor": "Pugliatti", "date": "2026-12-25", "column": "D"}, out["fixed_assignments"])
        self.assertIn({"doctor": "Trio", "date": "2026-12-24", "column": "J"}, out["fixed_assignments"])


class SlotTests(unittest.TestCase):
    def test_festive_days_with_eve_nights(self):
        data = fs.parse(YAML)
        slots = fs.slots(D(2026, 12, 1), D(2027, 1, 7), data["festivi_extra"])
        keys = [(s["date"], s["shift"]) for s in slots]

        # Immacolata: notte del 7 (prefestivo) + mattina/pomeriggio/notte dell'8
        self.assertIn(("2026-12-07", "Notte"), keys)
        self.assertNotIn(("2026-12-07", "Mattina"), keys)
        for sh in ("Mattina", "Pomeriggio", "Notte"):
            self.assertIn(("2026-12-08", sh), keys)
        # 25 festivo → nessun "prefestivo" separato per il 26
        self.assertEqual(sum(1 for k in keys if k[0] == "2026-12-25"), 3)
        self.assertIn(("2026-12-24", "Notte"), keys)
        self.assertIn(("2026-12-31", "Mattina"), keys)       # festivo extra
        self.assertIn(("2027-01-05", "Notte"), keys)          # pre-Epifania
        labels = {s["date"]: s["label"] for s in slots}
        self.assertEqual(labels["2026-12-25"], "Natale")
        self.assertEqual(labels["2026-12-24"], "Pre-Natale")
        self.assertEqual(labels["2026-12-31"], "S. Silvestro")

    def test_easter_sunday_is_included(self):
        slots = fs.slots(D(2027, 3, 26), D(2027, 3, 30), [])
        self.assertIn(("2027-03-28", "Mattina"), [(s["date"], s["shift"]) for s in slots])
        self.assertIn(("2027-03-27", "Notte"), [(s["date"], s["shift"]) for s in slots])


class DrawTests(unittest.TestCase):
    def setUp(self):
        self.slots = fs.slots(D(2026, 12, 1), D(2027, 1, 7), [{"date": "2026-12-31", "name": "S. Silvestro"}])

    def draw(self, entries=(), **kw):
        return fs.draw(self.slots, list(entries), day_doctors=DAY_DOCS, night_doctors=NIGHT_DOCS, seed=kw.pop("seed", 1), **kw)

    def test_fills_every_slot_respecting_participants(self):
        out = self.draw()
        by_key = {(e["date"], e["shift"]): e["doctor"] for e in out}

        self.assertEqual(set(by_key), {(s["date"], s["shift"]) for s in self.slots})
        for (d, sh), doc in by_key.items():
            self.assertIn(doc, NIGHT_DOCS if sh == "Notte" else DAY_DOCS)

    def test_keeps_existing_assignments_when_only_empty(self):
        out = self.draw([{"date": "2026-12-25", "doctor": "Grimaldi", "shift": "Mattina", "note": ""}])
        self.assertIn({"date": "2026-12-25", "doctor": "Grimaldi", "shift": "Mattina", "note": ""}, out)

    def test_balanced_and_no_doctor_twice_a_day_or_morning_after_night(self):
        out = self.draw()
        counts = {}
        for e in out:
            counts[e["doctor"]] = counts.get(e["doctor"], 0) + 1
        self.assertLessEqual(max(counts.values()) - min(counts.values()), 1, counts)
        per_day = {}
        for e in out:
            per_day.setdefault(e["date"], []).append(e["doctor"])
        for d, docs in per_day.items():
            self.assertEqual(len(docs), len(set(docs)), (d, docs))
        nights = {e["date"]: e["doctor"] for e in out if e["shift"] == "Notte"}
        for e in out:
            if e["shift"] in ("Mattina", "Pomeriggio"):
                prev = (D.fromisoformat(e["date"]) - dt.timedelta(days=1)).isoformat()
                self.assertNotEqual(nights.get(prev), e["doctor"], e)

    def test_same_seed_same_draw_other_seed_other_draw(self):
        self.assertEqual(self.draw(seed=5), self.draw(seed=5))
        self.assertNotEqual(self.draw(seed=5), self.draw(seed=6))


class CheckTests(unittest.TestCase):
    def test_warns_inactive_double_and_night_then_morning(self):
        entries = [
            {"date": "2026-12-25", "doctor": "Cimino", "shift": "Pomeriggio", "note": ""},
            {"date": "2026-12-26", "doctor": "Trio", "shift": "Mattina", "note": ""},
            {"date": "2026-12-26", "doctor": "Trio", "shift": "Notte", "note": ""},
            {"date": "2026-12-25", "doctor": "Cusmà", "shift": "Notte", "note": ""},
            {"date": "2026-12-26", "doctor": "Cusmà", "shift": "Pomeriggio", "note": ""},
        ]
        warnings = " | ".join(fs.check(entries, active_doctors={"Trio", "Cusmà"}))

        self.assertIn("Cimino", warnings)
        self.assertIn("non attivo", warnings)
        self.assertIn("Trio", warnings)
        self.assertIn("due turni", warnings)
        self.assertIn("Cusmà", warnings)
        self.assertIn("dopo la notte", warnings)


if __name__ == "__main__":
    unittest.main()
