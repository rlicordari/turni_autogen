import datetime as dt
import unittest

import unavailability_calendar as ucal


def urow(day, shift, note=""):
    return {"id": f"u{day}{shift}", "Data": dt.date(2026, 10, day), "Fascia": shift, "Note": note}


def prow(day, shift, priority="media", note=""):
    return {"id": f"p{day}{shift}", "Data": dt.date(2026, 10, day), "Fascia": shift, "Priorita": priority, "Note": note}


def apply(unav, pref, event):
    return ucal.apply_event(unav, pref, event, year=2026, month=10)


def shifts_on(rows, day):
    return sorted(r["Fascia"] for r in rows if r["Data"].day == day)


class ApplyEventTests(unittest.TestCase):
    def test_day_can_hold_several_shifts(self):
        unav, pref = apply([], [], {"type": "set_day", "kind": "unav", "day": 15, "shifts": ["Pomeriggio", "Notte"], "note": "corso"})

        self.assertEqual(shifts_on(unav, 15), ["Notte", "Pomeriggio"])
        self.assertTrue(all(r["Note"] == "corso" for r in unav))
        self.assertEqual(pref, [])

    def test_full_day_shifts_are_exclusive(self):
        unav, _ = apply([], [], {"type": "set_day", "kind": "unav", "day": 7, "shifts": ["Mattina", "Ferie"]})
        self.assertEqual(shifts_on(unav, 7), ["Ferie"])

        unav, _ = apply([], [], {"type": "set_day", "kind": "unav", "day": 7, "shifts": ["Notte", "Tutto il giorno"]})
        self.assertEqual(shifts_on(unav, 7), ["Tutto il giorno"])

    def test_set_day_replaces_only_that_day_and_kind(self):
        unav = [urow(7, "Mattina"), urow(8, "Notte")]
        pref = [prow(7, "Notte")]

        unav, pref = apply(unav, pref, {"type": "set_day", "kind": "unav", "day": 7, "shifts": ["Pomeriggio"]})

        self.assertEqual(shifts_on(unav, 7), ["Pomeriggio"])
        self.assertEqual(shifts_on(unav, 8), ["Notte"])
        self.assertEqual(shifts_on(pref, 7), ["Notte"])

    def test_empty_selection_clears_the_day(self):
        unav, _ = apply([urow(7, "Mattina"), urow(7, "Notte")], [], {"type": "set_day", "kind": "unav", "day": 7, "shifts": []})
        self.assertEqual(unav, [])

    def test_range_sets_ferie_on_every_day(self):
        unav, _ = apply([urow(8, "Notte")], [], {"type": "range", "from": 7, "to": 12})

        self.assertEqual(sorted(r["Data"].day for r in unav), [7, 8, 9, 10, 11, 12])
        self.assertEqual({r["Fascia"] for r in unav}, {"Ferie"})

    def test_preferences_keep_priority_and_refuse_ferie(self):
        _, pref = apply([], [], {"type": "set_day", "kind": "pref", "day": 14, "shifts": ["Notte", "Ferie"], "priority": "alta"})

        self.assertEqual(shifts_on(pref, 14), ["Notte"])
        self.assertEqual(pref[0]["Priorita"], "alta")

    def test_invalid_events_change_nothing(self):
        unav = [urow(7, "Mattina")]
        for event in (
            {"type": "set_day", "kind": "unav", "day": 32, "shifts": ["Mattina"]},
            {"type": "set_day", "kind": "boh", "day": 7, "shifts": ["Mattina"]},
            {"type": "set_day", "kind": "unav", "day": 7, "shifts": ["Colazione"]},
            {"type": "range", "from": 12, "to": 7},
            {"type": "sconosciuto"},
        ):
            new_unav, new_pref = apply(unav, [], event)
            if event.get("shifts") == ["Colazione"]:
                self.assertEqual(new_unav, [], "fascia sconosciuta = nessuna fascia valida")
            else:
                self.assertEqual(new_unav, unav, event)
            self.assertEqual(new_pref, [])


def edit(seq, event, year=2026, month=10):
    return {**event, "seq": seq, "year": year, "month": month}


def apply_value(unav, pref, value, ack=None):
    return ucal.apply_value(unav, pref, value, year=2026, month=10, ack=ack or {})


class ApplyValueTests(unittest.TestCase):
    """The component resends every edit until the page acknowledges it."""

    def test_two_taps_in_one_value_are_both_applied(self):
        value = {"cid": "a", "edits": [
            edit(1, {"type": "set_day", "kind": "unav", "day": 3, "shifts": ["Mattina"]}),
            edit(2, {"type": "set_day", "kind": "unav", "day": 4, "shifts": ["Notte"]}),
        ]}

        unav, _, ack, nav = apply_value([], [], value)

        self.assertEqual([(r["Data"].day, r["Fascia"]) for r in unav], [(3, "Mattina"), (4, "Notte")])
        self.assertEqual(ack, {"cid": "a", "seq": 2})
        self.assertIsNone(nav)

    def test_acknowledged_edits_are_not_applied_twice(self):
        value = {"cid": "a", "edits": [edit(1, {"type": "set_day", "kind": "unav", "day": 3, "shifts": ["Mattina"]})]}
        unav, pref, ack, _ = apply_value([], [], value)
        unav = [{**r, "Fascia": "Notte"} for r in unav]      # modificato poi da altro (es. ferie lunghe)

        again, _, ack2, _ = apply_value(unav, pref, value, ack)

        self.assertEqual(again, unav)
        self.assertEqual(ack2, ack)

    def test_edits_are_applied_in_sequence_order(self):
        value = {"cid": "a", "edits": [
            edit(2, {"type": "set_day", "kind": "unav", "day": 8, "shifts": ["Notte"]}),
            edit(1, {"type": "range", "from": 7, "to": 9}),
        ]}

        unav, _, _, _ = apply_value([], [], value)

        self.assertEqual(shifts_on(unav, 8), ["Notte"])
        self.assertEqual(shifts_on(unav, 9), ["Ferie"])

    def test_new_component_instance_starts_from_zero(self):
        value = {"cid": "b", "edits": [edit(1, {"type": "set_day", "kind": "unav", "day": 3, "shifts": ["Mattina"]})]}

        unav, _, ack, _ = apply_value([], [], value, ack={"cid": "a", "seq": 9})

        self.assertEqual(len(unav), 1)
        self.assertEqual(ack, {"cid": "b", "seq": 1})

    def test_edit_of_another_month_is_dropped_but_acknowledged(self):
        value = {"cid": "a", "edits": [edit(1, {"type": "set_day", "kind": "unav", "day": 3, "shifts": ["Mattina"]}, month=11)]}

        unav, _, ack, _ = apply_value([], [], value)

        self.assertEqual(unav, [])
        self.assertEqual(ack, {"cid": "a", "seq": 1})

    def test_month_change_comes_after_the_edits(self):
        value = {"cid": "a",
                 "edits": [edit(1, {"type": "set_day", "kind": "unav", "day": 3, "shifts": ["Mattina"]})],
                 "nav": {"seq": 2, "year": 2026, "month": 11}}

        unav, _, ack, nav = apply_value([], [], value)

        self.assertEqual(len(unav), 1)
        self.assertEqual(nav, (2026, 11))
        self.assertEqual(ack, {"cid": "a", "seq": 2})
        self.assertIsNone(apply_value(unav, [], value, ack)[3], "cambio mese già confermato")

    def test_garbage_changes_nothing(self):
        for value in (None, "x", {"cid": "a", "edits": "x"}, {"cid": "a", "edits": [{"seq": "?"}]}):
            unav, pref, ack, nav = apply_value([urow(3, "Mattina")], [], value)
            self.assertEqual([r["Fascia"] for r in unav], ["Mattina"], value)
            self.assertIsNone(nav)


class PayloadTests(unittest.TestCase):
    def test_payload_marks_days_that_differ_from_server(self):
        sent_unav = [{"doctor": "Rossi", "date": "2026-10-03", "shift": "Tutto il giorno", "note": ""}]
        unav = [urow(3, "Tutto il giorno"), urow(7, "Ferie")]

        payload = ucal.calendar_payload(unav, [], sent_unav, [], year=2026, month=10, limits={"per_shift": 6})

        days = {d["day"]: d for d in payload["days"]}
        self.assertEqual(len(days), 31)
        self.assertEqual(payload["first_weekday"], 3)       # 1/10/2026 è giovedì
        self.assertFalse(days[3]["pending"])
        self.assertTrue(days[7]["pending"])
        self.assertEqual(days[7]["unav"], [{"shift": "Ferie", "note": ""}])
        self.assertEqual(payload["label"], "Ottobre 2026")
        self.assertEqual(payload["ack"], {})

    def test_payload_returns_the_acknowledgement(self):
        payload = ucal.calendar_payload([], [], [], [], year=2026, month=10, limits={}, ack={"cid": "a", "seq": 4})
        self.assertEqual(payload["ack"], {"cid": "a", "seq": 4})

    def test_removed_day_is_pending(self):
        sent_unav = [{"doctor": "Rossi", "date": "2026-10-03", "shift": "Notte", "note": ""}]

        payload = ucal.calendar_payload([], [], sent_unav, [], year=2026, month=10, limits={})

        self.assertTrue(payload["days"][2]["pending"])

    def test_national_holidays_are_flagged(self):
        payload = ucal.calendar_payload([], [], [], [], year=2026, month=11, limits={})
        self.assertTrue(payload["days"][0]["festive"])        # 1 novembre
        self.assertFalse(payload["days"][1]["festive"])       # 2 novembre, lunedì


if __name__ == "__main__":
    unittest.main()
