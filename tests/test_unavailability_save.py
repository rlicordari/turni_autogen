import datetime as dt
import unittest

import unavailability_store as ustore


class FakeStore:
    """In-memory stand-in for one doctor's CSV on GitHub (optimistic locking by sha)."""

    def __init__(self, rows=None):
        self.rows = [dict(r) for r in (rows or [])]
        self.version = 1
        self.puts = 0
        self.stale_reads = 0  # next N reads return the previous version (GitHub replica lag)
        self._previous = [dict(r) for r in self.rows]
        self._previous_version = self.version

    def load(self):
        if self.stale_reads > 0:
            self.stale_reads -= 1
            return [dict(r) for r in self._previous], f"sha{self._previous_version}"
        return [dict(r) for r in self.rows], f"sha{self.version}"

    def save(self, rows, sha):
        if sha != f"sha{self.version}":
            raise ustore.ShaConflictError("409 conflict")
        self._previous, self._previous_version = self.rows, self.version
        self.rows = [dict(r) for r in rows]
        self.version += 1
        self.puts += 1
        return {"content_sha": f"sha{self.version}", "commit_sha": f"commit{self.version}"}


def row(doctor, day, shift, note=""):
    return {"doctor": doctor, "date": day, "shift": shift, "note": note, "updated_at": "t0", "priority": "media"}


def save(store, entries_by_month, base_rows, **kw):
    base_signatures = {
        mk: ustore.month_signature(base_rows, "Migliorato", mk[0], mk[1]) for mk in entries_by_month
    }
    return ustore.save_doctor_months(
        load_fn=store.load,
        save_fn=store.save,
        doctor="Migliorato",
        entries_by_month=entries_by_month,
        base_signatures=base_signatures,
        updated_at="2026-09-30T10:00:00Z",
        sleep_fn=lambda _s: None,
        **kw,
    )


class SaveDoctorMonthsTests(unittest.TestCase):
    def test_stale_editor_cannot_erase_days_saved_by_another_session(self):
        # Il PC ha aperto ottobre vuoto; poi dal telefono sono stati salvati 7-8 ottobre.
        base_rows = []
        store = FakeStore([row("Migliorato", "2026-10-07", "Ferie"), row("Migliorato", "2026-10-08", "Ferie")])
        pc_entries = {(2026, 10): [(dt.date(2026, 10, 20), "Mattina", "")]}

        with self.assertRaises(ustore.MonthConflictError) as ctx:
            save(store, pc_entries, base_rows)

        self.assertEqual(ctx.exception.month_key, "2026-10")
        self.assertEqual(store.puts, 0)
        self.assertEqual(len(store.rows), 2)

    def test_change_in_another_month_is_merged_not_overwritten(self):
        base_rows = [row("Migliorato", "2026-10-04", "Tutto il giorno")]
        store = FakeStore(base_rows + [row("Migliorato", "2026-11-02", "Notte")])
        entries = {(2026, 10): [(dt.date(2026, 10, 4), "Tutto il giorno", ""), (dt.date(2026, 10, 7), "Ferie", "")]}

        outcome = save(store, entries, base_rows)

        self.assertTrue(outcome.changed)
        dates = sorted(r["date"] for r in store.rows)
        self.assertEqual(dates, ["2026-10-04", "2026-10-07", "2026-11-02"])
        self.assertEqual(outcome.diffs["2026-10"]["added_count"], 1)

    def test_unchanged_save_does_not_write(self):
        base_rows = [row("Migliorato", "2026-10-04", "Tutto il giorno")]
        store = FakeStore(base_rows)

        outcome = save(store, {(2026, 10): [(dt.date(2026, 10, 4), "Tutto il giorno", "")]}, base_rows)

        self.assertFalse(outcome.changed)
        self.assertEqual(store.puts, 0)

    def test_double_submit_after_success_is_a_no_op(self):
        base_rows = []
        store = FakeStore([])
        entries = {(2026, 10): [(dt.date(2026, 10, 7), "Ferie", "")]}

        first = save(store, entries, base_rows)
        # Il secondo tap parte dallo stesso editor (base vuota) ma il server ha gia' quei dati.
        second = save(store, entries, base_rows)

        self.assertTrue(first.changed)
        self.assertFalse(second.changed)
        self.assertEqual(store.puts, 1)

    def test_sha_conflict_is_retried_when_month_did_not_change(self):
        base_rows = []
        store = FakeStore([])
        calls = {"n": 0}
        real_save = store.save

        def flaky_save(rows, sha):
            calls["n"] += 1
            if calls["n"] == 1:
                raise ustore.ShaConflictError("409 conflict")
            return real_save(rows, sha)

        outcome = ustore.save_doctor_months(
            load_fn=store.load,
            save_fn=flaky_save,
            doctor="Migliorato",
            entries_by_month={(2026, 10): [(dt.date(2026, 10, 7), "Ferie", "")]},
            base_signatures={(2026, 10): []},
            updated_at="2026-09-30T10:00:00Z",
            sleep_fn=lambda _s: None,
        )

        self.assertTrue(outcome.changed)
        self.assertEqual(store.puts, 1)

    def test_read_back_lag_is_retried_without_writing_twice(self):
        base_rows = []
        store = FakeStore([])
        entries = {(2026, 10): [(dt.date(2026, 10, 7), "Ferie", "")]}
        real_save = store.save

        def save_then_lag(rows, sha):
            result = real_save(rows, sha)
            store.stale_reads = 2
            return result

        outcome = ustore.save_doctor_months(
            load_fn=store.load,
            save_fn=save_then_lag,
            doctor="Migliorato",
            entries_by_month=entries,
            base_signatures={(2026, 10): []},
            updated_at="2026-09-30T10:00:00Z",
            sleep_fn=lambda _s: None,
        )

        self.assertTrue(outcome.changed)
        self.assertEqual(store.puts, 1)
        self.assertEqual(outcome.commit_sha, "commit2")
        self.assertEqual([r["date"] for r in outcome.verified_rows], ["2026-10-07"])

    def test_write_that_landed_but_lost_its_response_still_counts_as_change(self):
        # Timeout lato client dopo che GitHub ha già applicato il commit: il retry
        # trova il mese già aggiornato. Deve risultare "modificato" (audit + mail).
        store = FakeStore([])
        real_save = store.save
        calls = {"n": 0}

        def save_lost_response(rows, sha):
            calls["n"] += 1
            real_save(rows, sha)
            if calls["n"] == 1:
                raise ustore.ShaConflictError("timeout: risposta persa")
            return {}

        outcome = ustore.save_doctor_months(
            load_fn=store.load,
            save_fn=save_lost_response,
            doctor="Migliorato",
            entries_by_month={(2026, 10): [(dt.date(2026, 10, 7), "Ferie", "")]},
            base_signatures={(2026, 10): []},
            updated_at="2026-09-30T10:00:00Z",
            sleep_fn=lambda _s: None,
        )

        self.assertTrue(outcome.changed)
        self.assertEqual(outcome.diffs["2026-10"]["added_count"], 1)
        self.assertEqual(store.puts, 1)

    def test_read_back_that_never_matches_raises(self):
        store = FakeStore([])
        real_save = store.save

        def save_then_lose(rows, sha):
            result = real_save(rows, sha)
            store.stale_reads = 99
            return result

        with self.assertRaises(ustore.SaveNotVerifiedError):
            ustore.save_doctor_months(
                load_fn=store.load,
                save_fn=save_then_lose,
                doctor="Migliorato",
                entries_by_month={(2026, 10): [(dt.date(2026, 10, 7), "Ferie", "")]},
                base_signatures={(2026, 10): []},
                updated_at="2026-09-30T10:00:00Z",
                sleep_fn=lambda _s: None,
            )


class MonthSignatureTests(unittest.TestCase):
    def test_signature_filters_doctor_and_month_and_normalizes_shift(self):
        rows = [
            row("Migliorato", "2026-10-07", "ferie", "x"),
            row("Migliorato", "2026-11-01", "Mattina"),
            row("Crea", "2026-10-07", "Mattina"),
        ]

        self.assertEqual(ustore.month_signature(rows, "Migliorato", 2026, 10), [("2026-10-07", "Ferie", "x")])

    def test_entries_signature_matches_month_signature(self):
        rows = [row("Migliorato", "2026-10-07", "Ferie")]
        entries = [(dt.date(2026, 10, 7), "Ferie", "")]

        self.assertEqual(ustore.entries_signature(entries), ustore.month_signature(rows, "Migliorato", 2026, 10))


if __name__ == "__main__":
    unittest.main()
