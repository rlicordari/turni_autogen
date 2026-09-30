import datetime as dt
import unittest

import unavailability_drafts as ud
import unavailability_store as ustore


def row(doctor, day, shift, note=""):
    return {"doctor": doctor, "date": day, "shift": shift, "note": note, "updated_at": "t0", "priority": "media"}


class DraftTests(unittest.TestCase):
    def test_draft_round_trips_entries(self):
        draft = ud.set_month(
            ud.empty_draft("Migliorato"),
            "2026-10",
            entries=[(dt.date(2026, 10, 7), "Ferie", ""), (dt.date(2026, 10, 8), "Ferie", "")],
            base_signature=[],
            updated_at="2026-09-29T11:55:00Z",
            session_id="s1",
        )

        restored = ud.from_text(ud.to_text(draft))

        self.assertEqual(
            ud.month_entries(restored, "2026-10"),
            [(dt.date(2026, 10, 7), "Ferie", ""), (dt.date(2026, 10, 8), "Ferie", "")],
        )
        self.assertEqual(restored["months"]["2026-10"]["updated_at"], "2026-09-29T11:55:00Z")

    def test_editor_resumes_unsent_draft_built_on_current_server_data(self):
        official = [row("Migliorato", "2026-10-04", "Tutto il giorno")]
        official_sig = ustore.month_signature(official, "Migliorato", 2026, 10)
        draft = ud.set_month(
            ud.empty_draft("Migliorato"),
            "2026-10",
            entries=[(dt.date(2026, 10, 4), "Tutto il giorno", ""), (dt.date(2026, 10, 7), "Ferie", "")],
            base_signature=official_sig,
            updated_at="2026-09-29T11:55:00Z",
            session_id="s1",
        )

        start = ud.editor_start(draft, "2026-10", official_sig)

        self.assertEqual(start["action"], "resume")
        self.assertEqual(len(start["entries"]), 2)

    def test_editor_offers_draft_when_server_changed_after_it(self):
        official = [row("Migliorato", "2026-10-20", "Mattina")]
        draft = ud.set_month(
            ud.empty_draft("Migliorato"),
            "2026-10",
            entries=[(dt.date(2026, 10, 7), "Ferie", "")],
            base_signature=[],
            updated_at="2026-09-29T11:55:00Z",
            session_id="s1",
        )

        start = ud.editor_start(draft, "2026-10", ustore.month_signature(official, "Migliorato", 2026, 10))

        self.assertEqual(start["action"], "offer")

    def test_editor_uses_server_data_when_draft_matches_it(self):
        official = [row("Migliorato", "2026-10-07", "Ferie")]
        sig = ustore.month_signature(official, "Migliorato", 2026, 10)
        draft = ud.set_month(
            ud.empty_draft("Migliorato"), "2026-10",
            entries=[(dt.date(2026, 10, 7), "Ferie", "")], base_signature=[],
            updated_at="t", session_id="s1",
        )

        self.assertEqual(ud.editor_start(draft, "2026-10", sig)["action"], "official")
        self.assertEqual(ud.editor_start(ud.empty_draft("Migliorato"), "2026-10", sig)["action"], "official")

    def test_pending_summary_lists_only_months_that_differ_from_server(self):
        official = [row("Migliorato", "2026-10-04", "Tutto il giorno"), row("Migliorato", "2026-11-02", "Notte")]
        draft = ud.empty_draft("Migliorato")
        draft = ud.set_month(
            draft, "2026-10",
            entries=[(dt.date(2026, 10, 7), "Ferie", ""), (dt.date(2026, 10, 8), "Ferie", "")],
            base_signature=[], updated_at="2026-09-29T11:55:00Z", session_id="s1",
        )
        draft = ud.set_month(
            draft, "2026-11",
            entries=[(dt.date(2026, 11, 2), "Notte", "")],
            base_signature=[], updated_at="2026-09-29T11:56:00Z", session_id="s1",
        )

        pending = ud.pending_summary(draft, official, ["2026-10", "2026-11"])

        self.assertEqual(len(pending), 1)
        self.assertEqual(pending[0]["month"], "2026-10")
        self.assertEqual(pending[0]["added"], 2)
        self.assertEqual(pending[0]["removed"], 1)

    def test_remove_month_drops_only_that_month(self):
        draft = ud.set_month(
            ud.empty_draft("Migliorato"), "2026-10",
            entries=[(dt.date(2026, 10, 7), "Ferie", "")], base_signature=[],
            updated_at="t", session_id="s1",
        )

        self.assertEqual(ud.remove_month(draft, "2026-10")["months"], {})


class FakeDraftServer:
    """Stand-in for the GitHub draft file: read-modify-write, can be made to fail."""

    def __init__(self, drafts=None):
        self.drafts = drafts or {}
        self.writes = 0
        self.fail = False
        self.loads = 0

    def load(self, doctor):
        self.loads += 1
        if self.fail:
            raise RuntimeError("GitHub saturo")
        return self.drafts.get(doctor) or ud.empty_draft(doctor)

    def update(self, doctor, apply_fn):
        if self.fail:
            raise RuntimeError("GitHub saturo")
        self.drafts[doctor] = apply_fn(self.drafts.get(doctor) or ud.empty_draft(doctor))
        self.writes += 1


class MemoryDraftStoreTests(unittest.TestCase):
    def setUp(self):
        self.t = [1000.0]
        self.server = FakeDraftServer()
        self.store = ud.MemoryDraftStore(
            load_fn=self.server.load,
            persist_fn=self.server.update,
            clock=lambda: self.t[0],
            persist_min_interval=120,
        )

    def put(self, mk, days, doctor="Migliorato"):
        self.store.set_month(
            doctor, mk,
            entries=[(dt.date(int(mk[:4]), int(mk[5:7]), d), "Ferie", "") for d in days],
            base_signature=[], updated_at="t", session_id="s",
        )

    def test_edits_are_in_memory_at_once_and_persisted_with_throttle(self):
        self.put("2026-10", [7])
        self.store.flush("Migliorato")
        self.assertEqual(self.server.writes, 1)          # prima copia subito

        self.put("2026-10", [7, 8])
        self.store.flush("Migliorato")
        self.assertEqual(self.server.writes, 1)          # entro 2 minuti: solo memoria
        self.assertEqual(len(ud.month_entries(self.store.get("Migliorato"), "2026-10")), 2)

        self.t[0] += 121
        self.store.flush("Migliorato")
        self.assertEqual(self.server.writes, 2)
        self.assertEqual(len(ud.month_entries(self.server.drafts["Migliorato"], "2026-10")), 2)

    def test_failed_persist_is_retried_later(self):
        self.server.fail = True
        self.put("2026-10", [7])
        self.store.flush("Migliorato", force=True)
        self.assertEqual(self.server.writes, 0)

        self.server.fail = False
        self.store.flush("Migliorato", force=True)
        self.assertEqual(self.server.writes, 1)

    def test_persist_merges_with_months_already_on_server(self):
        self.server.drafts["Migliorato"] = ud.set_month(
            ud.empty_draft("Migliorato"), "2026-11",
            entries=[(dt.date(2026, 11, 3), "Notte", "")], base_signature=[], updated_at="t", session_id="old",
        )
        self.server.fail = True           # all'avvio GitHub non risponde: memoria vuota
        self.put("2026-10", [7])
        self.server.fail = False
        self.store.flush("Migliorato", force=True)

        self.assertEqual(sorted(self.server.drafts["Migliorato"]["months"]), ["2026-10", "2026-11"])

    def test_removed_month_is_removed_on_server_too(self):
        self.put("2026-10", [7])
        self.store.flush("Migliorato", force=True)
        self.store.remove_month("Migliorato", "2026-10")
        self.store.flush("Migliorato", force=True)

        self.assertEqual(self.server.drafts["Migliorato"]["months"], {})

    def test_server_draft_is_loaded_once(self):
        self.store.get("Migliorato")
        self.store.get("Migliorato")
        self.assertEqual(self.server.loads, 1)

    def test_all_drafts_lists_doctors_in_memory(self):
        self.put("2026-10", [7], doctor="Migliorato")
        self.put("2026-10", [8], doctor="Crea")
        self.assertEqual(sorted(self.store.all_drafts()), ["Crea", "Migliorato"])


if __name__ == "__main__":
    unittest.main()
