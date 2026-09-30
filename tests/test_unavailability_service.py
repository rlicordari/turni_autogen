import csv
import datetime as dt
import io
import tempfile
import threading
import time
import unittest
from pathlib import Path
from unittest import mock

import github_utils
import smtplib
import unavailability_service as svc
import unavailability_store as ustore
from tests.fakes import FakeClock, FakeGithubHTTP, FakeSMTP

MONTH = (2026, 10)


def cfg():
    return svc.ServiceConfig(
        github={"owner": "o", "repo": "r", "token": "t", "branch": "main"},
        smtp={"host": "smtp.test", "port": 587, "from": "turni@example.org", "starttls": False},
        receipt_cc=["admin@example.org", "utic@polime.it"],
        app_build="test",
    )


def rows_of(gh, doctor):
    text = gh.text(f"data/unavailability/unavail_{svc.doctor_slug(doctor)}.csv")
    return list(csv.DictReader(io.StringIO(text))) if text.strip() else []


def audit_rows(gh, mk="2026-10"):
    text = gh.text(f"data/unavailability_audit/unavailability_audit_{mk}.csv")
    return list(csv.DictReader(io.StringIO(text))) if text.strip() else []


def job(doctor, days, base=(), shift="Ferie", email=None, origin="sessione-1"):
    return svc.make_job(
        doctor=doctor,
        doctor_email=email or f"{svc.doctor_slug(doctor)}@example.org",
        entries_by_month={MONTH: [(dt.date(2026, 10, d), shift, "") for d in days]},
        base_signatures={MONTH: list(base)},
        origin=origin,
    )


class ServiceTestCase(unittest.TestCase):
    chaos = 0.0

    def setUp(self):
        self.gh = FakeGithubHTTP(chaos=self.chaos)
        self.clock = FakeClock()
        FakeSMTP.sent = []
        github_utils.reset_state()
        for target, attr, value in (
            (github_utils, "_http", self.gh),
            (github_utils, "_now", self.clock.now),
            (github_utils, "_sleep", self.clock.sleep),
            (github_utils, "WRITE_MIN_INTERVAL", 1.0),
            (smtplib, "SMTP", FakeSMTP),
        ):
            p = mock.patch.object(target, attr, value)
            p.start()
            self.addCleanup(p.stop)
        self.addCleanup(github_utils.reset_state)


class ConcurrencyUnderChaosTests(ServiceTestCase):
    chaos = 0.3  # 30% delle richieste: rate limit, 502, timeout o reset di rete

    def test_many_doctors_saving_at_once_all_land_with_audit_and_receipt(self):
        doctors = [f"Medico {i:02d}" for i in range(25)]
        errors = []
        barrier = threading.Barrier(len(doctors))

        def run(i, doctor):
            barrier.wait()
            try:
                result = svc.save_now(cfg(), job(doctor, [1 + i % 28, 2 + i % 27]), max_wait=10_000)
                svc.write_audit(cfg(), result, max_wait=10_000)
                svc.send_receipt(cfg(), result)
            except Exception as e:  # pragma: no cover - reported below
                errors.append((doctor, repr(e)))

        threads = [threading.Thread(target=run, args=(i, d)) for i, d in enumerate(doctors)]
        for t in threads:
            t.start()
        for t in threads:
            t.join()

        self.assertEqual(errors, [])
        for i, doctor in enumerate(doctors):
            self.assertEqual(
                sorted(r["date"] for r in rows_of(self.gh, doctor)),
                sorted({f"2026-10-{1 + i % 28:02d}", f"2026-10-{2 + i % 27:02d}"}),
            )
        # Un'unica riga di audit per medico nel file condiviso del mese, nessuna persa.
        self.assertEqual(sorted(r["doctor"] for r in audit_rows(self.gh)), sorted(doctors))
        self.assertEqual(len(FakeSMTP.sent), len(doctors))
        # Le scritture del processo non si sono mai sovrapposte sul branch.
        self.assertEqual(self.gh.max_put_in_flight, 1)

    def test_same_doctor_two_devices_same_month_never_silently_overwrites(self):
        outcomes = []
        barrier = threading.Barrier(2)

        def run(days):
            barrier.wait()
            try:
                svc.save_now(cfg(), job("Migliorato", days), max_wait=10_000)
                outcomes.append(("ok", days))
            except ustore.MonthConflictError:
                outcomes.append(("conflict", days))

        threads = [threading.Thread(target=run, args=(d,)) for d in ([7, 8], [20])]
        for t in threads:
            t.start()
        for t in threads:
            t.join()

        self.assertEqual(sorted(o[0] for o in outcomes), ["conflict", "ok"])
        winner = next(days for status, days in outcomes if status == "ok")
        self.assertEqual(sorted(int(r["date"][-2:]) for r in rows_of(self.gh, "Migliorato")), winner)

    def test_same_doctor_two_devices_different_months_both_land(self):
        nov = svc.make_job(
            doctor="Migliorato",
            doctor_email="m@example.org",
            entries_by_month={(2026, 11): [(dt.date(2026, 11, 3), "Notte", "")]},
            base_signatures={(2026, 11): []},
        )
        threads = [
            threading.Thread(target=svc.save_now, args=(cfg(), job("Migliorato", [7])), kwargs={"max_wait": 10_000}),
            threading.Thread(target=svc.save_now, args=(cfg(), nov), kwargs={"max_wait": 10_000}),
        ]
        for t in threads:
            t.start()
        for t in threads:
            t.join()

        self.assertEqual(sorted(r["date"] for r in rows_of(self.gh, "Migliorato")), ["2026-10-07", "2026-11-03"])


class QueueTests(ServiceTestCase):
    def make_queue(self, **kw):
        return svc.SaveQueue(cfg(), clock=self.clock.now, **kw)

    def test_save_is_queued_while_github_is_down_and_completed_when_back(self):
        self.gh.down = True
        q = self.make_queue()
        j = job("Migliorato", [7, 8, 9, 10, 11, 12])

        with self.assertRaises(github_utils.GithubUnavailable):
            svc.save_now(cfg(), j, max_wait=25)
        q.enqueue(j)
        q.process_due()                      # ancora giù: resta in coda
        self.assertEqual(len(q.pending("Migliorato")), 1)
        self.assertEqual(rows_of(self.gh, "Migliorato"), [])

        self.gh.down = False
        self.clock.sleep(3600)
        q.process_due()

        self.assertEqual(q.pending("Migliorato"), [])
        self.assertEqual(len(rows_of(self.gh, "Migliorato")), 6)
        self.assertEqual(len(audit_rows(self.gh)), 1)
        self.assertEqual(len(FakeSMTP.sent), 1)
        self.assertIn("mer 07/10", FakeSMTP.sent[0].get_content())
        results = q.take_results("Migliorato")
        self.assertEqual(results[0]["status"], "done")
        self.assertEqual(q.take_results("Migliorato"), [])

    def test_background_conflict_is_never_applied_and_doctor_is_notified(self):
        q = self.make_queue()
        j = job("Migliorato", [7], base=[])
        q.enqueue(j)
        # Nel frattempo un altro dispositivo ha salvato ottobre.
        svc.save_now(cfg(), job("Migliorato", [20]), max_wait=10)

        q.process_due()

        self.assertEqual([r["date"] for r in rows_of(self.gh, "Migliorato")], ["2026-10-20"])
        self.assertEqual(q.take_results("Migliorato")[0]["status"], "conflict")
        mail = FakeSMTP.sent[-1]
        self.assertIn("NON", mail["Subject"])
        self.assertEqual(mail["To"], "migliorato@example.org")

    def test_repeated_saves_of_same_doctor_are_merged_in_queue(self):
        q = self.make_queue()
        q.enqueue(job("Migliorato", [7]))
        q.enqueue(job("Migliorato", [7, 8]))

        pending = q.pending("Migliorato")
        self.assertEqual(len(pending), 1)
        self.assertEqual(len(pending[0].entries_by_month[MONTH]), 2)

    def test_queue_survives_process_restart(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "queue.json"
            self.make_queue(persist_path=path).enqueue(job("Migliorato", [7, 8]))

            reloaded = self.make_queue(persist_path=path)

            self.assertEqual(len(reloaded.pending("Migliorato")), 1)
            reloaded.process_due()
            self.assertEqual(len(rows_of(self.gh, "Migliorato")), 2)


class SaveOrEnqueueTests(ServiceTestCase):
    def make_queue(self):
        return svc.SaveQueue(cfg(), clock=self.clock.now)

    def test_interactive_save_supersedes_queued_one_without_false_conflict(self):
        q = self.make_queue()
        self.gh.down = True
        status, _ = q.save_or_enqueue(job("Migliorato", [7]), max_wait=5)
        self.assertEqual(status, "queued")

        self.gh.down = False
        self.clock.sleep(61)  # GitHub aveva chiesto di attendere 60 s
        status, result = q.save_or_enqueue(job("Migliorato", [7, 8]), max_wait=5)
        q.process_due()

        self.assertEqual(status, "saved")
        self.assertEqual(q.pending(), [])
        self.assertEqual(sorted(r["date"] for r in rows_of(self.gh, "Migliorato")), ["2026-10-07", "2026-10-08"])
        self.assertFalse(any("NON" in str(m["Subject"]) for m in FakeSMTP.sent))

    def test_queued_months_are_saved_together_with_new_interactive_save(self):
        q = self.make_queue()
        self.gh.down = True
        nov = svc.make_job(
            doctor="Migliorato", doctor_email="m@example.org",
            entries_by_month={(2026, 11): [(dt.date(2026, 11, 3), "Notte", "")]},
            base_signatures={(2026, 11): []},
            origin="sessione-1",
        )
        q.save_or_enqueue(nov, max_wait=5)

        self.gh.down = False
        self.clock.sleep(61)
        status, result = q.save_or_enqueue(job("Migliorato", [7]), max_wait=5)

        self.assertEqual(status, "saved")
        self.assertEqual(sorted(r["date"] for r in rows_of(self.gh, "Migliorato")), ["2026-10-07", "2026-11-03"])
        self.assertEqual(sorted(result.outcome.diffs), ["2026-10", "2026-11"])

    def test_other_device_cannot_overwrite_a_month_waiting_in_queue(self):
        q = self.make_queue()
        self.gh.down = True
        q.save_or_enqueue(job("Migliorato", [7], origin="telefono"), max_wait=5)

        with self.assertRaises(ustore.MonthConflictError):
            q.save_or_enqueue(job("Migliorato", [20], origin="pc"), max_wait=5)

        pending = q.pending("Migliorato")[0]
        self.assertEqual([d.day for d, _s, _n in pending.entries_by_month[MONTH]], [7])

    def test_other_device_cannot_overwrite_a_month_just_landed_from_queue(self):
        q = self.make_queue()
        self.gh.down = True
        q.save_or_enqueue(job("Migliorato", [7], origin="telefono"), max_wait=5)
        self.gh.down = False
        self.clock.sleep(3600)
        q.process_due()                      # il telefono ha registrato il 7

        with self.assertRaises(ustore.MonthConflictError):
            q.save_or_enqueue(job("Migliorato", [20], origin="pc"), max_wait=5)
        self.assertEqual([r["date"] for r in rows_of(self.gh, "Migliorato")], ["2026-10-07"])

    def test_worker_and_interactive_save_of_same_doctor_are_serialized(self):
        q = self.make_queue()
        self.gh.down = True
        q.save_or_enqueue(job("Migliorato", [7]), max_wait=5)
        self.gh.down = False
        self.clock.sleep(3600)
        barrier = threading.Barrier(2)
        statuses = []

        def worker():
            barrier.wait()
            q.process_due()

        def interactive():
            barrier.wait()
            statuses.append(q.save_or_enqueue(job("Migliorato", [7, 9]), max_wait=5)[0])

        threads = [threading.Thread(target=worker), threading.Thread(target=interactive)]
        for t in threads:
            t.start()
        for t in threads:
            t.join()
        q.process_due()

        self.assertEqual(statuses, ["saved"])
        self.assertEqual(q.pending(), [])
        self.assertEqual(sorted(r["date"] for r in rows_of(self.gh, "Migliorato")), ["2026-10-07", "2026-10-09"])
        self.assertFalse(any("NON" in str(m["Subject"]) for m in FakeSMTP.sent))


class DraftIOTests(ServiceTestCase):
    chaos = 0.3

    def test_concurrent_draft_updates_of_different_months_are_all_kept(self):
        import unavailability_drafts as ud
        barrier = threading.Barrier(8)

        def write(day):
            barrier.wait()
            svc.update_draft(
                cfg(), "Migliorato",
                lambda d: ud.set_month(
                    d, f"2026-{day:02d}",
                    entries=[(dt.date(2026, day, 1), "Ferie", "")],
                    base_signature=[], updated_at="t", session_id="s",
                ),
                max_wait=10_000,
            )

        threads = [threading.Thread(target=write, args=(m,)) for m in range(1, 9)]
        for t in threads:
            t.start()
        for t in threads:
            t.join()

        draft = svc.load_draft(cfg(), "Migliorato", max_wait=10_000)
        self.assertEqual(sorted(draft["months"]), [f"2026-{m:02d}" for m in range(1, 9)])
        self.assertEqual(draft["doctor"], "Migliorato")


class WorkerThreadTests(unittest.TestCase):
    """Real background thread, real (short) time."""

    def test_worker_completes_queued_save(self):
        gh = FakeGithubHTTP()
        FakeSMTP.sent = []
        github_utils.reset_state()
        with mock.patch.object(github_utils, "_http", gh), \
                mock.patch.object(github_utils, "WRITE_MIN_INTERVAL", 0.0), \
                mock.patch.object(smtplib, "SMTP", FakeSMTP):
            q = svc.SaveQueue(cfg())
            q.enqueue(job("Migliorato", [7]))
            q.start_worker(interval=0.05)
            deadline = time.time() + 10
            while q.pending() and time.time() < deadline:
                time.sleep(0.05)
            q.stop_worker()

        self.assertEqual(q.pending(), [])
        self.assertEqual([r["date"] for r in rows_of(gh, "Migliorato")], ["2026-10-07"])
        self.assertEqual(len(FakeSMTP.sent), 1)


if __name__ == "__main__":
    unittest.main()
