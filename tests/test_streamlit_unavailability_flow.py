"""End-to-end test of the doctor unavailability page (real streamlit_app.py).

GitHub and SMTP are replaced by in-memory fakes; everything else is the real app.
"""

import base64
import csv
import datetime as dt
import hashlib
import io
import json
import smtplib
import unittest
from pathlib import Path
from unittest import mock

import os
import threading
import time

import streamlit as st
from streamlit.testing.v1 import AppTest

import github_utils
import unavailability_service as usvc
from tests.fakes import FakeClock, FakeGithubHTTP, FakeSMTP

APP = str(Path(__file__).resolve().parents[1] / "streamlit_app.py")
DOCTOR = "Migliorato"
PIN = "1234"
# La pagina propone il mese successivo a oggi.
_NEXT = (dt.date.today().replace(day=1) + dt.timedelta(days=32)).replace(day=1)
YY, MM = _NEXT.year, _NEXT.month
MK = f"{YY:04d}-{MM:02d}"
ROWS_KEY = f"unav_rows_{DOCTOR}_{YY}_{MM}"
OFFICIAL = "data/unavailability/unavail_migliorato.csv"
DRAFT = "data/unavailability_drafts/draft_migliorato.json"


def day(n: int) -> dt.date:
    return dt.date(YY, MM, n)


def official_rows(gh):
    text = gh.files.get(OFFICIAL, ("", ""))[0]
    return list(csv.DictReader(io.StringIO(text))) if text.strip() else []


class DoctorUnavailabilityFlowTests(unittest.TestCase):
    chaos = 0.0

    def setUp(self):
        # Risorse condivise dell'app (lease, bozze, coda) ripartono da zero a ogni test.
        usvc.stop_all_workers()
        st.cache_data.clear()
        st.cache_resource.clear()
        github_utils.reset_state()
        self.addCleanup(usvc.stop_all_workers)
        self.addCleanup(github_utils.reset_state)
        self.clock = FakeClock()
        self.gh = FakeGithubHTTP(chaos=self.chaos)
        salt = b"0123456789abcdef"
        self.gh.seed(
            f"data/doctor_auth/pin_migliorato.json",
            json.dumps({
                "doctor": DOCTOR,
                "salt_b64": base64.b64encode(salt).decode(),
                "hash_b64": base64.b64encode(hashlib.pbkdf2_hmac("sha256", PIN.encode(), salt, 1000)).decode(),
                "iters": 1000,
            }),
        )
        self.gh.seed("data/doctor_contacts.yml", f"{DOCTOR}:\n  email: migliorato@example.org\n")
        self.gh.seed(
            "data/unavailability_settings.yml",
            "unavailability_open: true\nmax_unavailability_per_shift: 6\nmax_weekend_days: 4\n"
            "receipt_cc_emails:\n- admin@example.org\n- utic@polime.it\n",
        )
        FakeSMTP.sent = []
        patches = [
            mock.patch.object(github_utils, "_http", self.gh),
            mock.patch.object(github_utils, "_now", self.clock.now),
            mock.patch.object(github_utils, "_sleep", self.clock.sleep),
            mock.patch.object(smtplib, "SMTP", FakeSMTP),
            mock.patch.dict(os.environ, {"TURNI_QUEUE_WORKER_INTERVAL": "0.05"}),
        ]
        for p in patches:
            p.start()
            self.addCleanup(p.stop)

    def new_session(self) -> AppTest:
        at = AppTest.from_file(APP, default_timeout=60)
        at.secrets["auth"] = {"admin_pin": "9999"}
        at.secrets["github_unavailability"] = {"token": "t", "owner": "o", "repo": "r", "branch": "main"}
        at.secrets["smtp"] = {"host": "smtp.test", "port": 587, "from": "turni@example.org", "starttls": False}
        at.run()
        self.assertFalse(at.exception, at.exception)
        return at

    def login(self, at: AppTest) -> AppTest:
        at.selectbox(key="login_doctor").set_value(DOCTOR)
        at.run()
        at.text_input(key="login_pin").input(PIN)
        next(b for b in at.button if b.label == "Accedi").click()
        at.run()
        self.assertFalse(at.exception, at.exception)
        return at

    def add_ferie(self, at: AppTest, start: dt.date, end: dt.date) -> None:
        at.date_input(key=f"{ROWS_KEY}__ferie_start").set_value(start)
        at.date_input(key=f"{ROWS_KEY}__ferie_end").set_value(end)
        at.button(key=f"{ROWS_KEY}__ferie_add").click()
        at.run()
        self.assertFalse(at.exception, at.exception)

    def save(self, at: AppTest) -> None:
        at.button(key=f"save_unav_{YY}_{MM}").click()
        at.run()
        self.assertFalse(at.exception, at.exception)

    def page_text(self, at: AppTest) -> str:
        parts = [e.value for e in list(at.success) + list(at.error) + list(at.warning) + list(at.info)]
        return "\n".join(str(p) for p in parts)

    def test_unsent_period_is_kept_in_draft_and_save_sends_receipt(self):
        at = self.login(self.new_session())
        self.add_ferie(at, day(7), day(12))

        # Non ancora inviato: l'archivio ufficiale e' vuoto, ma la bozza ha i 6 giorni.
        self.assertEqual(official_rows(self.gh), [])
        draft = json.loads(self.gh.files[DRAFT][0])
        self.assertEqual(len(draft["months"][MK]["entries"]), 6)
        self.assertIn("Modifiche NON inviate", self.page_text(at))

        self.save(at)

        self.assertEqual(
            sorted(r["date"] for r in official_rows(self.gh)),
            [day(d).isoformat() for d in range(7, 13)],
        )
        self.assertNotIn(MK, json.loads(self.gh.files[DRAFT][0])["months"])
        self.assertEqual(len(FakeSMTP.sent), 1)
        mail = FakeSMTP.sent[0]
        self.assertEqual(mail["To"], "migliorato@example.org")
        self.assertEqual(mail["Cc"], "admin@example.org, utic@polime.it")
        self.assertIn(f"{day(7):%d/%m}  Ferie", mail.get_content())
        self.assertIn("verificato sul server", self.page_text(at))

        # Secondo tap: nessuna nuova scrittura e nessuna seconda mail.
        official_sha_before = self.gh.files[OFFICIAL][1]
        self.save(at)
        self.assertEqual(self.gh.files[OFFICIAL][1], official_sha_before)
        self.assertEqual(len(FakeSMTP.sent), 1)
        self.assertIn("Nessuna modifica da inviare", self.page_text(at))

    def test_unsent_draft_is_restored_at_next_login(self):
        at = self.login(self.new_session())
        self.add_ferie(at, day(7), day(12))
        # Il medico chiude la pagina senza salvare; riapre da un altro dispositivo.
        other = self.login(self.new_session())

        restored = other.session_state[ROWS_KEY]
        self.assertEqual(len(restored), 6)
        self.assertIn("NON inviate", self.page_text(other))
        self.assertEqual(official_rows(self.gh), [])

    def test_stale_tab_cannot_erase_days_saved_from_another_device(self):
        pc = self.login(self.new_session())          # scheda aperta sul PC, ottobre vuoto
        phone = self.login(self.new_session())       # poi il telefono salva 7-12 ottobre
        self.add_ferie(phone, day(7), day(12))
        self.save(phone)
        self.assertEqual(len(official_rows(self.gh)), 6)
        # Il medico esce dal telefono: il PC (scheda vecchia) non viene buttato fuori.
        next(b for b in phone.button if b.label == "Esci / cambia medico").click()
        phone.run()

        self.add_ferie(pc, day(20), day(20))
        self.save(pc)

        self.assertEqual(len(official_rows(self.gh)), 6)
        self.assertIn("BLOCCATO", self.page_text(pc))

    def wait_until(self, cond, timeout=15.0):
        deadline = time.time() + timeout
        while time.time() < deadline:
            if cond():
                return True
            time.sleep(0.05)
        return False

    def test_save_is_queued_when_github_is_saturated_and_completed_later(self):
        at = self.login(self.new_session())
        self.add_ferie(at, day(7), day(12))

        self.gh.down = True                      # GitHub risponde solo "rate limit"
        self.save(at)

        self.assertIn("IN CODA", self.page_text(at))
        self.assertEqual(official_rows(self.gh), [])
        self.assertEqual(FakeSMTP.sent, [])

        self.gh.down = False                     # GitHub torna disponibile
        self.clock.sleep(3600)
        self.assertTrue(self.wait_until(lambda: len(official_rows(self.gh)) == 6), "il worker non ha completato")
        self.assertTrue(self.wait_until(lambda: len(FakeSMTP.sent) == 1))
        self.assertIn(f"{day(7):%d/%m}  Ferie", FakeSMTP.sent[0].get_content())

        at.run()                                 # il medico ricarica la pagina
        self.assertIn("in coda completato", self.page_text(at))
        self.assertNotIn("IN CODA", self.page_text(at))

    DOCTORS = ["Migliorato", "Crea", "Cimino", "Allegra", "Licordari", "Trio", "Pugliatti", "Colarusso"]

    def seed_doctor(self, doctor):
        salt = b"0123456789abcdef"
        slug = usvc.doctor_slug(doctor)
        self.gh.seed(
            f"data/doctor_auth/pin_{slug}.json",
            json.dumps({
                "doctor": doctor,
                "salt_b64": base64.b64encode(salt).decode(),
                "hash_b64": base64.b64encode(hashlib.pbkdf2_hmac("sha256", PIN.encode(), salt, 1000)).decode(),
                "iters": 1000,
            }),
        )

    def test_saves_of_many_open_sessions_all_land(self):
        contacts = "".join(f"{d}:\n  email: {usvc.doctor_slug(d)}@example.org\n" for d in self.DOCTORS)
        self.gh.seed("data/doctor_contacts.yml", contacts)
        for doctor in self.DOCTORS:
            self.seed_doctor(doctor)
        sessions = {}
        for i, doctor in enumerate(self.DOCTORS):
            at = self.new_session()
            at.selectbox(key="login_doctor").set_value(doctor)
            at.run()
            at.text_input(key="login_pin").input(PIN)
            next(b for b in at.button if b.label == "Accedi").click()
            at.run()
            rows_key = f"unav_rows_{doctor}_{YY}_{MM}"
            at.date_input(key=f"{rows_key}__ferie_start").set_value(day(1 + i))
            at.date_input(key=f"{rows_key}__ferie_end").set_value(day(3 + i))
            at.button(key=f"{rows_key}__ferie_add").click()
            at.run()
            self.assertFalse(at.exception, at.exception)
            sessions[doctor] = at

        # AppTest non regge sessioni in thread paralleli: qui le 8 sessioni restano aperte
        # e salvano una dopo l'altra (la concorrenza vera tra thread è coperta dai test
        # del servizio: ConcurrencyUnderChaosTests + stress).
        errors = []
        for doctor, at in sessions.items():
            at.button(key=f"save_unav_{YY}_{MM}").click()
            at.run()
            if at.exception:
                errors.append((doctor, at.exception))

        self.assertEqual(errors, [])
        for i, doctor in enumerate(self.DOCTORS):
            text = self.gh.text(f"data/unavailability/unavail_{usvc.doctor_slug(doctor)}.csv")
            dates = sorted(r["date"] for r in csv.DictReader(io.StringIO(text))) if text.strip() else []
            self.assertEqual(dates, [day(d).isoformat() for d in range(1 + i, 4 + i)], doctor)
            self.assertIn("verificato sul server", self.page_text(sessions[doctor]), doctor)
        audit = self.gh.text(f"data/unavailability_audit/unavailability_audit_{MK}.csv")
        self.assertEqual(sorted(r["doctor"] for r in csv.DictReader(io.StringIO(audit))), sorted(self.DOCTORS))
        self.assertEqual(len(FakeSMTP.sent), len(self.DOCTORS))
        self.assertEqual(self.gh.max_put_in_flight, 1)

    def admin_session(self, section: str) -> AppTest:
        at = self.new_session()
        at.sidebar.radio[0].set_value(section)
        at.run()
        at.text_input[0].input("9999")
        next(b for b in at.button if b.label == "Sblocca area Admin").click()
        at.run()
        self.assertFalse(at.exception, at.exception)
        return at

    def test_admin_sees_unsent_drafts_before_generating(self):
        doc = self.login(self.new_session())
        self.add_ferie(doc, day(7), day(12))   # inserite ma mai inviate

        admin = self.admin_session("⚙️ Admin — Genera turni")
        admin.date_input(key="custom_period_start_draft").set_value(day(1))
        admin.date_input(key="custom_period_end_draft").set_value(day(28))
        admin.radio(key="admin_period_mode_draft").set_value("Periodo personalizzato")
        next(b for b in admin.button if b.label == "Applica periodo").click()
        admin.run()

        self.assertFalse(admin.exception, admin.exception)
        self.assertIn("bozze NON inviate", self.page_text(admin))
        table = admin.dataframe[0].value
        self.assertEqual(list(table["Medico"]), [DOCTOR])
        self.assertEqual(list(table["Da aggiungere"]), [6])

    def test_admin_configuration_shows_receipt_copies(self):
        admin = self.admin_session("🔧 Admin — Configurazione")

        self.assertEqual(
            admin.text_area(key="receipt_cc_emails_input").value,
            "admin@example.org\nutic@polime.it",
        )



class DoctorUnavailabilityFlowUnderChaosTests(DoctorUnavailabilityFlowTests):
    """Same flows with 20% of GitHub requests failing (rate limits, 5xx, timeouts)."""

    chaos = 0.2


if __name__ == "__main__":
    unittest.main()
