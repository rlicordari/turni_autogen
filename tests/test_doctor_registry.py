import datetime as dt
import unittest
from pathlib import Path

import doctor_registry as dr
import turni_generator as tg
from turni_generator import DayRow

YAML = tg.load_rules(Path("Regole_Turni.yml"))


def card(**kw):
    c = dr.empty_card()
    c.update(kw)
    return c


class CardIsolationTests(unittest.TestCase):
    def test_saving_one_doctor_never_changes_another(self):
        # 18/11/2026: le colonne scelte per D'Angelo finirono su Allegra.
        reg = dr.set_doctor(dr.empty_registry(), "Allegra",
                            card(services={"J": {"level": "pool"}, "V": {"level": "pool"}}))
        before = dr.normalize(reg)["doctors"]["Allegra"]

        reg = dr.set_doctor(reg, "D'Angelo", card(services={"Y": {"level": "pool"}, "EG": {"level": "pool"}}))

        self.assertEqual(dr.normalize(reg)["doctors"]["Allegra"], before)
        self.assertEqual(sorted(reg["doctors"]["D'Angelo"]["services"]), ["EG", "Y"])

    def test_round_trip_and_cleaning(self):
        reg = dr.set_doctor(dr.empty_registry(), "Zito", card(
            services={"J": {"level": "pool", "quota": {"type": "fixed", "value": 2}, "weekend_max": 1},
                      "XX": {"level": "pool"}, "Q": {"level": "boh"}},
        ))
        self.assertEqual(dr.from_text(dr.to_text(reg))["doctors"], reg["doctors"])
        self.assertEqual(list(reg["doctors"]["Zito"]["services"]), ["J"])   # servizi ignoti scartati
        never = dr.set_doctor(reg, "Zito", card(mai_notti=True, services={"J": {"level": "pool"}}))
        self.assertNotIn("J", never["doctors"]["Zito"]["services"])

    def test_invalid_name_rejected(self):
        with self.assertRaises(ValueError):
            dr.set_doctor(dr.empty_registry(), "<script>", card())


class MigrationTests(unittest.TestCase):
    def setUp(self):
        self.reg = dr.migrate(None, YAML)
        self.docs = self.reg["doctors"]

    def svc(self, doc, sid):
        return self.docs[doc]["services"].get(sid) or {}

    def test_every_hidden_per_person_rule_lands_in_the_card(self):
        self.assertEqual(self.svc("Crea", "V").get("doppio"), "titolare")
        self.assertTrue(self.svc("Crea", "AB").get("sabato"))
        self.assertEqual(self.svc("Pugliatti", "W").get("giorni_fissi"), ["Tue"])
        self.assertEqual(self.svc("Migliorato", "AC").get("level"), "pool")
        self.assertEqual(self.svc("Cimino", "U").get("quota"), {"type": "fixed", "value": 2})
        self.assertEqual(self.svc("De Gregorio", "I").get("quota"), {"type": "max", "value": 3})
        self.assertEqual(self.svc("Zito", "J").get("weekend_min"), 1)
        self.assertEqual(self.svc("Zito", "J").get("weekend_max"), 1)
        self.assertEqual(self.svc("Zito", "J").get("quota"), {"type": "fixed", "value": 2})
        self.assertFalse(self.svc("Calabrò", "J").get("notti_weekend"))
        self.assertEqual(self.docs["Manganaro"]["festivi_quota"], 2)
        self.assertTrue(self.docs["De Gregorio"]["mai_notti"])
        self.assertTrue(self.docs["Manganaro"]["mai_notti"])
        self.assertEqual(self.docs["Allegra"]["lunedi_V_U_con"], ["Crea", "Dattilo"])
        self.assertTrue(self.docs["Grimaldi"]["esente_weekend_liberi"])
        self.assertEqual(self.svc("Recupero", "T").get("quota"), {"type": "min", "value": 4, "target": 5})
        self.assertEqual(self.svc("Recupero", "W").get("quota"), {"type": "max", "value": 8})
        self.assertEqual(self.svc("Recupero", "L").get("quota"), {"type": "fixed", "value": 3})
        self.assertIsNotNone(self.docs["Dattilo"]["universitario"])

    def test_df_pair_is_d_pool_and_never_reserve_for_other_wards(self):
        self.assertEqual(dr.doctors_for(self.reg, "D"), ["Calabrò", "Grimaldi"])
        self.assertNotIn("H", self.docs["Grimaldi"]["services"])
        self.assertNotIn("EG", self.docs["Calabrò"]["services"])

    def test_admin_active_flag_wins_over_old_yaml_exclusion(self):
        v1 = {"doctors": {"D'Angelo": {"active": True, "columns": []}}}
        reg = dr.migrate(v1, YAML)
        self.assertTrue(reg["doctors"]["D'Angelo"]["active"])
        self.assertIn("D'Angelo: attivo ma solo di riserva", " ".join(dr.check(reg)))


class SolverConfigTests(unittest.TestCase):
    def test_no_doctor_can_come_from_outside_the_cards(self):
        everyone = set(dr.migrate(None, YAML)["doctors"])
        reg = dr.set_doctor(dr.empty_registry(), "Nuovo", card(services={"J": {"level": "pool"}}))

        cfg = dr.to_solver_cfg(YAML, reg)

        leaked = (dr.strings_in(cfg) & everyone) - {"Recupero"}
        self.assertEqual(leaked, set(), "nomi presi dal file regole o dal codice")
        self.assertEqual(tg.collect_doctors(cfg), ["Nuovo"])

    def test_inactive_doctor_is_never_a_candidate(self):
        reg = dr.migrate(None, YAML)
        reg = dr.set_doctor(reg, "Trio", {**reg["doctors"]["Trio"], "active": False})
        cfg = dr.to_solver_cfg(YAML, reg)
        days = [DayRow(dt.date(2026, 12, 1) + dt.timedelta(days=i), d, 2 + i)
                for i, d in enumerate(["Tue", "Wed", "Thu", "Fri", "Sat", "Sun", "Mon"])]

        slots = tg.slots_for_month(cfg, days, {})

        self.assertFalse(any("Trio" in s.allowed for s in slots))

    def test_reserve_only_when_nobody_of_the_pool_is_available(self):
        reg = dr.empty_registry()
        reg = dr.set_doctor(reg, "Primo", card(services={"H": {"level": "pool"}}))
        reg = dr.set_doctor(reg, "Riserva", card(services={"H": {"level": "riserva"}}))
        reg = dr.set_doctor(reg, "Estraneo", card(services={"K": {"level": "pool"}}))
        cfg = dr.to_solver_cfg(YAML, reg)
        day = DayRow(dt.date(2026, 12, 2), "Wed", 2)

        h = next(s for s in tg.slots_for_month(cfg, [day], {}) if s.columns == ["H"])
        self.assertEqual(h.allowed, ["Primo"])
        h = next(s for s in tg.slots_for_month(cfg, [day], {"Primo": {day.date: {"Any"}}}) if s.columns == ["H"])
        self.assertEqual(h.allowed, ["Riserva"])
        self.assertEqual(h.emergency_doctors, ["Riserva"])

    def test_card_rules_reach_the_solver(self):
        cfg = dr.to_solver_cfg(YAML, dr.migrate(None, YAML))
        rules = cfg["rules"]
        self.assertEqual(rules["V"]["friday_required_doctor"], "Crea")
        self.assertEqual(rules["W"]["tuesday_fixed"], "Pugliatti")
        self.assertEqual(rules["AB"]["saturday_only_doctor"], "Crea")
        self.assertEqual(rules["J"]["monthly_quotas"]["Zito"], 2)
        self.assertEqual(rules["Festivi"]["quotas"]["Manganaro"], 2)
        self.assertEqual(cfg["pool_quota_overrides"][("Cimino", "U")], {"type": "fixed", "value": 2})
        self.assertEqual(rules["U"]["monday_v_rules"], [{"v_doctor": "Allegra", "u_doctors": ["Crea", "Dattilo"]}])
        self.assertEqual(rules["T"]["recupero_min_per_month"], 4)


if __name__ == "__main__":
    unittest.main()


class CardFormTests(unittest.TestCase):
    """The services table of the card page → card (doctor_admin_ui.card_from_form)."""

    VALUES = dict(active=True, reperibilita=True, no_sabato=False, festivi_diurni=True, festivi_notti=True,
                  festivi_quota=None, mai_notti=False, esente_weekend=False, universitario=False,
                  uni_ratio=0.6, uni_notte=True, j_weekend=False, j_wmin=None, j_wmax=1, v_doppio="secondo",
                  lun_vu=[], w_days=["Tue"], ab_sat=False, note="")

    def test_table_edits_become_the_card(self):
        import doctor_admin_ui as ui
        base = card(services={"J": {"level": "pool"}, "K": {"level": "pool", "quota": {"type": "max", "value": 4}}})
        table = ui.services_table(dr.clean_card(base))
        idx = {sid: i for i, sid in enumerate(dr.SERVICES)}
        table.loc[idx["K"], "Assegnabile"] = "—"          # tolto
        table.loc[idx["V"], "Assegnabile"] = "Sì"          # aggiunto
        table.loc[idx["H"], "Assegnabile"] = "Riserva"
        table.loc[idx["J"], "Quota mensile"] = "esatta"
        table.loc[idx["J"], "N"] = 2

        out = ui.card_from_form(dr.clean_card(base), table, self.VALUES)

        self.assertEqual(sorted(out["services"]), ["H", "J", "V"])
        self.assertEqual(out["services"]["H"], {"level": "riserva"})
        self.assertEqual(out["services"]["J"]["quota"], {"type": "fixed", "value": 2})
        self.assertFalse(out["services"]["J"]["notti_weekend"])
        self.assertEqual(out["services"]["J"]["weekend_max"], 1)
        self.assertEqual(out["services"]["V"]["doppio"], "secondo")

    def test_changed_elsewhere_is_refused(self):
        import doctor_admin_ui as ui
        reg = dr.set_doctor(dr.empty_registry(), "Trio", card(note="vecchia"))
        base = reg["doctors"]["Trio"]
        reg_elsewhere = dr.set_doctor(reg, "Trio", card(note="altrove"))
        with self.assertRaises(ui.CardChanged):
            ui.replace_card_if_unchanged(reg_elsewhere, "Trio", base, card(note="mia"))
        ok = ui.replace_card_if_unchanged(reg, "Trio", base, card(note="mia"))
        self.assertEqual(ok["doctors"]["Trio"]["note"], "mia")
