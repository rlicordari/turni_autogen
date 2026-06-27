import datetime as dt
import unittest

from pool_config_store import audit_pool_config, normalize_pool_config
from turni_generator import DayRow, Slot, apply_pool_config, slots_for_month, solve_with_ortools


class SolverRegressionTests(unittest.TestCase):
    def test_fixed_assignment_outside_domain_is_output(self):
        cfg = {
            "rules": {"D_F": {"allowed": ["A", "B"]}},
            "global_constraints": {},
        }
        day = DayRow(dt.date(2026, 4, 1), "Wed", 2)
        slot = Slot(day, "2026-04-01-D", ["D"], ["A"], required=True, shift="Mattina", rule_tag="D_F.D")

        assignment, stats = solve_with_ortools(
            cfg,
            [day],
            [slot],
            fixed_assignments=[{"doctor": "B", "date": "2026-04-01", "column": "D"}],
        )

        self.assertEqual(assignment["2026-04-01-D"], "B")
        self.assertNotEqual(stats["status"], "PARTIAL")

    def test_fixed_assignment_cannot_put_never_in_j(self):
        cfg = {
            "rules": {
                "J": {
                    "pool_other": ["Allegra", "Manganaro"],
                    "never_in_J": ["Manganaro"],
                }
            },
            "global_constraints": {},
        }
        day = DayRow(dt.date(2026, 7, 20), "Mon", 21)
        slot = Slot(day, "2026-07-20-J", ["J"], ["Allegra"], required=True, shift="Notte", rule_tag="J")

        assignment, stats = solve_with_ortools(
            cfg,
            [day],
            [slot],
            fixed_assignments=[{"doctor": "Manganaro", "date": "2026-07-20", "column": "J"}],
        )

        self.assertEqual(assignment["2026-07-20-J"], "Allegra")
        self.assertIn("J.never_in_J", "\n".join(stats.get("warnings", [])))

    def test_h_emergency_not_added_when_primary_exists(self):
        cfg = {
            "rules": {
                "H": {"pool_mon_fri": ["Manganaro"]},
                "J": {"pool_other": ["Zito"]},
            },
            "global_constraints": {},
            "pool_critical_services": {"H": {"fallback": "any"}},
        }
        day = DayRow(dt.date(2026, 7, 15), "Wed", 16)

        slots = slots_for_month(cfg, [day], {})
        h_slot = next(s for s in slots if s.columns == ["H"])

        self.assertIn("Manganaro", h_slot.allowed)
        self.assertNotIn("Zito", h_slot.allowed)
        self.assertFalse(h_slot.emergency_doctors)

    def test_kt_share_does_not_fire_when_separate_doctors_are_available(self):
        cfg = {
            "rules": {
                "K": {"pool": ["A", "B"]},
                "T": {"pool": ["A", "B"]},
            },
            "global_constraints": {"relief_valves": {"enable_kt_share": False}},
            "pool_service_combinations": [
                {"columns": ["K", "T"], "same_day": True, "mode": "always"}
            ],
        }
        day = DayRow(dt.date(2026, 7, 13), "Mon", 14)
        slots = [
            Slot(day, "2026-07-13-K", ["K"], ["A", "B"], required=True, shift="Mattina", rule_tag="K"),
            Slot(day, "2026-07-13-T", ["T"], ["A", "B"], required=True, shift="Mattina", rule_tag="T"),
        ]

        assignment, _stats = solve_with_ortools(cfg, [day], slots)

        self.assertIsNotNone(assignment["2026-07-13-K"])
        self.assertIsNotNone(assignment["2026-07-13-T"])
        self.assertNotEqual(assignment["2026-07-13-K"], assignment["2026-07-13-T"])

    def test_kt_share_fires_to_avoid_required_blank(self):
        cfg = {
            "rules": {
                "K": {"pool": ["A"]},
                "T": {"pool": ["A"]},
            },
            "global_constraints": {"relief_valves": {"enable_kt_share": True}},
        }
        day = DayRow(dt.date(2026, 7, 13), "Mon", 14)
        slots = [
            Slot(day, "2026-07-13-K", ["K"], ["A"], required=True, shift="Mattina", rule_tag="K"),
            Slot(day, "2026-07-13-T", ["T"], ["A"], required=True, shift="Mattina", rule_tag="T"),
        ]

        assignment, stats = solve_with_ortools(cfg, [day], slots)

        self.assertEqual(assignment["2026-07-13-K"], "A")
        self.assertEqual(assignment["2026-07-13-T"], "A")
        self.assertNotEqual(stats["status"], "PARTIAL")

    def test_pool_config_normalizes_kt_always_to_fallback(self):
        cfg = {
            "schema_version": 1,
            "service_combinations": [
                {"columns": ["K", "T"], "same_day": True, "mode": "always"},
                {"columns": ["Q", "R"], "same_day": True, "mode": "always"},
            ],
        }

        normalized = normalize_pool_config(cfg)

        self.assertEqual(normalized["service_combinations"][0]["mode"], "fallback")
        self.assertEqual(normalized["service_combinations"][1]["mode"], "always")

    def test_pool_config_normalizes_non_editable_columns_out_of_doctor_columns(self):
        cfg = {
            "schema_version": 1,
            "doctors": {
                "A": {
                    "active": True,
                    "columns": ["C", "K", "AC", "AD"],
                    "festivi_diurni": True,
                    "festivi_notti": True,
                    "excluded_from_reperibilita": False,
                    "university_doctor": None,
                    "column_overrides": {},
                }
            },
        }

        normalized = normalize_pool_config(cfg)

        self.assertEqual(normalized["doctors"]["A"]["columns"], ["K"])

    def test_pool_config_empty_gui_pool_does_not_fall_back_to_yaml_pool(self):
        cfg_yaml = {
            "columns": {"D": "UTIC mattina", "F": "Supporto 118"},
            "rules": {"D_F": {"allowed": ["Old"]}},
            "global_constraints": {},
        }
        pool_cfg = {
            "schema_version": 1,
            "doctors": {
                "Old": {
                    "active": True,
                    "columns": [],
                    "festivi_diurni": True,
                    "festivi_notti": True,
                    "excluded_from_reperibilita": False,
                    "university_doctor": None,
                    "column_overrides": {},
                }
            },
        }

        merged = apply_pool_config(cfg_yaml, pool_cfg)

        self.assertEqual(merged["rules"]["D_F"]["allowed"], [])

    def test_pool_config_audit_blocks_never_in_j_from_gui(self):
        cfg_yaml = {
            "columns": {"J": "Notte"},
            "rules": {"J": {"never_in_J": ["Manganaro"]}},
        }
        pool_cfg = {
            "schema_version": 1,
            "doctors": {
                "Manganaro": {
                    "active": True,
                    "columns": ["J"],
                    "festivi_diurni": True,
                    "festivi_notti": True,
                    "excluded_from_reperibilita": False,
                    "university_doctor": None,
                    "column_overrides": {},
                }
            },
        }

        audit = audit_pool_config(pool_cfg, cfg_yaml)

        self.assertTrue(any("J.never_in_J" in err for err in audit["errors"]))

    def test_pool_config_monthly_targets_are_merged_for_solver(self):
        cfg_yaml = {
            "columns": {"K": "Letto"},
            "rules": {"K": {"pool": ["A", "B"]}},
            "global_constraints": {},
        }
        doctor_template = {
            "active": True,
            "columns": ["K"],
            "festivi_diurni": True,
            "festivi_notti": True,
            "excluded_from_reperibilita": False,
            "university_doctor": None,
            "column_overrides": {},
        }
        pool_cfg = {
            "schema_version": 1,
            "doctors": {"A": dict(doctor_template), "B": dict(doctor_template)},
            "column_settings": {"K": {"monthly_target": 1, "counts_as": 1}},
        }

        merged = apply_pool_config(cfg_yaml, pool_cfg)

        self.assertEqual(merged["pool_monthly_targets"]["K"], 1)

    def test_pool_config_saturday_day_exclusion_is_merged_for_solver(self):
        cfg_yaml = {
            "columns": {"K": "Letto"},
            "rules": {"K": {"pool": ["A", "B"]}},
            "global_constraints": {},
        }
        doctor_template = {
            "active": True,
            "columns": ["K"],
            "festivi_diurni": True,
            "festivi_notti": True,
            "excluded_from_reperibilita": False,
            "university_doctor": None,
            "column_overrides": {},
        }
        pool_cfg = {
            "schema_version": 1,
            "doctors": {
                "A": {**doctor_template, "exclude_saturday_day": True},
                "B": dict(doctor_template),
            },
        }

        merged = apply_pool_config(cfg_yaml, pool_cfg)

        self.assertEqual(merged["pool_saturday_day_excluded"], {"A"})

    def test_saturday_day_exclusion_removes_morning_and_afternoon_but_not_night_or_c(self):
        cfg = {
            "columns": {"C": "Reperibilita", "H": "UTIC pomeriggio", "J": "Notte", "K": "Letto"},
            "rules": {
                "C_reperibilita": {"excluded": []},
                "H": {"pool_mon_fri": ["A", "B"]},
                "J": {"pool_other": ["A", "B"]},
                "K": {"pool": ["A", "B"]},
            },
            "global_constraints": {},
            "pool_saturday_day_excluded": {"A"},
        }
        sat = DayRow(dt.date(2026, 7, 4), "Sat", 5)
        fri = DayRow(dt.date(2026, 7, 3), "Fri", 4)

        slots = slots_for_month(cfg, [fri, sat], {})
        sat_k = next(s for s in slots if s.slot_id == "2026-07-04-K")
        sat_h = next(s for s in slots if s.slot_id == "2026-07-04-H")
        sat_j = next(s for s in slots if s.slot_id == "2026-07-04-J")
        sat_c = next(s for s in slots if s.slot_id == "2026-07-04-C")
        fri_k = next(s for s in slots if s.slot_id == "2026-07-03-K")

        self.assertNotIn("A", sat_k.allowed)
        self.assertNotIn("A", sat_h.allowed)
        self.assertIn("A", sat_j.allowed)
        self.assertIn("A", sat_c.allowed)
        self.assertIn("A", fri_k.allowed)

    def test_saturday_day_exclusion_applies_after_critical_fallback(self):
        cfg = {
            "columns": {"H": "UTIC pomeriggio"},
            "rules": {"H": {"pool_mon_fri": ["A"]}},
            "global_constraints": {},
            "pool_critical_services": {"H": {"fallback": "any"}},
            "pool_saturday_day_excluded": {"A"},
        }
        sat = DayRow(dt.date(2026, 7, 4), "Sat", 5)

        slots = slots_for_month(cfg, [sat], {})
        sat_h = next(s for s in slots if s.slot_id == "2026-07-04-H")

        self.assertNotIn("A", sat_h.allowed)


if __name__ == "__main__":
    unittest.main()
