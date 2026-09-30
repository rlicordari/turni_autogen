import datetime as dt
import multiprocessing
import tempfile
import unittest
from pathlib import Path

import generation_memory as gm
from pool_config_store import audit_pool_config, normalize_pool_config
from turni_generator import (
    DayRow,
    Slot,
    _repair_heavy_festive_assignments,
    assign_reperibilita_C,
    apply_pool_config,
    create_period_template_xlsx,
    load_template_days,
    solve_across_months,
    slots_for_month,
    solve_with_ortools,
)


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

    def test_pool_config_f_only_doctor_does_not_enter_df_primary_pair(self):
        cfg_yaml = {
            "columns": {"D": "UTIC mattina", "F": "Supporto 118"},
            "rules": {"D_F": {"allowed": ["Old"]}},
            "global_constraints": {},
        }
        pool_cfg = {
            "schema_version": 1,
            "doctors": {
                "Grimaldi": {
                    "active": True,
                    "columns": ["D", "F"],
                    "festivi_diurni": False,
                    "festivi_notti": False,
                    "excluded_from_reperibilita": False,
                    "university_doctor": None,
                    "column_overrides": {},
                },
                "Rubino": {
                    "active": True,
                    "columns": ["F"],
                    "festivi_diurni": True,
                    "festivi_notti": True,
                    "excluded_from_reperibilita": False,
                    "university_doctor": None,
                    "column_overrides": {},
                },
            },
        }

        merged = apply_pool_config(cfg_yaml, pool_cfg)

        self.assertEqual(merged["rules"]["D_F"]["allowed"], ["Grimaldi"])

    def test_l_relief_blank_penalty_has_solver_floor(self):
        cfg = {
            "rules": {"L": {"days": ["Mon"], "pool_other": ["A"]}},
            "global_constraints": {
                "relief_valves": {"allow_blank_columns": {"L": 20000}},
            },
        }
        day = DayRow(dt.date(2026, 8, 3), "Mon", 2)

        slots = slots_for_month(cfg, [day], {})
        l_slot = next(s for s in slots if s.columns == ["L"])

        self.assertFalse(l_slot.required)
        self.assertGreaterEqual(l_slot.blank_penalty, 20_000_000)

    def test_df_fallback_prefers_stable_consecutive_block(self):
        cfg = {
            "rules": {
                "D_F": {
                    "allowed": ["Grimaldi", "Calabrò"],
                    "days": "Mon-Sat",
                    "pattern_3_3": True,
                    "pattern_doc1": "Grimaldi",
                    "pattern_doc2": "Calabrò",
                    "enable_df_share": True,
                    "fallback_balance_penalty": 1,
                    "fallback_switch_penalty": 1000,
                    "fallback_block_days": 3,
                },
                "H": {"pool_mon_fri": ["A", "B", "C"]},
            },
            "global_constraints": {},
        }
        days = [
            DayRow(dt.date(2026, 8, 3), "Mon", 2),
            DayRow(dt.date(2026, 8, 4), "Tue", 3),
            DayRow(dt.date(2026, 8, 5), "Wed", 4),
        ]
        unav = {
            "Grimaldi": {d.date: {"Any"} for d in days},
            "Calabrò": {d.date: {"Any"} for d in days},
        }

        slots = slots_for_month(cfg, days, unav)
        assignment, stats = solve_with_ortools(cfg, days, slots)

        self.assertNotEqual(stats["status"], "PARTIAL")
        d_docs = [assignment[f"{d.date}-D"] for d in days]
        f_docs = [assignment[f"{d.date}-F"] for d in days]
        self.assertEqual(d_docs, f_docs)
        self.assertEqual(len(set(d_docs)), 1)

    def test_df_fallback_uses_h_pool_before_any_available_doctor(self):
        cfg = {
            "rules": {
                "D_F": {
                    "allowed": ["Grimaldi", "Calabrò"],
                    "days": "Mon-Sat",
                },
                "H": {"pool_mon_fri": ["Rubino"]},
                "J": {"pool_other": ["Dattilo"]},
            },
            "global_constraints": {},
        }
        day = DayRow(dt.date(2026, 8, 3), "Mon", 2)
        unav = {
            "Grimaldi": {day.date: {"Any"}},
            "Calabrò": {day.date: {"Any"}},
        }

        slots = slots_for_month(cfg, [day], unav)
        d_slot = next(s for s in slots if s.columns == ["D"])
        f_slot = next(s for s in slots if s.columns == ["F"])

        self.assertEqual(d_slot.allowed, ["Rubino"])
        self.assertEqual(f_slot.allowed, ["Rubino"])

    def test_df_single_available_primary_covers_both_columns(self):
        cfg = {
            "rules": {
                "D_F": {
                    "allowed": ["Grimaldi", "Calabrò"],
                    "days": "Mon-Sat",
                    "pattern_3_3": True,
                    "pattern_doc1": "Grimaldi",
                    "pattern_doc2": "Calabrò",
                    "enable_df_share": True,
                },
                "H": {"pool_mon_fri": ["Rubino", "Manganaro"]},
            },
            "global_constraints": {},
        }
        day = DayRow(dt.date(2026, 8, 21), "Fri", 2)
        unav = {"Calabrò": {day.date: {"Any"}}}

        slots = slots_for_month(cfg, [day], unav)
        assignment, stats = solve_with_ortools(cfg, [day], slots)

        self.assertNotEqual(stats["status"], "PARTIAL")
        self.assertEqual(assignment[f"{day.date}-D"], "Grimaldi")
        self.assertEqual(assignment[f"{day.date}-F"], "Grimaldi")

    def test_h_prefers_dedicated_candidate_to_preserve_k_pool(self):
        cfg = {
            "rules": {
                "H": {
                    "pool_mon_fri": ["Migliorato", "Manganaro"],
                    "reserve_multi_service_penalty": 1_000_000,
                    "reserve_for_columns": ["K"],
                },
                "K": {"pool": ["Manganaro", "Altro"]},
            },
            "global_constraints": {},
        }
        day = DayRow(dt.date(2026, 8, 20), "Thu", 2)
        slots = [
            Slot(day, f"{day.date}-H", ["H"], ["Migliorato", "Manganaro"], required=True, shift="Pomeriggio", rule_tag="H"),
            Slot(day, f"{day.date}-K", ["K"], ["Manganaro", "Altro"], required=True, shift="Mattina", rule_tag="K"),
        ]

        assignment, stats = solve_with_ortools(cfg, [day], slots)

        self.assertNotEqual(stats["status"], "PARTIAL")
        self.assertEqual(assignment[f"{day.date}-H"], "Migliorato")
        self.assertNotEqual(assignment[f"{day.date}-H"], "Manganaro")

    def test_k_blank_is_avoided_by_reshuffling_flexible_services(self):
        cfg = {
            "rules": {
                "H": {"pool_mon_fri": ["Migliorato", "Licordari", "Manganaro"]},
                "K": {"pool": ["Licordari", "Manganaro"], "no_consecutive_days_same_doctor": True},
                "Q": {"pool": ["Manganaro", "Zito"]},
            },
            "global_constraints": {
                "solver_max_time_seconds": 10,
                "k_pool_preserve_threshold": 5,
                "k_pool_preserve_penalty": 80_000_000,
            },
        }
        day20 = DayRow(dt.date(2026, 8, 20), "Thu", 2)
        day21 = DayRow(dt.date(2026, 8, 21), "Fri", 3)
        slots = [
            Slot(day20, "2026-08-20-H", ["H"], ["Migliorato"], required=True, shift="Pomeriggio", rule_tag="H"),
            Slot(day20, "2026-08-20-K", ["K"], ["Manganaro"], required=True, shift="Mattina", rule_tag="K"),
            Slot(day21, "2026-08-21-H", ["H"], ["Migliorato", "Licordari", "Manganaro"], required=True, shift="Pomeriggio", rule_tag="H"),
            Slot(day21, "2026-08-21-K", ["K"], ["Licordari", "Manganaro"], required=True, shift="Mattina", rule_tag="K"),
            Slot(day21, "2026-08-21-Q", ["Q"], ["Manganaro", "Zito"], required=True, shift="Mattina", rule_tag="Q"),
        ]

        assignment, stats = solve_with_ortools(cfg, [day20, day21], slots)

        self.assertNotEqual(stats["status"], "PARTIAL")
        self.assertEqual(assignment["2026-08-21-K"], "Licordari")
        self.assertEqual(assignment["2026-08-21-H"], "Manganaro")
        self.assertEqual(assignment["2026-08-21-Q"], "Zito")

    def test_eg_pool_is_preserved_for_k_when_k_has_outside_alternative(self):
        cfg = {
            "rules": {
                "E_G": {
                    "allowed": ["Manganaro", "Allegra"],
                    "protect_pool_penalty": 35_000_000,
                },
                "K": {"pool": ["Manganaro", "Colarusso"]},
            },
            "global_constraints": {},
        }
        day = DayRow(dt.date(2026, 8, 20), "Thu", 2)
        slots = [
            Slot(day, f"{day.date}-EG", ["E", "G"], ["Manganaro", "Allegra"], required=True, shift="Mattina", rule_tag="E_G"),
            Slot(day, f"{day.date}-K", ["K"], ["Manganaro", "Colarusso"], required=True, shift="Mattina", rule_tag="K"),
        ]

        assignment, stats = solve_with_ortools(cfg, [day], slots)

        self.assertNotEqual(stats["status"], "PARTIAL")
        self.assertEqual(assignment[f"{day.date}-K"], "Colarusso")
        self.assertIn(assignment[f"{day.date}-EG"], {"Manganaro", "Allegra"})

    def test_j_blank_week_overrides_are_reported_in_month_stats(self):
        cfg = {
            "columns": {"J": "Notte"},
            "rules": {"J": {"pool_other": ["A"], "thursday_blank": True}},
            "global_constraints": {"night_spacing_days_min": 1},
        }
        days = [
            DayRow(dt.date(2026, 8, 26), "Wed", 2),
            DayRow(dt.date(2026, 8, 27), "Thu", 3),
            DayRow(dt.date(2026, 8, 30), "Sun", 6),
        ]

        _slots, _assignment, stats = solve_across_months(
            cfg,
            days,
            {},
            j_blank_week_overrides={"2026-W35": ["2026-08-26", "2026-08-30"]},
        )

        self.assertEqual(
            stats["months"]["2026-08"]["j_blank_week_overrides"]["2026-W35"],
            ["2026-08-26", "2026-08-30"],
        )

    def test_generic_weekend_night_cap_is_soft_not_forced_blank(self):
        cfg = {
            "rules": {
                "J": {
                    "pool_other": ["A", "B", "C"],
                    "weekend_night_cap_penalty": 30_000_000,
                }
            },
            "global_constraints": {
                "night_spacing_days_min": 1,
                "night_off": {"same_day": True, "next_day": True},
            },
        }
        days = [
            DayRow(dt.date(2026, 8, 23) + dt.timedelta(days=i),
                   ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"][(dt.date(2026, 8, 23) + dt.timedelta(days=i)).weekday()],
                   i + 2)
            for i in range(8)
        ]
        day1 = days[0]
        day2 = days[-1]
        slots = [
            Slot(day1, "2026-08-23-J", ["J"], ["A"], required=True, shift="Notte", rule_tag="J"),
            Slot(day2, "2026-08-30-J", ["J"], ["A"], required=True, shift="Notte", rule_tag="J"),
        ]

        assignment, stats = solve_with_ortools(cfg, days, slots)

        self.assertNotEqual(stats["status"], "PARTIAL")
        self.assertEqual(assignment["2026-08-30-J"], "A")

    def test_sunday_j_balance_rotates_eligible_doctors(self):
        cfg = {
            "rules": {
                "J": {
                    "pool_other": ["A", "B"],
                    "free_night_quota_penalty": 0,
                    "free_night_spread_penalty": 0,
                    "weekend_night_cap_penalty": 0,
                    "sunday_night_balance_penalty": 40_000_000,
                }
            },
            "global_constraints": {"night_spacing_days_min": 1},
        }
        days = [
            DayRow(dt.date(2026, 8, 2), "Sun", 2),
            DayRow(dt.date(2026, 8, 9), "Sun", 9),
        ]
        slots = [
            Slot(day, f"{day.date}-J", ["J"], ["A", "B"], required=True, shift="Notte", rule_tag="J")
            for day in days
        ]

        assignment, stats = solve_with_ortools(cfg, days, slots)

        self.assertNotEqual(stats["status"], "PARTIAL")
        self.assertNotEqual(assignment["2026-08-02-J"], assignment["2026-08-09-J"])

    def test_free_j_spread_penalty_avoids_three_to_one_distribution(self):
        cfg = {
            "rules": {
                "J": {
                    "pool_other": ["A", "B"],
                    "free_night_quota_penalty": 0,
                    "free_night_spread_penalty": 20_000_000,
                    "weekend_night_cap_penalty": 0,
                    "sunday_night_balance_penalty": 0,
                }
            },
            "global_constraints": {"night_spacing_days_min": 1},
        }
        days = [
            DayRow(dt.date(2026, 8, 3), "Mon", 2),
            DayRow(dt.date(2026, 8, 4), "Tue", 3),
            DayRow(dt.date(2026, 8, 5), "Wed", 4),
            DayRow(dt.date(2026, 8, 6), "Thu", 5),
        ]
        slots = [
            Slot(day, f"{day.date}-J", ["J"], ["A", "B"], required=True, shift="Notte", rule_tag="J")
            for day in days
        ]

        assignment, stats = solve_with_ortools(cfg, days, slots)

        self.assertNotEqual(stats["status"], "PARTIAL")
        counts = {"A": 0, "B": 0}
        for day in days:
            counts[assignment[f"{day.date}-J"]] += 1
        self.assertEqual(sorted(counts.values()), [2, 2])

    def test_j_weekday_and_weekend_balances_are_independent(self):
        cfg = {
            "rules": {
                "J": {
                    "pool_other": ["A", "B"],
                    "free_night_quota_penalty": 0,
                    "free_night_spread_penalty": 0,
                    "weekday_night_spread_penalty": 25_000_000,
                    "weekend_night_cap_penalty": 0,
                    "weekend_night_spread_penalty": 30_000_000,
                    "sunday_night_balance_penalty": 0,
                }
            },
            "global_constraints": {"night_spacing_days_min": 1},
        }
        days = [
            DayRow(dt.date(2026, 8, 3), "Mon", 2),
            DayRow(dt.date(2026, 8, 4), "Tue", 3),
            DayRow(dt.date(2026, 8, 8), "Sat", 7),
            DayRow(dt.date(2026, 8, 9), "Sun", 8),
        ]
        slots = [
            Slot(day, f"{day.date}-J", ["J"], ["A", "B"], required=True, shift="Notte", rule_tag="J")
            for day in days
        ]

        assignment, stats = solve_with_ortools(cfg, days, slots)

        self.assertNotEqual(stats["status"], "PARTIAL")
        weekday_counts = {"A": 0, "B": 0}
        weekend_counts = {"A": 0, "B": 0}
        for day in days:
            doc = assignment[f"{day.date}-J"]
            if day.dow in ("Sat", "Sun"):
                weekend_counts[doc] += 1
            else:
                weekday_counts[doc] += 1
        self.assertEqual(sorted(weekday_counts.values()), [1, 1])
        self.assertEqual(sorted(weekend_counts.values()), [1, 1])

    def test_j_prefers_two_weekday_two_festive_over_three_one_split(self):
        cfg = {
            "rules": {
                "J": {
                    "pool_other": ["A", "B"],
                    "free_night_quota_penalty": 0,
                    "free_night_spread_penalty": 20_000_000,
                    "weekday_night_spread_penalty": 0,
                    "weekend_night_cap_penalty": 0,
                    "weekend_night_spread_penalty": 0,
                    "sunday_night_balance_penalty": 0,
                    "j_type_split_penalty": 35_000_000,
                }
            },
            "global_constraints": {"night_spacing_days_min": 1},
        }
        raw_days = [
            (dt.date(2026, 8, 3), "Mon"),
            (dt.date(2026, 8, 4), "Tue"),
            (dt.date(2026, 8, 5), "Wed"),
            (dt.date(2026, 8, 6), "Thu"),
            (dt.date(2026, 8, 8), "Sat"),
            (dt.date(2026, 8, 9), "Sun"),
            (dt.date(2026, 8, 15), "Sat"),
            (dt.date(2026, 8, 16), "Sun"),
        ]
        days = [DayRow(day, dow, i + 2) for i, (day, dow) in enumerate(raw_days)]
        slots = [
            Slot(day, f"{day.date}-J", ["J"], ["A", "B"], required=True, shift="Notte", rule_tag="J")
            for day in days
        ]

        assignment, stats = solve_with_ortools(cfg, days, slots)

        self.assertNotEqual(stats["status"], "PARTIAL")
        split = {"A": {"weekday": 0, "festive": 0}, "B": {"weekday": 0, "festive": 0}}
        for day in days:
            doc = assignment[f"{day.date}-J"]
            key = "festive" if day.dow in ("Sat", "Sun") else "weekday"
            split[doc][key] += 1
        self.assertEqual(split["A"], {"weekday": 2, "festive": 2})
        self.assertEqual(split["B"], {"weekday": 2, "festive": 2})

    def test_festive_de_hi_balance_counts_morning_and_afternoon_together(self):
        cfg = {
            "rules": {
                "Festivi": {
                    "pool": ["A", "B", "C", "D"],
                    "sunday_de_hi_balance_penalty": 25_000_000,
                }
            },
            "global_constraints": {},
        }
        days = [
            DayRow(dt.date(2026, 8, 2), "Sun", 2),
            DayRow(dt.date(2026, 8, 9), "Sun", 9),
        ]
        slots = []
        for day in days:
            slots.append(Slot(day, f"{day.date}-DE", ["D", "E"], ["A", "B", "C", "D"], required=True, shift="Mattina", rule_tag="Festivo_DE"))
            slots.append(Slot(day, f"{day.date}-HI", ["H", "I"], ["A", "B", "C", "D"], required=True, shift="Pomeriggio", rule_tag="Festivo_HI"))

        assignment, stats = solve_with_ortools(cfg, days, slots)

        self.assertNotEqual(stats["status"], "PARTIAL")
        counts = {"A": 0, "B": 0, "C": 0, "D": 0}
        for value in assignment.values():
            counts[value] += 1
        self.assertEqual(sorted(counts.values()), [1, 1, 1, 1])

    def test_festive_day_quota_avoids_reusing_doctor_with_forced_festive(self):
        cfg = {
            "rules": {
                "Festivi": {
                    "pool": ["A", "B", "C", "D"],
                    "festivi_diurni_quota_penalty": 70_000_000,
                }
            },
            "global_constraints": {},
        }
        days = [
            DayRow(dt.date(2026, 8, 2), "Sun", 2),
            DayRow(dt.date(2026, 8, 9), "Sun", 9),
            DayRow(dt.date(2026, 8, 15), "Sat", 15),
            DayRow(dt.date(2026, 8, 16), "Sun", 16),
        ]
        slots = [
            Slot(days[0], "2026-08-02-DE", ["D", "E"], ["A"], required=True, shift="Mattina", rule_tag="Festivo_DE"),
            Slot(days[1], "2026-08-09-HI", ["H", "I"], ["A", "B", "C", "D"], required=True, shift="Pomeriggio", rule_tag="Festivo_HI"),
            Slot(days[2], "2026-08-15-DE", ["D", "E"], ["A", "B", "C", "D"], required=True, shift="Mattina", rule_tag="Festivo_DE"),
            Slot(days[3], "2026-08-16-HI", ["H", "I"], ["A", "B", "C", "D"], required=True, shift="Pomeriggio", rule_tag="Festivo_HI"),
        ]

        assignment, stats = solve_with_ortools(cfg, days, slots)

        self.assertNotEqual(stats["status"], "PARTIAL")
        counts = {"A": 0, "B": 0, "C": 0, "D": 0}
        for value in assignment.values():
            counts[value] += 1
        self.assertEqual(counts["A"], 1)
        self.assertEqual(sorted(counts.values()), [1, 1, 1, 1])

    def test_festive_day_balance_carries_across_custom_period_months(self):
        cfg = {
            "rules": {
                "Festivi": {
                    "pool": ["A", "B", "C", "D"],
                    "festivi_diurni_quota_penalty": 250_000_000,
                    "festivi_diurni_concentration_penalty": 90_000_000,
                }
            },
            "global_constraints": {},
        }
        days = [
            DayRow(dt.date(2026, 8, 2), "Sun", 2),
            DayRow(dt.date(2026, 9, 6), "Sun", 3),
        ]

        _slots, assignment, stats = solve_across_months(
            cfg,
            days,
            {},
            fixed_assignments=[
                {"doctor": "A", "date": "2026-08-02", "column": "D"},
                {"doctor": "B", "date": "2026-08-02", "column": "H"},
            ],
        )

        self.assertNotEqual(stats["status"], "PARTIAL")
        september_docs = {
            assignment["2026-09-06-DE"],
            assignment["2026-09-06-HI"],
        }
        self.assertEqual(september_docs, {"C", "D"})

    def test_festive_j_and_festive_day_duties_are_alternatives(self):
        cfg = {
            "rules": {
                "J": {
                    "pool_other": ["A", "B", "C", "D"],
                    "free_night_quota_penalty": 0,
                    "free_night_spread_penalty": 0,
                    "weekday_night_spread_penalty": 0,
                    "weekend_night_cap_penalty": 0,
                    "weekend_night_spread_penalty": 0,
                    "sunday_night_balance_penalty": 0,
                    "j_type_split_penalty": 0,
                    "festive_j_day_overlap_penalty": 80_000_000,
                    "festive_total_spread_penalty": 35_000_000,
                },
                "Festivi": {"pool": ["A", "B", "C", "D"]},
            },
            "global_constraints": {"night_spacing_days_min": 1},
        }
        raw_slots = [
            (dt.date(2026, 8, 2), "Sun", "J", ["J"], "Notte", "J"),
            (dt.date(2026, 8, 9), "Sun", "J", ["J"], "Notte", "J"),
            (dt.date(2026, 8, 16), "Sun", "DE", ["D", "E"], "Mattina", "Festivo_DE"),
            (dt.date(2026, 8, 23), "Sun", "HI", ["H", "I"], "Pomeriggio", "Festivo_HI"),
        ]
        days = [DayRow(day, dow, i + 2) for i, (day, dow, *_rest) in enumerate(raw_slots)]
        slots = [
            Slot(
                days[i],
                f"{day}-{suffix}",
                columns,
                ["A", "B", "C", "D"],
                required=True,
                shift=shift,
                rule_tag=tag,
            )
            for i, (day, _dow, suffix, columns, shift, tag) in enumerate(raw_slots)
        ]

        assignment, stats = solve_with_ortools(cfg, days, slots)

        self.assertNotEqual(stats["status"], "PARTIAL")
        counts = {"A": 0, "B": 0, "C": 0, "D": 0}
        for value in assignment.values():
            counts[value] += 1
        self.assertEqual(sorted(counts.values()), [1, 1, 1, 1])

    def test_weekend_j_and_festive_day_share_one_heavy_festive_quota(self):
        cfg = {
            "rules": {
                "J": {
                    "pool_other": ["A", "B", "C", "D"],
                    "free_night_quota_penalty": 0,
                    "free_night_spread_penalty": 0,
                    "weekday_night_spread_penalty": 0,
                    "weekend_night_cap_penalty": 0,
                    "weekend_night_spread_penalty": 0,
                    "sunday_night_balance_penalty": 0,
                    "j_type_split_penalty": 0,
                    "festive_j_day_overlap_penalty": 0,
                    "festive_total_spread_penalty": 0,
                    "heavy_festive_quota_penalty": 300_000_000,
                    "heavy_festive_concentration_penalty": 120_000_000,
                },
                "Festivi": {"pool": ["A", "B", "C", "D"]},
            },
            "global_constraints": {"night_spacing_days_min": 1},
        }
        raw_slots = [
            (dt.date(2026, 8, 2), "Sun", "J", ["J"], "Notte", "J"),
            (dt.date(2026, 8, 9), "Sun", "J", ["J"], "Notte", "J"),
            (dt.date(2026, 8, 16), "Sun", "DE", ["D", "E"], "Mattina", "Festivo_DE"),
            (dt.date(2026, 8, 23), "Sun", "HI", ["H", "I"], "Pomeriggio", "Festivo_HI"),
        ]
        days = [DayRow(day, dow, i + 2) for i, (day, dow, *_rest) in enumerate(raw_slots)]
        slots = [
            Slot(days[i], f"{day}-{suffix}", columns, ["A", "B", "C", "D"], required=True, shift=shift, rule_tag=tag)
            for i, (day, _dow, suffix, columns, shift, tag) in enumerate(raw_slots)
        ]

        assignment, stats = solve_with_ortools(cfg, days, slots)

        self.assertNotEqual(stats["status"], "PARTIAL")
        counts = {"A": 0, "B": 0, "C": 0, "D": 0}
        for value in assignment.values():
            counts[value] += 1
        self.assertEqual(sorted(counts.values()), [1, 1, 1, 1])

    def test_heavy_festive_load_uses_all_candidate_doctors_before_reusing_one(self):
        cfg = {
            "rules": {
                "J": {
                    "pool_other": ["A", "B", "C"],
                    "free_night_quota_penalty": 0,
                    "free_night_spread_penalty": 0,
                    "weekend_night_cap_penalty": 0,
                    "weekend_night_spread_penalty": 0,
                    "sunday_night_balance_penalty": 0,
                    "festive_total_spread_penalty": 0,
                    "heavy_festive_quota_penalty": 0,
                    "heavy_festive_concentration_penalty": 0,
                    "heavy_festive_consecutive_penalty": 0,
                    "heavy_festive_unused_candidate_penalty": 450_000_000,
                },
                "Festivi": {"pool": ["A", "B", "C"]},
            },
            "global_constraints": {"night_spacing_days_min": 1},
        }
        days = [
            DayRow(dt.date(2026, 8, 2), "Sun", 2),
            DayRow(dt.date(2026, 8, 9), "Sun", 9),
            DayRow(dt.date(2026, 8, 16), "Sun", 16),
        ]
        slots = [
            Slot(days[0], "2026-08-02-J", ["J"], ["A"], required=True, shift="Notte", rule_tag="J"),
            Slot(days[1], "2026-08-09-DE", ["D", "E"], ["A", "B", "C"], required=True, shift="Mattina", rule_tag="Festivo_DE"),
            Slot(days[2], "2026-08-16-HI", ["H", "I"], ["A", "B", "C"], required=True, shift="Pomeriggio", rule_tag="Festivo_HI"),
        ]

        assignment, stats = solve_with_ortools(cfg, days, slots)

        self.assertNotEqual(stats["status"], "PARTIAL")
        counts = {"A": 0, "B": 0, "C": 0}
        for value in assignment.values():
            counts[value] += 1
        self.assertEqual(sorted(counts.values()), [1, 1, 1])

    def test_second_look_moves_heavy_slot_to_unused_available_candidate(self):
        cfg = {
            "rules": {"J": {"pool_other": ["Cusmà", "Licordari"]}},
            "global_constraints": {
                "night_spacing_days_min": 5,
                "night_off": {"same_day": True, "next_day": True},
            },
        }
        days = [
            DayRow(dt.date(2026, 8, 2), "Sun", 2),
            DayRow(dt.date(2026, 8, 9), "Sun", 9),
        ]
        slots = [
            Slot(days[0], "2026-08-02-J", ["J"], ["Cusmà", "Licordari"], required=True, shift="Notte", rule_tag="J"),
            Slot(days[1], "2026-08-09-HI", ["H", "I"], ["Cusmà"], required=True, shift="Pomeriggio", rule_tag="Festivo_HI"),
        ]
        assignment = {
            "2026-08-02-J": "Cusmà",
            "2026-08-09-HI": "Cusmà",
        }

        repaired, swaps = _repair_heavy_festive_assignments(cfg, slots, assignment)

        self.assertEqual(repaired["2026-08-02-J"], "Licordari")
        self.assertEqual(repaired["2026-08-09-HI"], "Cusmà")
        self.assertEqual(len(swaps), 1)

    def test_second_look_does_not_change_fixed_heavy_slot(self):
        cfg = {
            "rules": {"Festivi": {"pool": ["A", "B"]}},
            "global_constraints": {},
        }
        day = DayRow(dt.date(2026, 8, 2), "Sun", 2)
        slots = [
            Slot(day, "2026-08-02-DE", ["D", "E"], ["A", "B"], required=True, shift="Mattina", rule_tag="Festivo_DE"),
            Slot(day, "2026-08-02-HI", ["H", "I"], ["A"], required=True, shift="Pomeriggio", rule_tag="Festivo_HI"),
        ]
        assignment = {
            "2026-08-02-DE": "A",
            "2026-08-02-HI": "A",
        }

        repaired, swaps = _repair_heavy_festive_assignments(
            cfg,
            slots,
            assignment,
            fixed_assignments=[{"doctor": "A", "date": "2026-08-02", "column": "D"}],
        )

        self.assertEqual(repaired["2026-08-02-DE"], "A")
        self.assertEqual(swaps, [])

    def test_heavy_festive_duty_avoids_consecutive_j_then_hi_same_doctor(self):
        cfg = {
            "rules": {
                "J": {
                    "pool_other": ["A", "B"],
                    "free_night_quota_penalty": 0,
                    "free_night_spread_penalty": 0,
                    "weekend_night_cap_penalty": 0,
                    "weekend_night_spread_penalty": 0,
                    "sunday_night_balance_penalty": 0,
                    "festive_total_spread_penalty": 0,
                    "heavy_festive_quota_penalty": 0,
                    "heavy_festive_concentration_penalty": 0,
                    "heavy_festive_consecutive_penalty": 220_000_000,
                },
                "Festivi": {"pool": ["A", "B"]},
            },
            "global_constraints": {"night_spacing_days_min": 1},
        }
        days = [
            DayRow(dt.date(2026, 8, 2), "Sun", 2),
            DayRow(dt.date(2026, 8, 9), "Sun", 9),
        ]
        slots = [
            Slot(days[0], "2026-08-02-J", ["J"], ["A"], required=True, shift="Notte", rule_tag="J"),
            Slot(days[1], "2026-08-09-HI", ["H", "I"], ["A", "B"], required=True, shift="Pomeriggio", rule_tag="Festivo_HI"),
        ]

        assignment, stats = solve_with_ortools(cfg, days, slots)

        self.assertNotEqual(stats["status"], "PARTIAL")
        self.assertEqual(assignment["2026-08-02-J"], "A")
        self.assertEqual(assignment["2026-08-09-HI"], "B")

    def test_repeated_same_festive_day_type_is_penalized(self):
        cfg = {
            "rules": {
                "J": {
                    "pool_other": ["A", "B"],
                    "festive_j_day_overlap_penalty": 0,
                    "festive_total_spread_penalty": 0,
                    "festive_same_type_repeat_penalty": 45_000_000,
                },
                "Festivi": {"pool": ["A", "B"]},
            },
            "global_constraints": {},
        }
        days = [
            DayRow(dt.date(2026, 8, 2), "Sun", 2),
            DayRow(dt.date(2026, 8, 9), "Sun", 9),
        ]
        slots = [
            Slot(days[0], "2026-08-02-DE", ["D", "E"], ["A", "B"], required=True, shift="Mattina", rule_tag="Festivo_DE"),
            Slot(days[1], "2026-08-09-DE", ["D", "E"], ["A", "B"], required=True, shift="Mattina", rule_tag="Festivo_DE"),
        ]

        assignment, stats = solve_with_ortools(cfg, days, slots)

        self.assertNotEqual(stats["status"], "PARTIAL")
        self.assertNotEqual(assignment["2026-08-02-DE"], assignment["2026-08-09-DE"])

    def test_festive_day_duty_avoids_consecutive_festive_dates_even_if_type_changes(self):
        cfg = {
            "rules": {
                "J": {
                    "pool_other": ["A", "B"],
                    "festive_j_day_overlap_penalty": 0,
                    "festive_total_spread_penalty": 0,
                    "festive_same_type_repeat_penalty": 0,
                    "festive_day_consecutive_penalty": 70_000_000,
                },
                "Festivi": {"pool": ["A", "B"]},
            },
            "global_constraints": {},
        }
        days = [
            DayRow(dt.date(2026, 8, 2), "Sun", 2),
            DayRow(dt.date(2026, 8, 9), "Sun", 9),
        ]
        slots = [
            Slot(days[0], "2026-08-02-DE", ["D", "E"], ["A", "B"], required=True, shift="Mattina", rule_tag="Festivo_DE"),
            Slot(days[1], "2026-08-09-HI", ["H", "I"], ["A", "B"], required=True, shift="Pomeriggio", rule_tag="Festivo_HI"),
        ]

        assignment, stats = solve_with_ortools(cfg, days, slots)

        self.assertNotEqual(stats["status"], "PARTIAL")
        self.assertNotEqual(assignment["2026-08-02-DE"], assignment["2026-08-09-HI"])

    def test_provisional_c_reperibilita_does_not_force_nearby_j_blank(self):
        cfg = {
            "rules": {
                "C_reperibilita": {
                    "excluded": [],
                    "constraints": ["not_night_next_2_days", "not_night_prev_2_days", "not_night_same_day"],
                    "max_per_doctor": 3,
                    "spacing_min_days": 0,
                },
                "J": {"pool_other": ["A"]},
            },
            "global_constraints": {
                "night_spacing_days_min": 1,
                "night_off": {"same_day": True, "next_day": True},
            },
        }
        days = [
            DayRow(dt.date(2026, 8, 22), "Sat", 2),
            DayRow(dt.date(2026, 8, 23), "Sun", 3),
            DayRow(dt.date(2026, 8, 24), "Mon", 4),
        ]
        slots = [
            Slot(days[0], "2026-08-22-C", ["C"], ["A"], required=True, shift="Any", rule_tag="C_reperibilita"),
            Slot(days[2], "2026-08-24-J", ["J"], ["A"], required=True, shift="Notte", rule_tag="J"),
        ]

        assignment, stats = solve_with_ortools(cfg, days, slots)

        self.assertNotEqual(stats["status"], "PARTIAL")
        self.assertEqual(assignment["2026-08-24-J"], "A")

    def test_tiny_i_ab_pool_is_preserved_for_k_when_k_has_outside_alternative(self):
        cfg = {
            "rules": {
                "I": {"distribution_pool": ["Allegra", "Crea"]},
                "AB": {"fallback_pool": ["Allegra", "Crea"]},
                "K": {"pool": ["Allegra", "Colarusso"]},
            },
            "global_constraints": {"tiny_pool_protect_penalty": 45_000_000},
        }
        day = DayRow(dt.date(2026, 8, 20), "Thu", 2)
        slots = [
            Slot(day, f"{day.date}-I", ["I"], ["Allegra", "Crea"], required=True, shift="Pomeriggio", rule_tag="I"),
            Slot(day, f"{day.date}-AB", ["AB"], ["Allegra", "Crea"], required=True, shift="Mattina", rule_tag="AB"),
            Slot(day, f"{day.date}-K", ["K"], ["Allegra", "Colarusso"], required=True, shift="Mattina", rule_tag="K"),
        ]

        assignment, stats = solve_with_ortools(cfg, [day], slots)

        self.assertNotEqual(stats["status"], "PARTIAL")
        self.assertEqual(assignment[f"{day.date}-K"], "Colarusso")
        self.assertEqual({assignment[f"{day.date}-I"], assignment[f"{day.date}-AB"]}, {"Allegra", "Crea"})

    def test_night_before_tiny_i_ab_pool_prefers_outside_doctor(self):
        cfg = {
            "rules": {
                "J": {"pool_other": ["Allegra", "Zito"]},
                "I": {"distribution_pool": ["Allegra", "Crea"]},
                "AB": {"fallback_pool": ["Allegra", "Crea"]},
            },
            "global_constraints": {
                "night_off": {"same_day": True, "next_day": True},
                "night_spacing_days_min": 1,
                "next_day_tiny_pool_night_penalty": 60_000_000,
            },
        }
        day1 = DayRow(dt.date(2026, 8, 19), "Wed", 2)
        day2 = DayRow(dt.date(2026, 8, 20), "Thu", 3)
        slots = [
            Slot(day1, f"{day1.date}-J", ["J"], ["Allegra", "Zito"], required=True, shift="Notte", rule_tag="J"),
            Slot(day2, f"{day2.date}-I", ["I"], ["Allegra", "Crea"], required=True, shift="Pomeriggio", rule_tag="I"),
            Slot(day2, f"{day2.date}-AB", ["AB"], ["Allegra", "Crea"], required=True, shift="Mattina", rule_tag="AB"),
        ]

        assignment, stats = solve_with_ortools(cfg, [day1, day2], slots)

        self.assertNotEqual(stats["status"], "PARTIAL")
        self.assertEqual(assignment[f"{day1.date}-J"], "Zito")
        self.assertEqual({assignment[f"{day2.date}-I"], assignment[f"{day2.date}-AB"]}, {"Allegra", "Crea"})

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

    def test_calabro_can_cover_saturday_df_unless_gui_flag_excludes_her(self):
        cfg = {
            "columns": {"D": "UTIC mattina", "F": "Supporto 118"},
            "rules": {
                "D_F": {
                    "allowed": ["Grimaldi", "Calabrò"],
                    "days": "Mon-Sat",
                }
            },
            "global_constraints": {},
        }
        sat = DayRow(dt.date(2026, 8, 29), "Sat", 2)

        slots = slots_for_month(cfg, [sat], {})
        sat_d = next(s for s in slots if s.slot_id == "2026-08-29-D")

        self.assertIn("Calabrò", sat_d.allowed)

        cfg_flagged = {**cfg, "pool_saturday_day_excluded": {"Calabrò"}}
        flagged_slots = slots_for_month(cfg_flagged, [sat], {})
        flagged_d = next(s for s in flagged_slots if s.slot_id == "2026-08-29-D")

        self.assertNotIn("Calabrò", flagged_d.allowed)

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

    def test_disabled_column_is_not_scheduled(self):
        cfg = {
            "columns": {"Q": "ECO base", "K": "Letto"},
            "rules": {
                "Q": {"days": "Mon-Sat", "pool": ["A"]},
                "K": {"pool": ["A"]},
            },
            "global_constraints": {},
        }
        day = DayRow(dt.date(2026, 8, 3), "Mon", 2)

        slots = slots_for_month(cfg, [day], {}, disabled_columns=["Q"])

        self.assertFalse(any("Q" in s.columns for s in slots))
        self.assertTrue(any(s.columns == ["K"] for s in slots))

    def test_disabled_column_partially_trims_combined_slot(self):
        cfg = {
            "columns": {"E": "Cardiologia mattina", "G": "Riabilitazione"},
            "rules": {"E_G": {"days": "Mon-Sat", "allowed": ["A"]}},
            "global_constraints": {},
        }
        day = DayRow(dt.date(2026, 8, 3), "Mon", 2)

        slots = slots_for_month(cfg, [day], {}, disabled_columns=["G"])
        eg_slot = next(s for s in slots if s.rule_tag == "E_G")

        self.assertEqual(eg_slot.columns, ["E"])

    def test_create_period_template_spans_custom_range(self):
        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "period.xlsx"
            create_period_template_xlsx(
                "Regole_Turni.yml",
                dt.date(2026, 8, 1),
                dt.date(2026, 9, 6),
                out,
            )

            wb, _ws, days = load_template_days(out)
            try:
                self.assertEqual(len(days), 37)
                self.assertEqual(days[0].date, dt.date(2026, 8, 1))
                self.assertEqual(days[-1].date, dt.date(2026, 9, 6))
            finally:
                wb.close()

    def test_solve_across_custom_period_carries_night_into_next_month(self):
        cfg = {
            "columns": {"J": "Notte", "K": "Letto"},
            "rules": {
                "J": {"pool_other": ["A"]},
                "K": {"pool": ["A", "B"]},
            },
            "global_constraints": {
                "night_spacing_days_min": 5,
                "night_off": {"same_day": True, "next_day": True},
            },
        }
        days = [
            DayRow(dt.date(2026, 8, 31), "Mon", 2),
            DayRow(dt.date(2026, 9, 1), "Tue", 3),
        ]

        slots, _assignment, _stats = solve_across_months(
            cfg,
            days,
            {},
            fixed_assignments=[{"doctor": "A", "date": "2026-08-31", "column": "J"}],
        )
        sep_k = next(s for s in slots if s.slot_id == "2026-09-01-K")

        self.assertNotIn("A", sep_k.allowed)
        self.assertIn("B", sep_k.allowed)

    def test_monthly_target_without_fixed_override_does_not_crash(self):
        cfg = {
            "columns": {"K": "Letto"},
            "rules": {"K": {"pool": ["A", "B"], "balance_weight": 100}},
            "global_constraints": {},
            "pool_monthly_targets": {"K": 1},
        }
        day = DayRow(dt.date(2026, 8, 3), "Mon", 2)
        slot = Slot(day, "2026-08-03-K", ["K"], ["A", "B"], required=True, shift="Mattina", rule_tag="K")

        assignment, stats = solve_with_ortools(cfg, [day], [slot])

        self.assertIn(assignment["2026-08-03-K"], {"A", "B"})
        self.assertNotEqual(stats["status"], "INFEASIBLE")

    def test_reperibilita_c_partial_month_does_not_require_monthly_minimum(self):
        docs = [f"Doc{i}" for i in range(13)]
        cfg = {
            "rules": {
                "C_reperibilita": {
                    "min_per_doctor": 2,
                    "max_per_doctor": 3,
                    "target_per_doctor": 2,
                    "spacing_min_days": 3,
                    "excluded": [],
                    "constraints": [],
                }
            },
            "global_constraints": {},
        }
        days = [
            DayRow(dt.date(2026, 9, day), ["Tue", "Wed", "Thu", "Fri", "Sat", "Sun"][idx], idx + 2)
            for idx, day in enumerate(range(1, 7))
        ]
        slots = [
            Slot(d, f"{d.date}-C", ["C"], docs, required=True, shift="Any", rule_tag="C_reperibilita")
            for d in days
        ]

        assignment, diag = assign_reperibilita_C(cfg, days, slots, {})

        self.assertTrue(all(assignment.get(f"{d.date}-C") for d in days))
        self.assertIn("OK", diag["C_reperibilita_diag"]["status"])

    def test_reperibilita_c_desired_distribution_is_not_a_hard_cap(self):
        docs = ["A", "B", "C"]
        cfg = {
            "rules": {
                "C_reperibilita": {
                    "min_per_doctor": 1,
                    "max_per_doctor": 2,
                    "target_per_doctor": 1,
                    "spacing_min_days": 0,
                    "excluded": [],
                    "constraints": [],
                }
            },
            "global_constraints": {},
        }
        days = [
            DayRow(dt.date(2026, 8, 3), "Mon", 2),
            DayRow(dt.date(2026, 8, 4), "Tue", 3),
            DayRow(dt.date(2026, 8, 5), "Wed", 4),
            DayRow(dt.date(2026, 8, 6), "Thu", 5),
            DayRow(dt.date(2026, 8, 7), "Fri", 6),
        ]
        slots = [
            Slot(days[0], "2026-08-03-C", ["C"], ["C"], required=True, shift="Any", rule_tag="C_reperibilita"),
            Slot(days[1], "2026-08-04-C", ["C"], ["C"], required=True, shift="Any", rule_tag="C_reperibilita"),
            Slot(days[2], "2026-08-05-C", ["C"], ["A"], required=True, shift="Any", rule_tag="C_reperibilita"),
            Slot(days[3], "2026-08-06-C", ["C"], ["A"], required=True, shift="Any", rule_tag="C_reperibilita"),
            Slot(days[4], "2026-08-07-C", ["C"], ["B"], required=True, shift="Any", rule_tag="C_reperibilita"),
        ]

        assignment, diag = assign_reperibilita_C(cfg, days, slots, {})

        self.assertEqual(assignment["2026-08-03-C"], "C")
        self.assertEqual(assignment["2026-08-04-C"], "C")
        self.assertIn("OK", diag["C_reperibilita_diag"]["status"])
        self.assertGreaterEqual(diag["C_reperibilita_diag"]["effective_max_per_doctor"], 2)

    def test_j_partial_month_does_not_cap_free_doctors_at_zero(self):
        cfg = {
            "rules": {
                "J": {
                    "pool_other": ["A", "B", "C", "D", "E", "Zito", "Calabrò", "Dattilo"],
                    "monthly_quotas": {
                        "Zito": 2,
                        "Calabrò": 2,
                        "Dattilo": 2,
                    },
                }
            },
            "global_constraints": {
                "night_spacing_days_min": 5,
                "night_off": {"same_day": True, "next_day": True},
            },
        }
        days = [
            DayRow(dt.date(2026, 9, day), ["Tue", "Wed", "Thu", "Fri", "Sat"][idx], idx + 2)
            for idx, day in enumerate(range(1, 6))
        ]
        doctors = ["A", "B", "C", "D", "E", "Zito", "Calabrò", "Dattilo"]
        slots = [
            Slot(d, f"{d.date}-J", ["J"], doctors, required=True, shift="Notte", rule_tag="J")
            for d in days
        ]

        assignment, stats = solve_with_ortools(cfg, days, slots)

        self.assertNotEqual(stats["status"], "PARTIAL")
        self.assertTrue(all(assignment.get(f"{d.date}-J") for d in days))
        self.assertTrue(any(assignment.get(f"{d.date}-J") in {"A", "B", "C", "D", "E"} for d in days))

    def test_j_quota_uses_prior_generation_memory_as_monthly_consumed_quota(self):
        cfg = {
            "rules": {
                "J": {
                    "pool_other": ["Zito", "A", "B", "C"],
                    "monthly_quotas": {"Zito": 2},
                }
            },
            "global_constraints": {
                "night_spacing_days_min": 1,
                "night_off": {"same_day": True, "next_day": False},
            },
        }
        days = [
            DayRow(dt.date(2026, 9, 7 + idx), ["Mon", "Tue", "Wed"][idx], idx + 2)
            for idx in range(3)
        ]
        slots = [
            Slot(d, f"{d.date}-J", ["J"], ["Zito", "A", "B", "C"], required=True, shift="Notte", rule_tag="J")
            for d in days
        ]
        prior_usage = {
            "counts": {"2026-09": {"Zito": {"J": 1}}},
            "night_dates_by_doc": {"2026-09": {"Zito": ["2026-09-02"]}},
        }

        assignment, stats = solve_with_ortools(cfg, days, slots, prior_usage=prior_usage)

        zito_current = sum(1 for d in days if assignment.get(f"{d.date}-J") == "Zito")
        self.assertLessEqual(zito_current, 1)
        self.assertNotEqual(stats["status"], "PARTIAL")

    def test_solve_across_uses_prior_generation_memory_for_night_spacing(self):
        cfg = {
            "rules": {
                "J": {"pool_other": ["Zito", "A"]},
                "K": {"pool": ["Zito", "A"]},
            },
            "global_constraints": {
                "night_spacing_days_min": 5,
                "night_off": {"same_day": True, "next_day": True},
            },
        }
        days = [
            DayRow(dt.date(2026, 9, 7), "Mon", 2),
            DayRow(dt.date(2026, 9, 8), "Tue", 3),
        ]
        prior_usage = {
            "counts": {"2026-09": {"Zito": {"J": 1}}},
            "night_dates_by_doc": {"2026-09": {"Zito": ["2026-09-05"]}},
        }

        slots, _assignment, _stats = solve_across_months(cfg, days, {}, prior_usage=prior_usage)
        j_slots = [s for s in slots if s.columns == ["J"]]

        self.assertTrue(j_slots)
        self.assertTrue(all("Zito" not in s.allowed for s in j_slots))

    def test_free_j_pool_balance_counts_prior_generation_memory(self):
        cfg = {
            "rules": {
                "J": {"pool_other": ["A", "B", "C"]},
            },
            "global_constraints": {
                "night_spacing_days_min": 1,
                "night_off": {"same_day": True, "next_day": False},
            },
        }
        days = [
            DayRow(dt.date(2026, 9, 7), "Mon", 2),
            DayRow(dt.date(2026, 9, 8), "Tue", 3),
        ]
        slots = [
            Slot(d, f"{d.date}-J", ["J"], ["A", "B", "C"], required=True, shift="Notte", rule_tag="J")
            for d in days
        ]
        prior_usage = {
            "counts": {"2026-09": {"A": {"J": 1}}},
            "night_dates_by_doc": {"2026-09": {"A": ["2026-09-02"]}},
        }

        assignment, stats = solve_with_ortools(cfg, days, slots, prior_usage=prior_usage)

        self.assertNotEqual(stats["status"], "PARTIAL")
        self.assertEqual(0, sum(1 for d in days if assignment.get(f"{d.date}-J") == "A"))

    def test_partial_period_weekday_afternoons_are_not_prior_festive_days(self):
        # A ha fatto 3 pomeriggi H feriali nella prima settimana salvata:
        # non sono festivi e non devono togliergli le domeniche del resto del mese.
        memory = gm.append_version(
            gm.empty_memory(),
            version_id="w1",
            label="Ottobre prima settimana",
            start_date=dt.date(2026, 10, 1),
            end_date=dt.date(2026, 10, 7),
            assignments={
                "2026-10-05": {"H": ["A"]},
                "2026-10-06": {"H": ["A"]},
                "2026-10-07": {"H": ["A"]},
            },
        )
        prior_usage = gm.build_solver_prior_usage(memory, dt.date(2026, 10, 8), dt.date(2026, 10, 31))
        cfg = {"rules": {"Festivi": {"pool": ["A", "B"]}}, "global_constraints": {}}
        days = [DayRow(dt.date(2026, 10, 11), "Sun", 2), DayRow(dt.date(2026, 10, 18), "Sun", 3)]
        slots = [
            Slot(d, f"{d.date}-DE", ["D", "E"], ["A", "B"], required=True, shift="Mattina", rule_tag="Festivo_DE")
            for d in days
        ]

        assignment, stats = solve_with_ortools(cfg, days, slots, prior_usage=prior_usage)

        self.assertNotEqual(stats["status"], "PARTIAL")
        self.assertEqual(sorted(assignment.values()), ["A", "B"])

    def test_previous_month_carryover_nights_do_not_count_as_current_month_weekend_usage(self):
        # Le notti di settembre servono solo per spacing/smonto a inizio ottobre:
        # non devono pesare sul bilanciamento weekend/feriali di ottobre.
        cfg = {
            "rules": {"J": {"pool_other": ["A", "B"]}},
            "global_constraints": {
                "night_spacing_days_min": 5,
                "night_off": {"same_day": True, "next_day": True},
            },
        }
        days = [DayRow(dt.date(2026, 10, 10), "Sat", 2), DayRow(dt.date(2026, 10, 19), "Mon", 3)]
        slots = [
            Slot(d, f"{d.date}-J", ["J"], ["A", "B"], required=True, shift="Notte", rule_tag="J")
            for d in days
        ]
        prior_usage = {"counts": {}, "night_dates_by_doc": {"2026-10": {"A": ["2026-09-26"]}}}
        # Unico criterio rimasto a parità di bilanciamento: A preferisce il sabato.
        prefs = [{"doctor": "A", "date": "2026-10-10", "shift": "Notte", "priority": "bassa"}]

        assignment, stats = solve_with_ortools(
            cfg, days, slots, availability_preferences=prefs, prior_usage=prior_usage
        )

        self.assertNotEqual(stats["status"], "PARTIAL")
        self.assertEqual(assignment["2026-10-10-J"], "A")

    def _twelve_working_days(self, start: dt.date):
        days = []
        d = start
        while len(days) < 12:
            if d.weekday() < 6:
                days.append(DayRow(d, ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"][d.weekday()], len(days) + 2))
            d += dt.timedelta(days=1)
        return days

    def test_q_cap_does_not_leave_coverable_slots_blank_when_pool_is_on_leave(self):
        # Pool Q di 6 medici ma 4 in ferie: A e B devono coprire tutto,
        # il tetto anti-dominanza non puo' produrre buchi evitabili.
        cfg = {"rules": {"Q": {"pool": ["A", "B", "C", "D", "E", "F"]}}, "global_constraints": {}}
        days = self._twelve_working_days(dt.date(2026, 10, 5))
        slots = [
            Slot(d, f"{d.date}-Q", ["Q"], ["A", "B"], required=True, shift="Mattina", rule_tag="Q")
            for d in days
        ]

        assignment, stats = solve_with_ortools(cfg, days, slots)

        self.assertNotEqual(stats["status"], "PARTIAL")
        self.assertTrue(all(assignment[s.slot_id] for s in slots))

    def test_t_cap_does_not_leave_coverable_slots_blank_when_pool_is_on_leave(self):
        cfg = {"rules": {"T": {"pool": ["A", "B", "C", "D", "E", "F"]}}, "global_constraints": {}}
        days = self._twelve_working_days(dt.date(2026, 10, 5))
        slots = [
            Slot(d, f"{d.date}-T", ["T"], ["A", "B"], required=True, shift="Mattina", rule_tag="T")
            for d in days
        ]

        assignment, stats = solve_with_ortools(cfg, days, slots)

        self.assertNotEqual(stats["status"], "PARTIAL")
        self.assertTrue(all(assignment[s.slot_id] for s in slots))

    def test_infeasible_fixed_assignments_are_reported_by_diagnostic_retry(self):
        cfg = {
            "rules": {"K": {"pool": ["A", "B"]}, "Q": {"pool": ["A", "B"]}},
            "global_constraints": {},
        }
        day = DayRow(dt.date(2026, 10, 5), "Mon", 2)
        slots = [
            Slot(day, "2026-10-05-K", ["K"], ["A", "B"], required=True, shift="Mattina", rule_tag="K"),
            Slot(day, "2026-10-05-Q", ["Q"], ["A", "B"], required=True, shift="Mattina", rule_tag="Q"),
        ]
        fixed = [
            {"doctor": "A", "date": "2026-10-05", "column": "K"},
            {"doctor": "A", "date": "2026-10-05", "column": "Q"},
        ]

        with self.assertRaises(RuntimeError) as ctx:
            solve_with_ortools(cfg, [day], slots, fixed_assignments=fixed)

        self.assertIn("fixed_assignments impossibili", str(ctx.exception))

    def test_diagnostic_retry_terminates_when_only_night_constraints_conflict(self):
        # J fisse incompatibili con lo spacing: il modello principale e' infeasible,
        # il retry diagnostico deve testare i gruppi e terminare (niente blocco
        # multi-thread di OR-Tools, vedi CLAUDE.md su num_search_workers).
        ctx = multiprocessing.get_context("spawn")
        queue = ctx.Queue()
        proc = ctx.Process(target=_run_conflicting_fixed_nights, args=(queue,))
        proc.start()
        proc.join(timeout=120)
        if proc.is_alive():
            proc.terminate()
            proc.join()
            self.fail("solve_with_ortools non termina nel retry diagnostico")
        self.assertIn("Slot critici", queue.get(timeout=5))


def _run_conflicting_fixed_nights(queue):
    cfg = {
        "rules": {"J": {"pool_other": ["A", "B"]}},
        "global_constraints": {
            "night_spacing_days_min": 5,
            "night_off": {"same_day": True, "next_day": False},
        },
    }
    days = [DayRow(dt.date(2026, 10, 5), "Mon", 2), DayRow(dt.date(2026, 10, 6), "Tue", 3)]
    slots = [
        Slot(d, f"{d.date}-J", ["J"], ["A", "B"], required=True, shift="Notte", rule_tag="J")
        for d in days
    ]
    fixed = [
        {"doctor": "A", "date": "2026-10-05", "column": "J"},
        {"doctor": "A", "date": "2026-10-06", "column": "J"},
    ]
    try:
        solve_with_ortools(cfg, days, slots, fixed_assignments=fixed)
        queue.put("NO_ERROR")
    except RuntimeError as e:
        queue.put(str(e))


if __name__ == "__main__":
    unittest.main()
