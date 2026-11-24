# -*- coding: utf-8 -*-
"""Doctor registry: one card per doctor, the ONLY source of who can do what.

Stored in the data repo as data/doctor_registry.json (schema 2). Every
per-person rule that used to live in Regole_Turni.yml pools, in hardcoded
names inside turni_generator.py or in pool_config.json now lives in the
doctor's own card:

    "Crea": {
      "active": true,
      "reperibilita": true,            # can do C
      "festivi_diurni": true,          # Sunday/holiday D+E and H+I
      "festivi_notti": true,           # J on Sundays/holidays
      "festivi_quota": null,           # holiday day shifts per month
      "no_sabato_diurno": false,
      "esente_weekend_liberi": false,  # exempt from "2 free weekends/month"
      "universitario": null,           # {"ratio": 0.6, "notte_doppia": true}
      "mai_notti": false,              # never J, not even from the holiday draw
      "lunedi_V_U_con": [],            # Monday on V -> U only one of these
      "services": {
        "EG": {"level": "pool"},
        "V":  {"level": "pool", "doppio": "titolare"},
        "AB": {"level": "pool", "sabato": true},
        "J":  {"level": "pool", "quota": {"type": "fixed", "value": 2},
               "notti_weekend": true, "weekend_min": null, "weekend_max": null},
        "W":  {"level": "pool", "giorni_fissi": ["Tue"]},
        "T":  {"level": "pool", "quota": {"type": "min", "value": 4, "target": 5}}
      },
      "note": ""
    }

Levels: "pool" = normally assigned there; "riserva" = only when nobody of the
pool is available that day. Saving touches one doctor at a time
(`set_doctor`), so editing a card can never change another doctor.

Pure logic, no Streamlit, no network (except the two thin GitHub helpers).
"""

from __future__ import annotations

import copy
import json
import re
from datetime import datetime, timezone
from typing import Optional

REGISTRY_PATH_DEFAULT = "data/doctor_registry.json"
SCHEMA_VERSION = 2
PLACEHOLDER = "Recupero"

# Services shown in a card, in display order: id -> (template columns, label).
# EG is one service: on weekdays E and G are always the same doctor.
SERVICES: dict[str, tuple[tuple[str, ...], str]] = {
    "D": (("D",), "UTIC mattina (coppia D/F)"),
    "F": (("F",), "Supporto 118"),
    "EG": (("E", "G"), "Cardiologia mattina + Riabilitazione"),
    "H": (("H",), "UTIC pomeriggio"),
    "I": (("I",), "Cardiologia pomeriggio"),
    "J": (("J",), "Notte"),
    "K": (("K",), "Letto"),
    "L": (("L",), "Padiglioni"),
    "Q": (("Q",), "ECO base"),
    "R": (("R",), "ECOSTRESS/ETE/TCD/Cardioncologia"),
    "S": (("S",), "Ecosala"),
    "T": (("T",), "Interni"),
    "U": (("U",), "Contr.PM"),
    "V": (("V",), "Sala PM"),
    "W": (("W",), "Ergometria/CPET"),
    "Y": (("Y",), "Ambulatori specialistici"),
    "Z": (("Z",), "Vascolare"),
    "AB": (("AB",), "Holter/Brugada/FA"),
    "AC": (("AC",), "Scintigrafia"),
}
COLUMN_TO_SERVICE = {c: sid for sid, (cols, _) in SERVICES.items() for c in cols}
LEVELS = ("pool", "riserva")
QUOTA_TYPES = ("fixed", "max", "min")
WEEKDAYS = ("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
DOPPIO_ROLES = ("titolare", "secondo")

_NAME_RE = re.compile(r"^[A-Za-zÀ-ÿ' .-]{2,40}$")


def norm_name(x) -> str:
    return re.sub(r"\s+", " ", str(x or "").strip())


def now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _sort_names(names) -> list[str]:
    return sorted(names, key=lambda s: (s == PLACEHOLDER, s.lower()))


# ── Cards ────────────────────────────────────────────────────────────────────

def empty_card() -> dict:
    return {
        "active": True,
        "reperibilita": True,
        "festivi_diurni": True,
        "festivi_notti": True,
        "festivi_quota": None,
        "no_sabato_diurno": False,
        "esente_weekend_liberi": False,
        "universitario": None,
        "mai_notti": False,
        "lunedi_V_U_con": [],
        "services": {},
        "note": "",
    }


def empty_registry() -> dict:
    return {
        "schema_version": SCHEMA_VERSION,
        "doctors": {},
        "column_settings": {},
        "service_combinations": [],
        "critical_services": {},
        "updated_at": None,
        "updated_by": None,
    }


def _int_or_none(v) -> Optional[int]:
    if v is None or v == "":
        return None
    try:
        return int(v)
    except (TypeError, ValueError):
        return None


def _clean_quota(q) -> Optional[dict]:
    if not isinstance(q, dict):
        return None
    qt = str(q.get("type") or "").strip().lower()
    val = _int_or_none(q.get("value"))
    if qt not in QUOTA_TYPES or val is None or val < 0:
        return None
    out = {"type": qt, "value": val}
    target = _int_or_none(q.get("target"))
    if target is not None and qt == "min":
        out["target"] = max(target, val)
    return out


def _clean_service(sid: str, spec) -> Optional[dict]:
    if not isinstance(spec, dict):
        return None
    level = str(spec.get("level") or "pool").strip().lower()
    if level not in LEVELS:
        return None
    out: dict = {"level": level}
    quota = _clean_quota(spec.get("quota"))
    if quota and level == "pool":
        out["quota"] = quota
    if sid == "J":
        out["notti_weekend"] = bool(spec.get("notti_weekend", True))
        for k in ("weekend_min", "weekend_max"):
            v = _int_or_none(spec.get(k))
            out[k] = v if v is not None and v >= 0 else None
    if sid == "V":
        role = spec.get("doppio")
        out["doppio"] = role if role in DOPPIO_ROLES else None
    if sid == "W":
        out["giorni_fissi"] = [d for d in WEEKDAYS if d in set(spec.get("giorni_fissi") or [])]
    if sid == "AB":
        out["sabato"] = bool(spec.get("sabato", False))
    return out


def clean_card(card) -> dict:
    """Card with every field present and valid (unknown fields dropped)."""
    base = empty_card()
    card = card if isinstance(card, dict) else {}
    for k in ("active", "reperibilita", "festivi_diurni", "festivi_notti",
              "no_sabato_diurno", "esente_weekend_liberi", "mai_notti"):
        if k in card:
            base[k] = bool(card[k])
    fq = _int_or_none(card.get("festivi_quota"))
    base["festivi_quota"] = fq if fq is not None and fq >= 0 else None
    uni = card.get("universitario")
    if isinstance(uni, dict):
        try:
            ratio = float(uni.get("ratio", 0.6))
        except (TypeError, ValueError):
            ratio = 0.6
        base["universitario"] = {
            "ratio": round(min(max(ratio, 0.05), 1.0), 2),
            "notte_doppia": bool(uni.get("notte_doppia", True)),
        }
    base["lunedi_V_U_con"] = sorted({norm_name(x) for x in (card.get("lunedi_V_U_con") or []) if norm_name(x)})
    services = {}
    for sid, spec in (card.get("services") or {}).items():
        if sid in SERVICES:
            cs = _clean_service(sid, spec)
            if cs:
                services[sid] = cs
    if base["mai_notti"]:
        services.pop("J", None)
    base["services"] = {sid: services[sid] for sid in SERVICES if sid in services}
    base["note"] = str(card.get("note") or "")[:500]
    return base


def normalize(reg) -> dict:
    reg = reg if isinstance(reg, dict) else {}
    out = empty_registry()
    out["doctors"] = {
        norm_name(name): clean_card(card)
        for name, card in (reg.get("doctors") or {}).items()
        if norm_name(name)
    }
    for k in ("column_settings", "critical_services"):
        if isinstance(reg.get(k), dict):
            out[k] = copy.deepcopy(reg[k])
    if isinstance(reg.get("service_combinations"), list):
        out["service_combinations"] = copy.deepcopy(reg["service_combinations"])
    out["updated_at"] = reg.get("updated_at")
    out["updated_by"] = reg.get("updated_by")
    return out


def to_text(reg: dict) -> str:
    return json.dumps(normalize(reg), ensure_ascii=False, indent=2) + "\n"


def from_text(text: str) -> dict:
    return normalize(json.loads(text or "{}"))


def valid_name(name: str) -> bool:
    return bool(_NAME_RE.match(norm_name(name)))


def set_doctor(reg: dict, name: str, card: dict, *, by: str = "admin") -> dict:
    """New registry with only `name`'s card replaced (or added)."""
    name = norm_name(name)
    if not valid_name(name):
        raise ValueError(f"Nome non valido: {name!r}")
    out = normalize(reg)
    out["doctors"][name] = clean_card(card)
    out["updated_at"] = now_iso()
    out["updated_by"] = by
    return out


def remove_doctor(reg: dict, name: str, *, by: str = "admin") -> dict:
    name = norm_name(name)
    out = normalize(reg)
    out["doctors"].pop(name, None)
    for card in out["doctors"].values():
        card["lunedi_V_U_con"] = [d for d in card["lunedi_V_U_con"] if d != name]
    out["updated_at"] = now_iso()
    out["updated_by"] = by
    return out


def set_column_settings(reg: dict, *, column_settings=None, service_combinations=None,
                        critical_services=None, by: str = "admin") -> dict:
    out = normalize(reg)
    if column_settings is not None:
        out["column_settings"] = copy.deepcopy(column_settings)
    if service_combinations is not None:
        out["service_combinations"] = copy.deepcopy(service_combinations)
    if critical_services is not None:
        out["critical_services"] = copy.deepcopy(critical_services)
    out["updated_at"] = now_iso()
    out["updated_by"] = by
    return out


def all_doctors(reg: dict) -> list[str]:
    return _sort_names(normalize(reg)["doctors"])


def active_doctors(reg: dict) -> list[str]:
    docs = normalize(reg)["doctors"]
    return _sort_names(d for d, c in docs.items() if c["active"])


def doctors_for(reg: dict, sid: str, level: str = "pool") -> list[str]:
    docs = normalize(reg)["doctors"]
    return _sort_names(
        d for d, c in docs.items()
        if c["active"] and (c["services"].get(sid) or {}).get("level") == level
    )


def check(reg: dict) -> list[str]:
    """Things the admin should know about (not blocking)."""
    reg = normalize(reg)
    docs = reg["doctors"]
    out = []
    for name, c in sorted(docs.items()):
        if not c["active"]:
            continue
        if name != PLACEHOLDER and not c["services"] and not c["reperibilita"]:
            out.append(f"{name}: attivo ma senza nessun servizio (non verrà mai assegnato).")
        elif name != PLACEHOLDER and not any(s["level"] == "pool" for s in c["services"].values()):
            out.append(f"{name}: attivo ma solo di riserva, nessun servizio normale.")
        for p in c["lunedi_V_U_con"]:
            if p not in docs:
                out.append(f"{name}: regola lunedì Sala PM/Contr.PM con un medico inesistente ({p}).")
    titolari_df = doctors_for(reg, "D")
    if titolari_df and len(titolari_df) != 2:
        out.append(
            f"UTIC mattina (D): la coppia D/F prevede 2 titolari, ora sono {len(titolari_df)} "
            f"({', '.join(titolari_df)}): l'alternanza 3+3 è disattivata."
        )
    v_tit = [d for d in doctors_for(reg, "V") if docs[d]["services"]["V"].get("doppio") == "titolare"]
    if len(v_tit) > 1:
        out.append(f"Sala PM: più di un titolare del turno doppio ({', '.join(v_tit)}), uso {v_tit[0]}.")
    w_tue = [d for d in doctors_for(reg, "W") if "Tue" in docs[d]["services"]["W"].get("giorni_fissi", [])]
    if len(w_tue) > 1:
        out.append(f"Ergometria: più medici fissi il martedì ({', '.join(w_tue)}), uso {w_tue[0]}.")
    ab_sat = [d for d in doctors_for(reg, "AB") if docs[d]["services"]["AB"].get("sabato")]
    if len(ab_sat) > 1:
        out.append(f"Holter: più medici per il sabato ({', '.join(ab_sat)}), uso {ab_sat[0]}.")
    return out


def overview(reg: dict) -> list[dict]:
    """One row per doctor: service id -> "●" pool / "○" riserva / ""."""
    reg = normalize(reg)
    rows = []
    for name in _sort_names(reg["doctors"]):
        c = reg["doctors"][name]
        row = {"Medico": name, "Attivo": "sì" if c["active"] else "no", "C": "●" if c["reperibilita"] else ""}
        for sid in SERVICES:
            lvl = (c["services"].get(sid) or {}).get("level")
            row[sid] = "●" if lvl == "pool" else ("○" if lvl == "riserva" else "")
        rows.append(row)
    return rows


# ── Migration from pool_config.json (v1) + Regole_Turni.yml ──────────────────

def _names(seq) -> list[str]:
    return [norm_name(x) for x in (seq or []) if norm_name(x)]


def j_quotas_names(rule_j: dict) -> list[str]:
    return _names((rule_j or {}).get("monthly_quotas") or {})


def migrate(pool_v1: Optional[dict], cfg_yaml: dict) -> dict:
    """Build the registry from the previous configuration, losing nothing.

    Pools come from pool_config.json when present (else from the YAML pools);
    every per-person rule hidden in the YAML or in the solver code becomes an
    explicit field of that doctor's card, including who the solver used as
    emergency cover ("riserva").
    """
    cfg_yaml = cfg_yaml or {}
    rules = cfg_yaml.get("rules") or {}
    gc = cfg_yaml.get("global_constraints") or {}
    pool_v1 = pool_v1 or {}
    v1_docs = {norm_name(k): v for k, v in (pool_v1.get("doctors") or {}).items()}
    abs_excl = set(_names(cfg_yaml.get("absolute_exclusions")))

    def r(key) -> dict:
        v = rules.get(key)
        return v if isinstance(v, dict) else {}

    yaml_pools = {
        "D": r("D_F").get("allowed"), "EG": r("E_G").get("allowed"),
        "H": r("H").get("pool_mon_fri"), "I": r("I").get("distribution_pool"),
        "J": r("J").get("pool_other"), "K": r("K").get("pool"), "L": r("L").get("pool_other"),
        "Q": r("Q").get("pool"), "R": r("R").get("pool"), "S": r("S").get("pool"),
        "T": r("T").get("pool"), "U": r("U").get("pool"), "V": r("V").get("pool"),
        "W": r("W").get("other_days_pool"), "Y": r("Y").get("other_pool"),
        "Z": r("Z").get("pool"), "AB": r("AB").get("fallback_pool"),
    }
    cols_by_doc: dict[str, set] = {}
    if v1_docs:
        for d, dc in v1_docs.items():
            cols = set()
            for c in dc.get("columns") or []:
                sid = COLUMN_TO_SERVICE.get(str(c).strip().upper())
                if sid:
                    cols.add(sid)
            cols_by_doc[d] = cols
    else:
        for sid, pool in yaml_pools.items():
            for d in _names(pool):
                cols_by_doc.setdefault(d, set()).add(sid)
        for d in _names(r("Festivi").get("pool")) + _names(gc.get("university_doctors")):
            cols_by_doc.setdefault(d, set())
    # Hidden solver behaviour made explicit
    for d in j_quotas_names(r("J")):
        cols_by_doc.setdefault(d, set()).add("J")  # J quota doctors joined J even outside pool_other
    ac_fixed = norm_name(r("AC").get("fixed"))
    if ac_fixed and "AC" in rules:
        cols_by_doc.setdefault(ac_fixed, set()).add("AC")
    if "L" in rules and PLACEHOLDER in cols_by_doc:
        cols_by_doc[PLACEHOLDER].add("L")  # L always accepted Recupero
    tue_fixed = norm_name(r("W").get("tuesday_fixed"))
    if tue_fixed and "W" in rules:
        cols_by_doc.setdefault(tue_fixed, set()).add("W")
    v_lead = norm_name(r("V").get("friday_required_doctor") or "Crea")
    v_partners = {"Dattilo", "Allegra"}
    ab_sat = norm_name(r("AB").get("saturday_only_doctor") or "Crea")
    ab_sat_on = int(r("AB").get("saturday_per_month", 0) or 0) > 0

    never_j = set(_names(r("J").get("never_in_J") or ["De Gregorio", "Manganaro"]))
    c_excl = set(_names(r("C_reperibilita").get("excluded")))
    fest_excl = set(_names(r("Festivi").get("excluded")))
    fest_quotas = {norm_name(k): _int_or_none(v) for k, v in (r("Festivi").get("quotas") or {}).items()}
    j_quotas = {norm_name(k): _int_or_none(v) for k, v in (r("J").get("monthly_quotas") or {}).items()}
    wk_min = {norm_name(k): _int_or_none(v) for k, v in (r("J").get("weekend_night_min_soft") or {}).items()}
    wk_max = {norm_name(k): _int_or_none(v) for k, v in (r("J").get("weekend_night_max_hard") or {}).items()}
    j_wex = set(_names(r("J").get("weekend_excluded_doctors")))
    uni_yaml = {norm_name(k): (v or {}) for k, v in (gc.get("university_doctors") or {}).items()}
    uni_ratio = float(gc.get("university_ratio", 0.6) or 0.6)
    wk_exempt = set(_names(gc.get("weekend_off_exempt")))
    cimino_u = _int_or_none(r("U").get("cimino_exact_per_month"))
    degregorio_i = _int_or_none(r("I").get("degregorio_max_weekdays"))
    df_titolari = {norm_name(r("D_F").get("pattern_doc1") or "Grimaldi"),
                   norm_name(r("D_F").get("pattern_doc2") or "Calabrò")}
    critical = pool_v1.get("critical_services") or {}

    reg = empty_registry()
    for name in sorted(set(cols_by_doc) | set(v1_docs)):
        dc = v1_docs.get(name) or {}
        card = empty_card()
        if dc:
            card["active"] = bool(dc.get("active", True))
            card["reperibilita"] = not bool(dc.get("excluded_from_reperibilita", False))
            card["festivi_diurni"] = bool(dc.get("festivi_diurni", True))
            card["festivi_notti"] = bool(dc.get("festivi_notti", True))
            card["no_sabato_diurno"] = bool(dc.get("exclude_saturday_day", False))
        else:
            card["reperibilita"] = name not in c_excl
            card["festivi_diurni"] = name not in fest_excl
        if name in abs_excl and not dc:
            # Without a pool_config card the YAML exclusion is all we know; with a
            # card, the admin's "Attivo" choice wins over the old YAML list.
            card["active"] = False
        if name == PLACEHOLDER:
            card.update(reperibilita=False, festivi_diurni=False, festivi_notti=False)
        card["festivi_quota"] = fest_quotas.get(name)
        card["esente_weekend_liberi"] = name in wk_exempt
        card["mai_notti"] = name in never_j
        uni = dc.get("university_doctor") if dc else None
        if isinstance(uni, dict) or name in uni_yaml:
            ratio = float((uni or {}).get("ratio", uni_ratio) or uni_ratio)
            nd = uni_yaml.get(name, {}).get("night_counts_double")
            card["universitario"] = {"ratio": ratio, "notte_doppia": bool(True if nd is None else nd)}

        services: dict = {}
        for sid in sorted(cols_by_doc.get(name, set())):
            if sid == "J" and card["mai_notti"]:
                continue
            services[sid] = {"level": "pool"}
        overrides = (dc.get("column_overrides") or {}) if dc else {}
        if "J" in services:
            jov = overrides.get("J") or {}
            q = _int_or_none(jov.get("monthly_quota"))
            if q is not None:
                services["J"]["quota"] = {"type": str(jov.get("quota_type") or "fixed"), "value": q}
            elif j_quotas.get(name) is not None:
                services["J"]["quota"] = {"type": "fixed", "value": j_quotas[name]}
            services["J"]["notti_weekend"] = bool(jov.get("weekend_nights", name not in j_wex))
            services["J"]["weekend_min"] = wk_min.get(name)
            services["J"]["weekend_max"] = wk_max.get(name)
        for col, ov in overrides.items():
            sid = COLUMN_TO_SERVICE.get(str(col).strip().upper())
            if sid and sid != "J" and sid in services and isinstance(ov, dict):
                q = _int_or_none(ov.get("monthly_quota"))
                if q is not None:
                    services[sid]["quota"] = {"type": str(ov.get("quota_type") or "fixed"), "value": q}
        if "V" in services:
            if name == v_lead:
                services["V"]["doppio"] = "titolare"
            elif name in v_partners:
                services["V"]["doppio"] = "secondo"
        if "AB" in services and name == ab_sat and ab_sat_on:
            services["AB"]["sabato"] = True
        if "W" in services and name == tue_fixed:
            services["W"]["giorni_fissi"] = ["Tue"]
        if "U" in services and name == "Cimino" and cimino_u:
            services["U"]["quota"] = {"type": "fixed", "value": cimino_u}
        if "I" in services and name == "De Gregorio" and degregorio_i:
            services["I"]["quota"] = {"type": "max", "value": degregorio_i}
        if name == PLACEHOLDER:
            t_min = _int_or_none(r("T").get("recupero_min_per_month"))
            if "T" in services and t_min is not None:
                services["T"]["quota"] = {
                    "type": "min", "value": t_min,
                    "target": _int_or_none(r("T").get("recupero_target_per_month")) or t_min,
                }
            for sid in ("Q", "W"):
                v = _int_or_none(r(sid).get("recupero_max_per_month"))
                if sid in services and v:
                    services[sid]["quota"] = {"type": "max", "value": v}
            v = _int_or_none(r("L").get("quota_recupero_per_month"))
            if "L" in services and v is not None:
                services["L"]["quota"] = {"type": "fixed", "value": v}
        if name == "Allegra" and r("U").get("v_allegra_monday_constraint"):
            card["lunedi_V_U_con"] = ["Crea", "Dattilo"]
        # Reserve: who the solver used as emergency cover before (critical
        # columns "any", H/E+G expansion, D/F last fallback), now in the card.
        if card["active"] and name != PLACEHOLDER:
            reserve = {"D", "F"}
            if name not in df_titolari:
                reserve |= {"EG", "H"}
                for col, spec in critical.items():
                    sid = COLUMN_TO_SERVICE.get(str(col).strip().upper())
                    fb = (spec or {}).get("fallback")
                    if sid and (fb == "any" or (isinstance(fb, list) and name in _names(fb))):
                        reserve.add(sid)
            for sid in reserve:
                if sid == "J" and card["mai_notti"]:
                    continue
                services.setdefault(sid, {"level": "riserva"})
        card["services"] = services
        reg["doctors"][name] = clean_card(card)

    reg["column_settings"] = copy.deepcopy(pool_v1.get("column_settings") or {})
    reg["service_combinations"] = copy.deepcopy(pool_v1.get("service_combinations") or [])
    reg["critical_services"] = copy.deepcopy(critical)
    reg["updated_at"] = now_iso()
    reg["updated_by"] = "migrazione"
    return normalize(reg)


# ── Solver configuration ─────────────────────────────────────────────────────

# Every YAML field that names doctors. They are emptied before the registry is
# applied: no doctor can come from anywhere else.
_NAME_LIST_FIELDS = (
    "allowed", "pool", "pool_other", "pool_mon_fri", "other_pool", "fallback_pool",
    "distribution_pool", "other_days_pool", "excluded", "never_in_J",
    "weekend_excluded_doctors",
)
_NAME_DICT_FIELDS = ("quotas", "monthly_quotas", "weekend_night_min_soft", "weekend_night_max_hard")
_NAME_SCALAR_FIELDS = (
    "friday_required_doctor", "tuesday_fixed", "saturday_only_doctor", "fixed", "prefer",
    "pattern_doc1", "pattern_doc2",
)
_OBSOLETE_FIELDS = (
    "cimino_exact_per_month", "degregorio_max_weekdays", "v_allegra_monday_constraint",
    "recupero_min_per_month", "recupero_target_per_month", "recupero_max_per_month",
    "quota_recupero_per_month",
)


def strip_names(cfg_yaml: dict) -> dict:
    """Copy of the YAML rules with every doctor name removed."""
    cfg = copy.deepcopy(cfg_yaml or {})
    for rule in (cfg.get("rules") or {}).values():
        if not isinstance(rule, dict):
            continue
        for k in _NAME_LIST_FIELDS:
            if k in rule:
                rule[k] = []
        for k in _NAME_DICT_FIELDS:
            if k in rule:
                rule[k] = {}
        for k in _NAME_SCALAR_FIELDS + _OBSOLETE_FIELDS:
            rule.pop(k, None)
    gc = cfg.setdefault("global_constraints", {})
    gc["university_doctors"] = {}
    gc["weekend_off_exempt"] = []
    cfg["absolute_exclusions"] = []
    cfg.pop("unavailability", None)
    return cfg


def to_solver_cfg(cfg_yaml: dict, reg: dict) -> dict:
    """Solver configuration: YAML column rules + doctors ONLY from the registry."""
    reg = normalize(reg)
    cfg = strip_names(cfg_yaml)
    rules = cfg.setdefault("rules", {})
    gc = cfg["global_constraints"]
    docs = reg["doctors"]
    active = active_doctors(reg)
    cfg["doctors"] = list(active)
    cfg["absolute_exclusions"] = _sort_names(d for d, c in docs.items() if not c["active"])

    def pool(sid: str) -> list[str]:
        return doctors_for(reg, sid, "pool")

    def svc(d, sid) -> dict:
        return docs[d]["services"].get(sid) or {}

    def rule(key: str) -> dict:
        return rules[key]

    if "D_F" in rules:
        df = rule("D_F")
        titolari = pool("D")
        df["allowed"] = titolari
        df["f_pool"] = pool("F")
        if len(titolari) == 2:
            df["pattern_doc1"], df["pattern_doc2"] = titolari
        else:
            df["pattern_3_3"] = False
    if "E_G" in rules:
        rule("E_G")["allowed"] = pool("EG")
    if "H" in rules:
        rule("H")["pool_mon_fri"] = pool("H")
        rule("H")["distribution_pool"] = pool("H")
    if "I" in rules:
        rule("I")["distribution_pool"] = pool("I")
    if "J" in rules:
        rj = rule("J")
        j_pool = pool("J")
        rj["pool_other"] = j_pool
        rj["never_in_J"] = [d for d in active if docs[d]["mai_notti"]]
        rj["monthly_quotas"] = {
            d: svc(d, "J")["quota"]["value"] for d in j_pool
            if (svc(d, "J").get("quota") or {}).get("type") == "fixed"
        }
        rj["weekend_excluded_doctors"] = [d for d in j_pool if not svc(d, "J").get("notti_weekend", True)]
        rj["weekend_night_min_soft"] = {
            d: svc(d, "J")["weekend_min"] for d in j_pool if svc(d, "J").get("weekend_min") is not None
        }
        rj["weekend_night_max_hard"] = {
            d: svc(d, "J")["weekend_max"] for d in j_pool if svc(d, "J").get("weekend_max") is not None
        }
    for sid, (key, field) in {
        "K": ("K", "pool"), "L": ("L", "pool_other"), "Q": ("Q", "pool"), "R": ("R", "pool"),
        "S": ("S", "pool"), "T": ("T", "pool"), "U": ("U", "pool"), "V": ("V", "pool"),
        "W": ("W", "other_days_pool"), "Y": ("Y", "other_pool"), "Z": ("Z", "pool"),
        "AB": ("AB", "fallback_pool"), "AC": ("AC", "pool"),
    }.items():
        if key in rules:
            rule(key)[field] = pool(sid)
    if "V" in rules:
        v_pool = pool("V")
        lead = [d for d in v_pool if svc(d, "V").get("doppio") == "titolare"]
        if lead:
            rule("V")["friday_required_doctor"] = lead[0]
        rule("V")["double_partners"] = [d for d in v_pool if svc(d, "V").get("doppio") == "secondo"]
    if "W" in rules:
        tue = [d for d in pool("W") if "Tue" in svc(d, "W").get("giorni_fissi", [])]
        if tue:
            rule("W")["tuesday_fixed"] = tue[0]
    if "AB" in rules:
        sat = [d for d in pool("AB") if svc(d, "AB").get("sabato")]
        if sat:
            rule("AB")["saturday_only_doctor"] = sat[0]
        else:
            rule("AB")["saturday_per_month"] = 0
    if "U" in rules:
        rule("U")["monday_v_rules"] = [
            {"v_doctor": d, "u_doctors": [p for p in docs[d]["lunedi_V_U_con"] if p in active]}
            for d in active if docs[d]["lunedi_V_U_con"]
        ]
    fest = rules.setdefault("Festivi", {})
    fest["pool"] = [d for d in active if d != PLACEHOLDER and docs[d]["festivi_diurni"]]
    fest["excluded"] = [d for d in active if d == PLACEHOLDER or not docs[d]["festivi_diurni"]]
    fest["quotas"] = {d: docs[d]["festivi_quota"] for d in active if docs[d]["festivi_quota"] is not None}
    if "C_reperibilita" in rules:
        rule("C_reperibilita")["excluded"] = _sort_names(
            [d for d in active if not docs[d]["reperibilita"]] + cfg["absolute_exclusions"]
        )
    cfg["pool_festivi_notti_excluded"] = {d for d in pool("J") if not docs[d]["festivi_notti"]}
    cfg["pool_saturday_day_excluded"] = {d for d in active if docs[d]["no_sabato_diurno"]}
    gc["weekend_off_exempt"] = [d for d in active if docs[d]["esente_weekend_liberi"]]
    gc["university_doctors"] = {
        d: {
            "type": "university",
            "ratio": docs[d]["universitario"]["ratio"],
            "night_counts_double": bool(docs[d]["universitario"]["notte_doppia"]),
        }
        for d in active if docs[d]["universitario"]
    }

    # Quotas other than J fixed (generic per doctor/column) + placeholder rules
    quota_overrides: dict = {}
    for d in active:
        for sid, spec in docs[d]["services"].items():
            q = spec.get("quota")
            if not q or spec.get("level") != "pool":
                continue
            if sid == "J" and q["type"] == "fixed":
                continue
            if d == PLACEHOLDER:
                if sid == "T" and q["type"] == "min" and "T" in rules:
                    rules["T"]["recupero_min_per_month"] = q["value"]
                    rules["T"]["recupero_target_per_month"] = q.get("target", q["value"])
                    continue
                if sid in ("Q", "W") and q["type"] == "max" and sid in rules:
                    rules[sid]["recupero_max_per_month"] = q["value"]
                    continue
                if sid == "L" and q["type"] == "fixed" and "L" in rules:
                    rules["L"]["quota_recupero_per_month"] = q["value"]
                    continue
            col = SERVICES[sid][0][0]
            quota_overrides[(d, col)] = {"type": q["type"], "value": int(q["value"])}
    cfg["pool_quota_overrides"] = quota_overrides

    # Reserve doctors per template column
    reserve_pools: dict[str, list[str]] = {}
    for sid, (cols, _) in SERVICES.items():
        rs = doctors_for(reg, sid, "riserva")
        if rs:
            for c in cols:
                reserve_pools[c] = rs
    cfg["reserve_pools"] = reserve_pools

    # Column-level settings (not per person)
    col_settings = reg.get("column_settings") or {}
    counts_as, targets = {}, {}
    for col, cs in col_settings.items():
        if not isinstance(cs, dict):
            continue
        if "counts_as" in cs:
            counts_as[col] = int(cs["counts_as"])
        if cs.get("monthly_target") is not None:
            targets[str(col).strip().upper()] = int(cs["monthly_target"])
    if counts_as:
        cfg["pool_counts_as"] = counts_as
    if targets:
        cfg["pool_monthly_targets"] = targets
    j_cs = col_settings.get("J") or {}
    if "spacing_min_days" in j_cs:
        gc["night_spacing_days_min"] = int(j_cs["spacing_min_days"])
    if "spacing_preferred_days" in j_cs:
        gc["night_spacing_days_preferred"] = int(j_cs["spacing_preferred_days"])
    combos = reg.get("service_combinations") or []
    cfg["pool_service_combinations"] = copy.deepcopy(combos)
    relief = gc.setdefault("relief_valves", {})
    for combo in combos:
        if tuple(sorted(combo.get("columns") or [])) == ("K", "T"):
            relief["enable_kt_share"] = combo.get("mode", "fallback") in ("fallback", "always")
    cfg["pool_critical_services"] = copy.deepcopy(reg.get("critical_services") or {})
    return cfg


def strings_in(obj) -> set[str]:
    """Every string anywhere in obj (keys and values): used to prove no names leak."""
    out = set()
    if isinstance(obj, dict):
        for k, v in obj.items():
            out |= strings_in(k) | strings_in(v)
    elif isinstance(obj, (list, tuple, set)):
        for v in obj:
            out |= strings_in(v)
    elif isinstance(obj, str):
        out.add(obj)
    return out


# ── GitHub I/O (thin) ────────────────────────────────────────────────────────

def load_from_github(owner, repo, token, branch="main", path=REGISTRY_PATH_DEFAULT):
    """(registry, sha), or (None, None) when the file does not exist yet."""
    from github_utils import get_file
    gf = get_file(owner, repo, path, token, branch)
    if gf is None:
        return None, None
    return from_text(gf.text), gf.sha


def save_to_github(reg, owner, repo, token, branch="main", sha=None,
                   path=REGISTRY_PATH_DEFAULT, message="Scheda medici aggiornata"):
    from github_utils import put_file
    return put_file(owner, repo, path, token, message, to_text(reg), branch, sha)
