# -*- coding: utf-8 -*-
"""Pool configuration store — overlay JSON su Regole_Turni.yml.

Gestisce load/save/validate/migrate del file data/pool_config.json su GitHub.
Struttura analoga a shift_history.py (storage JSON) e unavailability_store.py
(funzioni pure parse/validate).

Il JSON sovrascrive pool, quote e flag al momento della generazione turni.
Il YAML rimane invariato come template avanzato.
"""

from __future__ import annotations

import copy
import json
from datetime import datetime, timezone
from typing import Optional

POOL_CONFIG_PATH_DEFAULT = "data/pool_config.json"
SCHEMA_VERSION = 1

QUOTA_TYPES = {"fixed", "max", "min"}
COMBINATION_MODES = {"always", "fallback", "preferred"}

_DOCTOR_REQUIRED_KEYS = {
    "active",
    "columns",
    "festivi_diurni",
    "festivi_notti",
    "excluded_from_reperibilita",
    "university_doctor",
    "column_overrides",
}


# ── 1. Serializzazione pura ──────────────────────────────────────────────────

def pool_config_to_text(cfg: dict) -> str:
    return json.dumps(cfg, ensure_ascii=False, indent=2)


def pool_config_from_text(text: str) -> dict:
    return json.loads(text)


# ── 2. Skeleton vuoto ────────────────────────────────────────────────────────

def empty_pool_config() -> dict:
    return {
        "schema_version": SCHEMA_VERSION,
        "doctors": {},
        "column_settings": {
            "J": {
                "monthly_target": 2,
                "spacing_min_days": 5,
                "spacing_preferred_days": 7,
                "counts_as": 2,
            },
            "C": {
                "monthly_target": None,
                "spacing_min_days": 3,
                "counts_as": 0,
            },
        },
        "service_combinations": [],
        "critical_services": {},
        "updated_at": _now_iso(),
        "updated_by": "admin",
    }


# ── 3. Validazione ───────────────────────────────────────────────────────────

def validate_pool_config(cfg: dict) -> list[str]:
    """Ritorna lista di errori (vuota se OK)."""
    errs: list[str] = []

    if not isinstance(cfg, dict):
        return ["La configurazione non è un dizionario valido"]

    if cfg.get("schema_version") != SCHEMA_VERSION:
        errs.append(f"schema_version deve essere {SCHEMA_VERSION}, trovato: {cfg.get('schema_version')}")

    doctors = cfg.get("doctors", {})
    if not isinstance(doctors, dict):
        errs.append("'doctors' deve essere un dizionario")
    else:
        for name, dcfg in doctors.items():
            if not isinstance(dcfg, dict):
                errs.append(f"Medico '{name}': deve essere un dizionario")
                continue
            missing = _DOCTOR_REQUIRED_KEYS - dcfg.keys()
            if missing:
                errs.append(f"Medico '{name}': campi mancanti: {sorted(missing)}")
            if not isinstance(dcfg.get("columns", []), list):
                errs.append(f"Medico '{name}': 'columns' deve essere una lista")
            overrides = dcfg.get("column_overrides", {})
            if not isinstance(overrides, dict):
                errs.append(f"Medico '{name}': 'column_overrides' deve essere un dizionario")
            else:
                for col, ov in overrides.items():
                    if not isinstance(ov, dict):
                        errs.append(f"Medico '{name}', colonna '{col}': override deve essere un dizionario")
                        continue
                    qt = ov.get("quota_type")
                    if qt is not None and qt not in QUOTA_TYPES:
                        errs.append(f"Medico '{name}', colonna '{col}': quota_type '{qt}' non valido (ammessi: {sorted(QUOTA_TYPES)})")
                    mq = ov.get("monthly_quota")
                    if mq is not None and (not isinstance(mq, int) or mq < 0):
                        errs.append(f"Medico '{name}', colonna '{col}': monthly_quota deve essere intero >= 0")

    combos = cfg.get("service_combinations", [])
    if not isinstance(combos, list):
        errs.append("'service_combinations' deve essere una lista")
    else:
        for i, combo in enumerate(combos):
            if not isinstance(combo, dict):
                errs.append(f"service_combinations[{i}]: deve essere un dizionario")
                continue
            cols = combo.get("columns", [])
            if not isinstance(cols, list) or len(cols) != 2:
                errs.append(f"service_combinations[{i}]: 'columns' deve essere una lista di 2 lettere")
            mode = combo.get("mode")
            if mode not in COMBINATION_MODES:
                errs.append(f"service_combinations[{i}]: mode '{mode}' non valido (ammessi: {sorted(COMBINATION_MODES)})")

    critical = cfg.get("critical_services", {})
    if not isinstance(critical, dict):
        errs.append("'critical_services' deve essere un dizionario")
    else:
        for col, spec in critical.items():
            if not isinstance(spec, dict):
                errs.append(f"critical_services['{col}']: deve essere un dizionario")
                continue
            fb = spec.get("fallback")
            if fb != "any" and not isinstance(fb, list):
                errs.append(f"critical_services['{col}']: fallback deve essere 'any' o una lista di medici")

    col_settings = cfg.get("column_settings", {})
    if not isinstance(col_settings, dict):
        errs.append("'column_settings' deve essere un dizionario")
    else:
        c_cfg = col_settings.get("C", {})
        if isinstance(c_cfg, dict):
            ca = c_cfg.get("counts_as")
            if ca is not None and ca != 0:
                errs.append("column_settings.C.counts_as deve essere 0 (reperibilità non conta nel workload)")

    return errs


# ── 4. GitHub storage ────────────────────────────────────────────────────────

def load_pool_config_from_github(
    owner: str,
    repo: str,
    token: str,
    branch: str = "main",
    path: str = POOL_CONFIG_PATH_DEFAULT,
) -> tuple[dict, Optional[str]]:
    """Carica pool_config da GitHub. Ritorna ({}, None) se il file non esiste."""
    from github_utils import get_file

    gf = get_file(owner, repo, path, token, branch)
    if gf is None:
        return {}, None
    return pool_config_from_text(gf.text), gf.sha


def save_pool_config_to_github(
    cfg: dict,
    owner: str,
    repo: str,
    token: str,
    branch: str = "main",
    sha: Optional[str] = None,
    path: str = POOL_CONFIG_PATH_DEFAULT,
) -> dict:
    """Salva pool_config su GitHub. Ritorna risposta GitHub Contents API."""
    from github_utils import put_file

    text = pool_config_to_text(cfg)
    return put_file(
        owner,
        repo,
        path,
        token,
        "Aggiornamento configurazione pool medici",
        text,
        branch,
        sha,
    )


# ── 5. Migrazione da YAML ────────────────────────────────────────────────────

def migrate_from_yaml(cfg_yaml: dict) -> dict:
    """Costruisce un pool_config iniziale leggendo il YAML corrente.

    L'admin rifinisce i dettagli dopo la migrazione — questo è un punto di
    partenza automatico che rispecchia la configurazione attuale.
    """
    from turni_generator import collect_doctors, norm_name

    rules = cfg_yaml.get("rules", {})
    gc = cfg_yaml.get("global_constraints", {})
    abs_excl = {norm_name(x) for x in (cfg_yaml.get("absolute_exclusions") or [])}

    rJ = rules.get("J", {})
    rC = rules.get("C_reperibilita", {})
    rFest = rules.get("Festivi", {})

    j_pool_other = {norm_name(x) for x in (rJ.get("pool_other") or [])}
    j_monthly_quotas = rJ.get("monthly_quotas") or {}
    j_weekend_excluded = {norm_name(x) for x in (rJ.get("weekend_excluded_doctors") or [])}

    c_excluded = {norm_name(x) for x in (rC.get("excluded") or [])}

    fest_excl = {norm_name(x) for x in (rFest.get("excluded") or [])}
    fest_pool = {norm_name(x) for x in (rFest.get("pool") or [])}

    gc_uni = gc.get("university_doctors") or {}
    uni_ratio = float(gc.get("university_ratio", 0.6))

    all_doctors = collect_doctors(cfg_yaml)

    # Mappa colonna → chiave pool nel YAML e lista medici
    col_to_doctors = _build_col_to_doctors(rules, all_doctors)

    doctors_cfg: dict = {}
    for doc in all_doctors:
        dn = norm_name(doc)
        active = dn not in abs_excl

        columns = sorted(
            col for col, pool in col_to_doctors.items() if dn in pool
        )

        festivi_diurni = dn not in fest_excl and (dn in fest_pool or active)
        festivi_notti = dn in j_pool_other and dn not in fest_excl

        excluded_from_rep = dn in c_excluded

        uni_cfg = gc_uni.get(doc) or gc_uni.get(dn)
        university_doctor = {"ratio": uni_ratio} if uni_cfg else None

        column_overrides: dict = {}
        mq = j_monthly_quotas.get(doc) or j_monthly_quotas.get(dn)
        if mq is not None:
            column_overrides["J"] = {"monthly_quota": int(mq), "quota_type": "fixed"}
        if dn in j_weekend_excluded:
            column_overrides.setdefault("J", {})["weekend_nights"] = False

        doctors_cfg[doc] = {
            "active": active,
            "columns": columns,
            "festivi_diurni": festivi_diurni,
            "festivi_notti": festivi_notti,
            "excluded_from_reperibilita": excluded_from_rep,
            "university_doctor": university_doctor,
            "column_overrides": column_overrides,
        }

    # Aggiungi medici in absolute_exclusions come active=false (non in collect_doctors)
    for doc in (cfg_yaml.get("absolute_exclusions") or []):
        if doc not in doctors_cfg:
            doctors_cfg[doc] = {
                "active": False,
                "columns": [],
                "festivi_diurni": False,
                "festivi_notti": False,
                "excluded_from_reperibilita": True,
                "university_doctor": None,
                "column_overrides": {},
            }

    relief = gc.get("relief_valves") or {}
    service_combinations = []
    if relief.get("enable_kt_share", False):
        service_combinations.append({"columns": ["K", "T"], "same_day": True, "mode": "fallback"})

    cfg = empty_pool_config()
    cfg["doctors"] = doctors_cfg
    cfg["service_combinations"] = service_combinations
    cfg["updated_at"] = _now_iso()
    return cfg


# ── Helpers interni ──────────────────────────────────────────────────────────

def _now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _build_col_to_doctors(rules: dict, all_doctors: list[str]) -> dict[str, set[str]]:
    """Costruisce mappa colonna → set di nomi normalizzati dal YAML."""
    from turni_generator import norm_name

    POOL_KEYS = [
        "allowed", "pool", "pool_other", "pool_mon_fri", "other_pool",
        "fallback_pool", "distribution_pool", "other_days_pool",
    ]
    FIXED_KEYS = ["fixed", "tuesday_fixed", "friday_required_doctor"]

    col_to_docs: dict[str, set[str]] = {}

    col_rule_map = {
        "C": "C_reperibilita",
        "D": "D_F", "F": "D_F",
        "E": "E_G", "G": "E_G",
        "H": "H", "I": "I",
        "J": "J",
        "K": "K", "L": "L", "Q": "Q", "R": "R", "S": "S",
        "T": "T", "U": "U", "V": "V", "W": "W",
        "Y": "Y", "Z": "Z", "AA": "AA", "AB": "AB", "AC": "AC",
    }

    dn_all = {norm_name(d) for d in all_doctors}

    for col, rule_key in col_rule_map.items():
        rule = rules.get(rule_key, {})
        if not isinstance(rule, dict):
            continue
        pool: set[str] = set()
        for k in POOL_KEYS:
            if k in rule and isinstance(rule[k], list):
                pool |= {norm_name(x) for x in rule[k] if x}
        for k in FIXED_KEYS:
            if rule.get(k):
                pool.add(norm_name(rule[k]))
        # Includi solo medici riconosciuti da collect_doctors
        col_to_docs[col] = pool & dn_all

    return col_to_docs
