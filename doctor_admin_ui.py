# -*- coding: utf-8 -*-
"""Admin page "Medici": one card per doctor, saved one doctor at a time.

Why it is built this way (18/11/2026: the columns chosen for D'Angelo were
saved on Allegra):
- every widget key contains the doctor's name and a version number, so a
  value typed for one doctor can never be read back for another;
- the card is edited inside a form and saved with `update_registry`, which
  re-reads the registry from GitHub, refuses if that doctor was changed in the
  meantime and replaces ONLY that doctor's card;
- the services table has a fixed set of rows (no row can move or disappear).

The page receives its I/O as callables so it has no GitHub code of its own.
"""

from __future__ import annotations

from typing import Callable, Optional

import pandas as pd
import streamlit as st

import doctor_registry as dr

LEVEL_LABELS = {"": "—", "pool": "Sì", "riserva": "Riserva"}
LEVEL_FROM_LABEL = {v: k for k, v in LEVEL_LABELS.items()}
QUOTA_LABELS = {"": "—", "fixed": "esatta", "max": "massimo", "min": "minimo"}
QUOTA_FROM_LABEL = {v: k for k, v in QUOTA_LABELS.items()}
WEEKDAY_IT = {"Mon": "Lunedì", "Tue": "Martedì", "Wed": "Mercoledì", "Thu": "Giovedì",
              "Fri": "Venerdì", "Sat": "Sabato", "Sun": "Domenica"}
NEW_DOCTOR = "➕ Nuovo medico"


class CardChanged(Exception):
    """The card was changed elsewhere after this page loaded it."""


def services_table(card: dict) -> pd.DataFrame:
    rows = []
    for sid, (cols, label) in dr.SERVICES.items():
        spec = card["services"].get(sid) or {}
        q = spec.get("quota") or {}
        rows.append({
            "Servizio": f"{'+'.join(cols)} · {label}",
            "Assegnabile": LEVEL_LABELS.get(spec.get("level", ""), "—"),
            "Quota mensile": QUOTA_LABELS.get(q.get("type", ""), "—"),
            "N": q.get("value"),
            "Obiettivo (solo minimo)": q.get("target"),
        })
    return pd.DataFrame(rows)


def _num(v) -> Optional[int]:
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    try:
        return int(v)
    except (TypeError, ValueError):
        return None


def card_from_form(base: dict, table: pd.DataFrame, values: dict) -> dict:
    """New card from the form widgets; per-service options survive level changes."""
    card = dr.clean_card(base)
    services: dict = {}
    for (sid, _), (_, row) in zip(dr.SERVICES.items(), table.iterrows()):
        level = LEVEL_FROM_LABEL.get(str(row.get("Assegnabile") or "—"), "")
        if not level:
            continue
        spec = dict(base["services"].get(sid) or {})
        spec["level"] = level
        qtype = QUOTA_FROM_LABEL.get(str(row.get("Quota mensile") or "—"), "")
        qval = _num(row.get("N"))
        if qtype and qval is not None and level == "pool":
            spec["quota"] = {"type": qtype, "value": qval}
            target = _num(row.get("Obiettivo (solo minimo)"))
            if qtype == "min" and target is not None:
                spec["quota"]["target"] = target
        else:
            spec.pop("quota", None)
        services[sid] = spec
    if "J" in services:
        services["J"]["notti_weekend"] = bool(values["j_weekend"])
        services["J"]["weekend_min"] = _num(values["j_wmin"])
        services["J"]["weekend_max"] = _num(values["j_wmax"])
    if "V" in services:
        services["V"]["doppio"] = {"—": None, "titolare": "titolare", "secondo": "secondo"}[values["v_doppio"]]
    if "W" in services:
        services["W"]["giorni_fissi"] = list(values["w_days"])
    if "AB" in services:
        services["AB"]["sabato"] = bool(values["ab_sat"])
    card.update(
        active=bool(values["active"]),
        reperibilita=bool(values["reperibilita"]),
        festivi_diurni=bool(values["festivi_diurni"]),
        festivi_notti=bool(values["festivi_notti"]),
        festivi_quota=_num(values["festivi_quota"]),
        no_sabato_diurno=bool(values["no_sabato"]),
        esente_weekend_liberi=bool(values["esente_weekend"]),
        mai_notti=bool(values["mai_notti"]),
        universitario=({"ratio": float(values["uni_ratio"]), "notte_doppia": bool(values["uni_notte"])}
                       if values["universitario"] else None),
        lunedi_V_U_con=list(values["lun_vu"]),
        note=str(values["note"] or ""),
        services=services,
    )
    return dr.clean_card(card)


def replace_card_if_unchanged(reg: dict, name: str, base: Optional[dict], new_card: dict) -> dict:
    """Registry with `name`'s card replaced, or CardChanged if it moved meanwhile."""
    current = dr.normalize(reg)["doctors"].get(name)
    if base is None:
        if current is not None:
            raise CardChanged(f"Esiste già un medico chiamato {name}.")
    elif current != dr.clean_card(base):
        raise CardChanged(f"La scheda di {name} è stata modificata da un'altra sessione.")
    return dr.set_doctor(reg, name, new_card)


def render(
    *,
    load_registry: Callable[[], tuple],
    update_registry: Callable[[Callable[[dict], dict], str], dict],
    contacts: dict,
    save_email: Callable[[str, str], None],
) -> None:
    st.markdown("### 👩‍⚕️ Medici")
    st.caption(
        "Ogni medico ha la sua scheda: **solo da qui** il generatore sa chi può fare cosa. "
        "Si salva un medico alla volta: modificare una scheda non tocca mai gli altri."
    )
    flash = st.session_state.pop("_doc_flash", None)
    if flash:
        (st.success if flash[0] == "success" else st.error)(flash[1])

    reg, _sha, on_github = load_registry()
    reg = dr.normalize(reg)
    ver = int(st.session_state.get("doc_ver", 0))

    if not on_github:
        st.info(
            "Le schede non sono ancora nel repository dati: le vedi ricostruite dalla configurazione "
            "attuale (pool medici e file regole), con ogni regola personale riportata nella scheda. "
            "Controllale; al primo salvataggio vengono scritte tutte.",
            icon="🧭",
        )
        if st.button("💾 Salva tutte le schede così come sono", key="doc_save_all"):
            try:
                update_registry(lambda fresh: dr.normalize(reg), "Schede medici create dalla configurazione precedente")
                st.session_state["_doc_flash"] = ("success", "Schede medici salvate nel repository dati.")
            except Exception as e:
                st.session_state["_doc_flash"] = ("error", f"Schede non salvate: {e}")
            st.rerun()

    warnings = dr.check(reg)
    if warnings:
        st.warning("Da controllare:\n\n" + "\n".join(f"- {w}" for w in warnings), icon="⚠️")

    tab_card, tab_all, tab_cols = st.tabs(["👩‍⚕️ Scheda medico", "📋 Riepilogo", "⚙️ Impostazioni servizi"])

    with tab_card:
        names = dr.all_doctors(reg)
        pending = st.session_state.pop("_doc_sel_next", None)
        if pending is not None:
            st.session_state["doc_sel"] = pending if pending in names else (names[0] if names else NEW_DOCTOR)
        if st.session_state.get("doc_sel") not in names + [NEW_DOCTOR]:
            st.session_state.pop("doc_sel", None)
        sel = st.selectbox("Medico", names + [NEW_DOCTOR], key="doc_sel")
        if sel == NEW_DOCTOR:
            _render_new_doctor(reg, names, update_registry, ver)
        elif sel:
            _render_card(reg, sel, contacts, update_registry, save_email, ver)

    with tab_all:
        st.caption("● = assegnabile · ○ = solo di riserva (quando nessun altro è disponibile)")
        legend = ", ".join(f"{sid} = {label}" for sid, (_, label) in dr.SERVICES.items())
        st.dataframe(pd.DataFrame(dr.overview(reg)), hide_index=True, width="stretch")
        st.caption(legend)

    with tab_cols:
        _render_column_settings(reg, update_registry, ver)


def _render_new_doctor(reg, names, update_registry, ver):
    with st.form(f"doc_new::{ver}"):
        name = st.text_input("Cognome", key=f"doc_new_name::{ver}", help="Esattamente come nelle indisponibilità")
        copy_from = st.selectbox("Copia servizi e regole da (facoltativo)", ["—"] + names, key=f"doc_new_copy::{ver}")
        ok = st.form_submit_button("➕ Crea medico", key=f"doc_new_ok::{ver}")
    if ok:
        name = dr.norm_name(name)
        if not dr.valid_name(name):
            st.error("Nome non valido: usa solo lettere, spazi e apostrofo.")
            return
        base_card = dr.empty_card() if copy_from == "—" else dict(reg["doctors"][copy_from], note="")
        try:
            update_registry(lambda fresh: replace_card_if_unchanged(fresh, name, None, base_card),
                            f"Nuovo medico: {name}")
            st.session_state["_doc_flash"] = ("success", f"{name} creato. Completa la sua scheda.")
            st.session_state["_doc_sel_next"] = name
            st.session_state["doc_ver"] = ver + 1
        except Exception as e:
            st.session_state["_doc_flash"] = ("error", f"{name} non creato: {e}")
        st.rerun()


def _render_card(reg, name, contacts, update_registry, save_email, ver):
    base = reg["doctors"][name]
    k = f"{name}::{ver}"
    others = [d for d in dr.all_doctors(reg) if d not in (name, dr.PLACEHOLDER)]
    with st.form(f"doc_form::{k}"):
        st.markdown(f"#### {name}")
        c1, c2, c3 = st.columns(3)
        with c1:
            active = st.checkbox("Attivo", value=base["active"], key=f"d_active::{k}",
                                 help="Non attivo (malattia, maternità…): mai assegnato, i suoi turni sorteggiati restano vuoti e gialli.")
            reperibilita = st.checkbox("Reperibilità (C)", value=base["reperibilita"], key=f"d_rep::{k}")
            no_sabato = st.checkbox("No sabato mattina/pomeriggio", value=base["no_sabato_diurno"], key=f"d_nosat::{k}")
        with c2:
            festivi_diurni = st.checkbox("Festivi e domeniche di giorno", value=base["festivi_diurni"], key=f"d_fd::{k}")
            festivi_notti = st.checkbox("Notti nei festivi", value=base["festivi_notti"], key=f"d_fn::{k}")
            festivi_quota = st.number_input(
                "Festivi di giorno al mese (vuoto = nessuna quota)", min_value=0, max_value=10,
                value=base["festivi_quota"], step=1, key=f"d_fq::{k}",
                help="Quota fissa di domeniche/festivi diurni al mese. Vuoto = li distribuisce il bilanciamento.",
            )
        with c3:
            mai_notti = st.checkbox("Mai notti (nemmeno da sorteggio)", value=base["mai_notti"], key=f"d_never::{k}")
            esente_weekend = st.checkbox("Esente dai 2 weekend liberi al mese", value=base["esente_weekend_liberi"], key=f"d_wkex::{k}")
            universitario = st.checkbox("Universitario", value=bool(base["universitario"]), key=f"d_uni::{k}")
            uni = base["universitario"] or {"ratio": 0.6, "notte_doppia": True}
            uni_ratio = st.number_input("Quota oraria universitario", min_value=0.1, max_value=1.0,
                                        value=float(uni["ratio"]), step=0.05, key=f"d_unir::{k}")
            uni_notte = st.checkbox("La notte vale doppio (universitario)", value=bool(uni["notte_doppia"]), key=f"d_unin::{k}")
        email = st.text_input("Email (codici PIN e resoconti)",
                              value=str((contacts.get(name) or {}).get("email") or ""), key=f"d_email::{k}")

        st.markdown("**Servizi** — *Sì* = assegnabile; *Riserva* = solo se nessun altro è disponibile quel giorno")
        table = st.data_editor(
            services_table(base),
            column_config={
                "Servizio": st.column_config.TextColumn("Servizio", disabled=True, width="large"),
                "Assegnabile": st.column_config.SelectboxColumn("Assegnabile", options=list(LEVEL_LABELS.values()), required=True),
                "Quota mensile": st.column_config.SelectboxColumn("Quota mensile", options=list(QUOTA_LABELS.values()), required=True),
                "N": st.column_config.NumberColumn("N", min_value=0, max_value=31, step=1),
                "Obiettivo (solo minimo)": st.column_config.NumberColumn("Obiettivo", min_value=0, max_value=31, step=1),
            },
            hide_index=True, num_rows="fixed", width="stretch", key=f"d_svc::{k}",
        )

        st.markdown("**Regole particolari** (valgono solo se il servizio è assegnabile)")
        r1, r2, r3 = st.columns(3)
        j = base["services"].get("J") or {}
        with r1:
            j_weekend = st.checkbox("Notti anche sabato/domenica", value=j.get("notti_weekend", True), key=f"d_jw::{k}")
            j_wmin = st.number_input("Notti weekend: minimo (vuoto = nessuno)", min_value=0, max_value=8,
                                     value=j.get("weekend_min"), step=1, key=f"d_jwmin::{k}")
            j_wmax = st.number_input("Notti weekend: massimo (vuoto = nessuno)", min_value=0, max_value=8,
                                     value=j.get("weekend_max"), step=1, key=f"d_jwmax::{k}")
        with r2:
            v_role = (base["services"].get("V") or {}).get("doppio") or "—"
            v_doppio = st.selectbox("Sala PM, turno doppio", ["—", "titolare", "secondo"],
                                    index=["—", "titolare", "secondo"].index(v_role), key=f"d_vd::{k}",
                                    help="Titolare: c'è sempre nel turno doppio. Secondo: può affiancarlo.")
            lun_vu = st.multiselect("Lunedì in Sala PM → Contr.PM solo a", others,
                                    default=[d for d in base["lunedi_V_U_con"] if d in others], key=f"d_lvu::{k}")
        with r3:
            w_days = st.multiselect("Ergometria fisso il", list(dr.WEEKDAYS[:5]),
                                    default=[d for d in (base["services"].get("W") or {}).get("giorni_fissi", [])
                                             if d in dr.WEEKDAYS[:5]],
                                    format_func=lambda d: WEEKDAY_IT[d], key=f"d_wd::{k}")
            ab_sat = st.checkbox("Holter anche il sabato", value=(base["services"].get("AB") or {}).get("sabato", False),
                                 key=f"d_abs::{k}")
        note = st.text_area("Note", value=base["note"], key=f"d_note::{k}", height=68)
        submitted = st.form_submit_button(f"💾 Salva scheda di {name}", type="primary", key=f"d_save::{k}")

    if submitted:
        values = dict(
            active=active, reperibilita=reperibilita, no_sabato=no_sabato, festivi_diurni=festivi_diurni,
            festivi_notti=festivi_notti, festivi_quota=festivi_quota, mai_notti=mai_notti,
            esente_weekend=esente_weekend, universitario=universitario, uni_ratio=uni_ratio, uni_notte=uni_notte,
            j_weekend=j_weekend, j_wmin=j_wmin, j_wmax=j_wmax, v_doppio=v_doppio, lun_vu=lun_vu,
            w_days=w_days, ab_sat=ab_sat, note=note,
        )
        new_card = card_from_form(base, table, values)
        msgs = []
        try:
            if new_card != base:
                update_registry(lambda fresh: replace_card_if_unchanged(fresh, name, base, new_card),
                                f"Scheda medico aggiornata: {name}")
                msgs.append(f"Scheda di {name} salvata.")
            old_email = str((contacts.get(name) or {}).get("email") or "")
            if email.strip() != old_email:
                save_email(name, email)
                msgs.append("Email aggiornata." if email.strip() else "Email rimossa.")
            st.session_state["_doc_flash"] = ("success", " ".join(msgs) or "Nessuna modifica da salvare.")
            st.session_state["doc_ver"] = ver + 1
        except CardChanged as e:
            st.session_state["_doc_flash"] = ("error", f"{e} Niente è stato sovrascritto: ricarica e ripeti.")
            st.session_state["doc_ver"] = ver + 1
        except ValueError as e:
            st.session_state["_doc_flash"] = ("error", f"{e}: niente è stato salvato.")
        except Exception as e:
            st.session_state["_doc_flash"] = ("error", f"Scheda di {name} non salvata ({type(e).__name__}: {e}).")
        st.rerun()

    with st.expander(f"🗑️ Elimina {name}"):
        st.caption("Per un'assenza lunga basta togliere *Attivo*: la scheda resta pronta per il rientro.")
        sure = st.checkbox(f"Sì, elimina definitivamente la scheda di {name}", key=f"d_delok::{k}")
        if st.button("Elimina", key=f"d_del::{k}", disabled=not sure):
            def _remove(fresh):
                if dr.normalize(fresh)["doctors"].get(name) != base:
                    raise CardChanged(f"La scheda di {name} è stata modificata da un'altra sessione.")
                return dr.remove_doctor(fresh, name)
            try:
                update_registry(_remove, f"Medico eliminato: {name}")
                st.session_state["_doc_flash"] = ("success", f"Scheda di {name} eliminata.")
                st.session_state["_doc_sel_next"] = ""
            except Exception as e:
                st.session_state["_doc_flash"] = ("error", f"{name} non eliminato: {e}")
            st.session_state["doc_ver"] = ver + 1
            st.rerun()


def _render_column_settings(reg, update_registry, ver):
    st.caption("Impostazioni per servizio, uguali per tutti (chi può farle si decide nelle schede).")
    cs = reg["column_settings"]
    cols = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "Q", "R", "S", "T", "U", "V", "W", "Y", "Z", "AB"]
    df = pd.DataFrame([{
        "Colonna": c,
        "Target mensile": (cs.get(c) or {}).get("monthly_target"),
        "Distanza minima (gg)": int((cs.get(c) or {}).get("spacing_min_days", 0) or 0),
        "Distanza preferita (gg)": int((cs.get(c) or {}).get("spacing_preferred_days", 0) or 0),
        "Vale come": int((cs.get(c) or {}).get("counts_as", 0 if c == "C" else (2 if c == "J" else 1))),
    } for c in cols])
    crit_now = sorted(reg["critical_services"])
    with st.form(f"doc_cols::{ver}"):
        edited = st.data_editor(
            df, hide_index=True, num_rows="fixed", width="stretch", key=f"d_cs::{ver}",
            column_config={
                "Colonna": st.column_config.TextColumn("Colonna", disabled=True),
                "Target mensile": st.column_config.NumberColumn("Target", min_value=0, max_value=31, step=1),
                "Distanza minima (gg)": st.column_config.NumberColumn("Distanza min", min_value=0, max_value=30, step=1),
                "Distanza preferita (gg)": st.column_config.NumberColumn("Distanza pref.", min_value=0, max_value=30, step=1),
                "Vale come": st.column_config.NumberColumn("Vale come", min_value=0, max_value=4, step=1),
            },
        )
        kt = next((c for c in reg["service_combinations"] if sorted(c.get("columns") or []) == ["K", "T"]), None)
        kt_on = st.checkbox("Letto e Interni (K+T) allo stesso medico se manca personale", value=bool(kt), key=f"d_kt::{ver}")
        critical = st.multiselect(
            "Colonne indispensabili (se restano vuote la cella è gialla e lo segnalo)",
            cols, default=[c for c in crit_now if c in cols], key=f"d_crit::{ver}",
        )
        ok = st.form_submit_button("💾 Salva impostazioni servizi", key=f"d_cs_ok::{ver}")
    if ok:
        new_cs = {}
        for _, row in edited.iterrows():
            c = row["Colonna"]
            new_cs[c] = {
                "monthly_target": _num(row["Target mensile"]),
                "spacing_min_days": _num(row["Distanza minima (gg)"]) or 0,
                "spacing_preferred_days": _num(row["Distanza preferita (gg)"]) or 0,
                "counts_as": 0 if c == "C" else (_num(row["Vale come"]) if _num(row["Vale come"]) is not None else 1),
            }
        combos = [c for c in reg["service_combinations"] if sorted(c.get("columns") or []) != ["K", "T"]]
        if kt_on:
            combos.append({"columns": ["K", "T"], "same_day": True, "mode": "fallback"})
        crit = {c: (reg["critical_services"].get(c) or {"fallback": "any"}) for c in critical}
        base = (reg["column_settings"], reg["service_combinations"], reg["critical_services"])

        def _apply(fresh):
            f = dr.normalize(fresh)
            if (f["column_settings"], f["service_combinations"], f["critical_services"]) != base:
                raise CardChanged("Le impostazioni dei servizi sono state modificate da un'altra sessione.")
            return dr.set_column_settings(f, column_settings=new_cs, service_combinations=combos, critical_services=crit)
        try:
            update_registry(_apply, "Impostazioni servizi aggiornate")
            st.session_state["_doc_flash"] = ("success", "Impostazioni servizi salvate.")
        except Exception as e:
            st.session_state["_doc_flash"] = ("error", f"Impostazioni non salvate: {e}")
        st.session_state["doc_ver"] = ver + 1
        st.rerun()
