import io
import tempfile
from pathlib import Path

import streamlit as st
from openpyxl import load_workbook

from turni_generator import generate_schedule

st.set_page_config(page_title="Turni Autogenerator", layout="wide")

st.title("Turni Autogenerator – versione web (Streamlit)")

st.markdown(
    """
Questa versione gira **senza Tkinter** (Streamlit Cloud non supporta GUI desktop).

**Workflow:** carichi *Template Excel*, *Regole YAML* (o usi quelle del repo) e, opzionalmente, *Indisponibilità* → generi un **.xlsx** scaricabile.
"""
)

# --- Inputs ---
col1, col2 = st.columns(2)

with col1:
    template_up = st.file_uploader("Template turni (.xlsx)", type=["xlsx"], accept_multiple_files=False)
    # Evita input libero del nome foglio: è la causa più comune di KeyError su Streamlit Cloud.
    # Leggiamo i nomi reali dei fogli dal template e facciamo scegliere da dropdown.
    sheet_name = None
    if template_up is not None:
        try:
            wb_tmp = load_workbook(io.BytesIO(template_up.getvalue()), read_only=True, data_only=True)
            sheets = wb_tmp.sheetnames
        except Exception:
            sheets = []

        if sheets:
            st.caption("Fogli trovati nel template: " + ", ".join(sheets))
            opt = st.selectbox(
                "Seleziona foglio",
                options=["(foglio attivo / primo foglio)"] + sheets,
                index=0,
            )
            sheet_name = None if opt.startswith("(") else opt
        else:
            st.warning(
                "Non riesco a leggere i fogli del template: verrà usato il foglio attivo (primo foglio)."
            )
            sheet_name = None

with col2:
    use_repo_rules = st.checkbox("Usa Regole_Turni.yml del repo", value=True)
    rules_up = None
    if not use_repo_rules:
        rules_up = st.file_uploader("Regole (.yml/.yaml)", type=["yml", "yaml"], accept_multiple_files=False)

unav_up = st.file_uploader(
    "Indisponibilità (opzionale: .xlsx/.csv/.tsv)",
    type=["xlsx", "xls", "csv", "tsv"],
    accept_multiple_files=False,
)

st.divider()

run_btn = st.button("Genera turni", type="primary")

if run_btn:
    if template_up is None:
        st.error("Carica prima il **Template turni (.xlsx)**.")
        st.stop()

    if not use_repo_rules and rules_up is None:
        st.error("Hai disattivato 'Usa Regole_Turni.yml del repo': carica un file **Regole (.yml/.yaml)**.")
        st.stop()

    with tempfile.TemporaryDirectory() as td:
        td = Path(td)

        # Save template
        template_path = td / "template.xlsx"
        template_path.write_bytes(template_up.getvalue())

        # Save rules
        if use_repo_rules:
            repo_rules = Path(__file__).with_name("Regole_Turni.yml")
            if not repo_rules.exists():
                st.error("Non trovo 'Regole_Turni.yml' nel repo. Carica un file regole manualmente.")
                st.stop()
            rules_path = td / "Regole_Turni.yml"
            rules_path.write_bytes(repo_rules.read_bytes())
        else:
            rules_path = td / "Regole_Turni.yml"
            rules_path.write_bytes(rules_up.getvalue())

        # Save unavailability (optional)
        unav_path = None
        if unav_up is not None:
            unav_path = td / f"unavailability.{unav_up.name.split('.')[-1]}"
            unav_path.write_bytes(unav_up.getvalue())

        out_path = td / "turni_output.xlsx"

        with st.spinner("Calcolo in corso…"):
            stats, log_path = generate_schedule(
                template_xlsx=template_path,
                rules_yml=rules_path,
                out_xlsx=out_path,
                unavailability_path=unav_path,
	            sheet_name=sheet_name,
            )

        # --- Results ---
        st.success("Turni generati.")

        # Show solver summary
        status = (stats or {}).get("status", "")
        st.subheader("Esito solver")
        st.write(f"**Status:** {status}")

        months = (stats or {}).get("months") or {}
        greedy_months = [k for k, v in months.items() if isinstance(v, dict) and v.get("solver_error")]
        if greedy_months:
            st.warning(
                "OR-Tools non disponibile o schedule infeasible per: "
                + ", ".join(greedy_months)
                + ". In quei mesi è stato usato il greedy."
            )

        # Download output
        st.download_button(
            label="Scarica Excel generato",
            data=out_path.read_bytes(),
            file_name="turni_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Show log (if any)
        if log_path and Path(log_path).exists():
            st.subheader("Log")
            try:
                st.code(Path(log_path).read_text(encoding="utf-8", errors="ignore"))
            except Exception:
                st.code(Path(log_path).read_text(errors="ignore"))
