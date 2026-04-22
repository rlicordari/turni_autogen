# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Descrizione del progetto

Generatore automatico di turni ospedalieri per UTIC/Cardiologia. Legge un template Excel, un file di regole YAML e un file di indisponibilità mensili; risolve l'assegnazione tramite **OR-Tools CP-SAT** (con fallback greedy); produce un file Excel compilato. Espone anche una **web UI Streamlit** per l'inserimento autonomo delle indisponibilità da parte dei medici e la generazione dei turni da parte dell'admin.

## Installazione

```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate

pip install -r requirements.txt
```

Richiede Python 3.10+ (consigliato 3.11).

## Avvio

**CLI:**
```bash
python turni_generator.py \
  --template Turni_Febbraio_2026.xlsx \
  --rules Regole_Turni.yml \
  --unavailability unavailability.xlsx \
  --out Turni_Febbraio_2026_COMPILATI.xlsx
```

**App Streamlit:**
```bash
streamlit run streamlit_app.py
```

**GUI legacy (tkinter):**
```bash
python turni_generator.py --gui
```

## Architettura

| File | Ruolo |
|---|---|
| `turni_generator.py` | Solver principale: legge template Excel + regole YAML + indisponibilità, costruisce il modello CP-SAT, scrive l'Excel di output |
| `streamlit_app.py` | UI Streamlit: login medico (PIN + OTP email), inserimento indisponibilità, generazione turni per admin |
| `unavailability_store.py` | Funzioni pure per il datastore CSV: parsing, filtraggio, deduplicazione, serializzazione |
| `github_utils.py` | Lettura/scrittura via GitHub Contents API (archivia il CSV delle indisponibilità su una repo privata) |
| `xlsx_utils.py` | Genera il file XLSX delle indisponibilità dal CSV usando `unavailability_template.xlsx` |
| `Regole_Turni.yml` | File regole mensile: definizione colonne, pool medici, quote, vincoli, penalità |
| `data/doctor_contacts.yml` | Mappa nome medico → email (usata per gli OTP) |
| `Style_Template.xlsx` | Template di stile opzionale applicato alla generazione di nuovi template Excel mensili |

### Flusso dei dati

1. Le **regole** sono definite in `Regole_Turni.yml` — le lettere di colonna corrispondono ai tipi di turno, ciascuno con pool, quote, vincoli di spaziatura e pesi di penalità.
2. Le **indisponibilità** sono archiviate come CSV in una repo GitHub privata (`data/unavailability_store.csv`). I medici le inseriscono via Streamlit; l'admin può anche caricare un file Excel.
3. Il **solver** (`turni_generator.py`) traduce regole + indisponibilità in variabili e vincoli CP-SAT, risolve e compila il workbook openpyxl.
4. **Streamlit** (`streamlit_app.py`) orchestra il tutto per gli utenti web: gestisce auth PIN, OTP via SMTP, lease di sessione per medico (kick-out in caso di login concorrente) e invoca `turni_generator` in-process.

### Secrets Streamlit (necessari per il funzionamento completo)

```toml
[auth]
admin_pin = "..."

[doctor_pins]
"Cognome" = "1234"

[github_unavailability]
token  = "ghp_..."
owner  = "REPO_OWNER"
repo   = "REPO_NAME"
branch = "main"
path   = "data/unavailability_store.csv"

[smtp]
host     = "smtp.gmail.com"
port     = 587
username = "..."
password = "APP_PASSWORD"
from     = "..."
starttls = true
```

### Keep-alive (GitHub Actions)

`.github/workflows/keep_awake_selenium.yml` fa un ping all'app Streamlit ogni 3 ore tramite Selenium per evitare l'hibernation. Richiede il secret Actions `STREAMLIT_APP_URL`.

## Convenzioni importanti

- I nomi dei medici in `Regole_Turni.yml` devono corrispondere **esattamente** ai nomi nei file di indisponibilità e in `doctor_contacts.yml`.
- Valori ammessi per `Fascia`: `Mattina`, `Pomeriggio`, `Notte`, `Diurno` (= Mattina + Pomeriggio), `Tutto il giorno`.
- Le lettere di colonna (C, D, E, …) nel YAML corrispondono direttamente alle colonne Excel del template.
- `absolute_exclusions` nel YAML elenca i medici mai assegnati ad alcun turno. Attualmente esclusi: De Luca, Carciotto, Virga, Andò, Saporito, D'Angelo.
- **D'Angelo** è stata esclusa temporaneamente (aprile 2026). Per reinserirla, aggiungere "D'Angelo" nei seguenti pool/liste in `Regole_Turni.yml`:
  - `E_G.allowed` (Cardiologia mattina / Riabilitazione)
  - `Q.pool` (ECO base)
  - `T.pool` (Interni)
  - `U.pool` (Contr.PM)
  - `Y.other_pool` (Ambulatori specialistici)
  - `Z.pool` (Vascolare)
  - `AB.fallback_pool` (Holter/Brugada/FA)
  - Rimuoverla da `absolute_exclusions` e da `C_reperibilita.excluded` (se applicabile al mese).
- I medici universitari (`university_doctors`) possono avere `night_counts_double: true` per dimezzare la quota effettiva di notti.
- La sezione `relief_valves` definisce fallback ad alta penalità per evitare l'infeasibility (es. permettere una colonna vuota a costo elevato anziché fallire).

## Feature in sviluppo: Memoria Storica Turni

**Piano completo:** `docs/PLAN_historical_shifts.md`

**Stato avanzamento (aggiornare ad ogni step):**
- [x] Task 1: `shift_history.py` — parser Excel definitivo + aggregazione stats + normalizzazione nomi (commit 2303447)
- [x] Task 2: Storage su GitHub — load/save storico JSON (commit a01a611)
- [x] Task 3: Integrazione solver — soft constraints con `historical_stats`
- [x] Task 4: UI admin Streamlit — upload, tabella, grafici Plotly, eliminazione mese
- [x] Task 5: Test end-to-end e push

**Moduli coinvolti:**
| File | Modifica |
|---|---|
| `shift_history.py` | NUOVO — parser + aggregazione + storage GitHub |
| `turni_generator.py` | Aggiunto parametro `historical_stats` al solver |
| `streamlit_app.py` | Nuova sezione admin "Memoria Storica" |
