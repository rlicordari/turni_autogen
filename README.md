# Turni Autogenerator (UTIC/Cardiologia) – prototype multipiattaforma

Questo progetto legge:
- un **template Excel** dei turni (come `Turni_Febbraio_2026.xlsx`)
- un file **regole YAML** (come `Regole_Turni_UTIC_FebMar_2026_v3.yml`)
- un file **indisponibilità mensili** (Excel/CSV) con colonne: `Medico`, `Data`, `Fascia`

e genera un nuovo Excel compilato, rispettando vincoli e quote.



## Indisponibilità – logica del file (Excel/CSV)
Il file delle indisponibilità usa le colonne:

- **Medico**: deve combaciare (match esatto) con i nomi usati nelle regole
- **Data**: data Excel o `gg/mm/aaaa`
- **Fascia**: uno tra:
  - `Mattina` → blocca tutte le colonne mattutine
  - `Pomeriggio` → blocca tutte le colonne pomeridiane
  - `Notte` → blocca la colonna notte
  - `Diurno` (o `Giorno`) → blocca **Mattina + Pomeriggio** (ma lascia disponibile per la Notte)
  - `Tutto il giorno` (o `All day`) → blocca l’intera giornata (equivalente a full-day)

Note:
- Più righe per lo stesso medico e giorno si sommano (es. Mattina + Pomeriggio).
- Le colonne **AD/AE (Medici liberi)** possono rimanere vuote; se vengono valorizzate,
  contengono solo medici **non assegnati** e **senza indisponibilità** in quel giorno.
## Requisiti
- Python 3.10+ (consigliato 3.11)
- Windows/macOS/Linux

### Installazione (consigliato con venv)
```bash
python -m venv .venv
# macOS/Linux:
source .venv/bin/activate
# Windows:
# .venv\Scripts\activate

pip install -r requirements.txt
```

> Nota macOS: se vedi un prompt per “Command Line Developer Tools”, esegui:
> `xcode-select --install`

## Uso (CLI)
```bash
python turni_generator.py \
  --template Turni_Febbraio_2026.xlsx \
  --rules Regole_Turni_UTIC_FebMar_2026_v3.yml \
  --unavailability unavailability.xlsx \
  --out Turni_Febbraio_2026_COMPILATI.xlsx
```

## Uso (GUI)
```bash
python turni_generator.py --gui
```

## Indisponibilità mensili: formato atteso
- **Medico**: es. `Calabrò`
- **Data**: `gg/mm/aaaa` oppure data Excel
- **Fascia**: `Mattina` / `Pomeriggio` / `Notte`

Esempio:
| Medico | Data | Fascia | Note |
|---|---|---|---|
| Calabrò | 04/02/2026 | Mattina | congresso |

## Solver
Il solver principale usa **OR-Tools (CP-SAT)**.
Se OR-Tools non è disponibile, il programma prova un riempimento greedy e genera un report di conflitti.

## Output
- Compila le colonne operative definite nel YAML
- Mantiene formattazione del template
- Compila `medici liberi 1/2` (AD/AE) con i primi 2 medici non assegnati in giornata

---
Prototype: pensato per essere esteso (più colonne, priorità, pesi obiettivo, ecc.).

## Indisponibilità con privacy (nome + PIN)

L'app include una sezione "Indisponibilità (Medico)" dove ogni medico può inserire SOLO le proprie indisponibilità.
I dati vengono salvati in un file CSV su una **repo GitHub privata** (Contents API). I medici non possono vedere le indisponibilità altrui.

### Secrets (Streamlit Cloud)

Esempio `secrets.toml`:

```toml
[auth]
admin_pin = "ADMIN123"

[doctor_pins]
"Dattilo" = "1111"
"Migliorato" = "2222"
"Calabrò" = "3333"

[github_unavailability]
token  = "ghp_xxxxxxxxxxxxxxxxxxxxx"
owner  = "TUO_OWNER"
repo   = "TUO_REPO_PRIVATA"
branch = "main"
path   = "data/unavailability_store.csv"
```

Note:
- `token` deve avere permessi di lettura/scrittura sui contenuti della repo privata.
- `path` può essere un percorso dentro la repo (es. `data/...`). Il file verrà creato/aggiornato dalla app.
- In alternativa puoi usare chiavi piatte: `ADMIN_PIN`, `GITHUB_UNAV_TOKEN`, `GITHUB_UNAV_OWNER`, `GITHUB_UNAV_REPO`, `GITHUB_UNAV_BRANCH`, `GITHUB_UNAV_PATH`.

### Formato indisponibilità

Il formato Excel usato dall'admin segue `unavailability_template.xlsx` (foglio `Indisponibilita`) con colonne:
`Medico | Data | Fascia | Note`.

---

## PIN personalizzabile per medico (OTP via Email)

L’app supporta PIN **per-medico** (salvati come hash PBKDF2 + salt su GitHub) con:
- **impostazione iniziale / reset** tramite OTP inviato via **Email (SMTP)**
- **cambio PIN** in autonomia quando il medico è loggato
- **kick-out automatico**: se lo stesso medico accede da un altro browser/dispositivo, la sessione precedente viene disconnessa.

### File contatti medici
Aggiungere nella repo GitHub:
- `data/doctor_contacts.yml`

Esempio:
```yaml
Rossi Mario:
  email: "rossi.mario@gmail.com"
Bianchi Luca:
  email: "luca.bianchi@gmail.com"
```

### Secrets (Gmail SMTP)
In Streamlit Cloud → **Settings → Secrets** inserire (vedi anche `.streamlit/secrets.example.toml`):
```toml
[smtp]
host = "smtp.gmail.com"
port = 587
username = "tuoaccount@gmail.com"
password = "APP_PASSWORD_16_CARATTERI"
from = "tuoaccount@gmail.com"
starttls = true
```
> La `password` deve essere una **App Password** (richiede 2FA sul Google Account).

---

## Streamlit Community Cloud: evitare l'hibernation

Su Streamlit Community Cloud le app vanno in *hibernation* dopo un periodo senza traffico (policy della piattaforma).
Per ridurre al minimo lo stop, puoi usare un **ping periodico**:

- Workflow GitHub Actions pronto: `.github/workflows/keep_awake.yml`
- Aggiungi in GitHub → Settings → Secrets and variables → Actions il secret:
  - `STREAMLIT_APP_URL` = URL pubblico della tua app (es. `https://nomeapp.streamlit.app/`)

