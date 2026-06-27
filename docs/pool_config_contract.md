# Contratto Regole GUI/YAML

Questo documento definisce quali regole devono essere modificabili dalla GUI e
quali restano nello YAML come vincoli tecnici del solver.

## GUI: regole operative modificabili

La GUI `Gestione Pool Medici` e il file `data/pool_config.json` sono
autoritativi per:

- medici attivi/inattivi;
- email dei medici;
- colonne in cui ogni medico puo' comparire;
- abilitazione a reperibilita' C;
- esclusione dai turni diurni del sabato (mattina e pomeriggio, non notte);
- abilitazione a festivi diurni e notti festive;
- flag universitario;
- quote/limiti per singolo medico e colonna;
- target mensili globali per colonna;
- `counts_as` per il workload;
- combinazioni same-day esposte in GUI;
- servizi indispensabili e fallback di emergenza.

Se la GUI svuota un pool, il pool resta vuoto anche nel merge effettivo: il
solver non deve ripescare silenziosamente il vecchio pool YAML.

## YAML: vincoli tecnici avanzati

`Regole_Turni.yml` resta il template tecnico per:

- calendario e colonne del modello Excel;
- vincoli CP-SAT complessi e penalita' avanzate;
- regole speciali non ancora esposte in GUI;
- colonne automatiche o non operative;
- fallback tecnici di ultima istanza;
- parametri di debug e diagnostica solver.

Le regole YAML che limitano la sicurezza del dominio, come `J.never_in_J` e
`absolute_exclusions`, sono assolute: la GUI non puo' aggirarle.

## Regole particolari confermate

- `K+T` e' solo una valvola di emergenza. Il solver deve separare K e T quando
  esistono due medici disponibili.
- L'esclusione dal sabato vale solo per turni diurni (`Mattina`/`Pomeriggio`):
  non rimuove il medico da `J` notte e non rimuove la reperibilita' `C`.
- `J.never_in_J` prevale su pool, override e assegnazioni fisse.
- I servizi critici usano fallback solo quando il pool primario disponibile e'
  vuoto.
- Le colonne `AD:AG` sono solo riepilogo medici liberi.
- Le colonne automatiche `AA` e `AC` non sono gestite dal pool GUI.

## Validazione obbligatoria

Prima di salvare `pool_config.json`, la GUI deve eseguire l'audit in
`pool_config_store.audit_pool_config()` contro lo YAML corrente.

Errori bloccanti:

- medico in `J` se presente in `J.never_in_J`;
- colonne sconosciute o non esposte;
- pool primario vuoto su una colonna indispensabile;
- fallback verso medici sconosciuti o inattivi;
- nomi medico duplicati dopo normalizzazione;
- valori quota/target non validi.

Avvisi non bloccanti:

- pool vuoto su colonne non indispensabili;
- medico attivo senza colonne e senza reperibilita';
- spacing preferito minore dello spacing minimo.
