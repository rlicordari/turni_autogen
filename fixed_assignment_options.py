"""Validate fixed assignments returned by the Streamlit data editor."""

from datetime import date, datetime

import pandas as pd


def _cell_text(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def _parse_day(value: object) -> date | None:
    if value is None or pd.isna(value):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        raw = value.strip()
        for parser in (date.fromisoformat, lambda s: datetime.strptime(s, "%d/%m/%Y").date()):
            try:
                return parser(raw)
            except ValueError:
                pass
    return None


def parse_fixed_assignment_rows(
    rows: pd.DataFrame,
    *,
    start: date,
    end: date,
    doctors: list[str],
    shift_labels: dict[str, str],
) -> list[dict[str, str]]:
    """Parse all editor rows or reject the whole form with a useful error.

    Streamlit returns ISO date strings when a row is added to an initially empty
    editor. Existing rows may instead contain date or pandas Timestamp values.
    """
    if not isinstance(rows, pd.DataFrame):
        raise ValueError("La tabella delle assegnazioni fisse non è leggibile.")

    result: list[dict[str, str]] = []
    seen: dict[tuple[str, str], str] = {}
    for index, (_, row) in enumerate(rows.iterrows(), start=1):
        doctor = _cell_text(row.get("Medico"))
        raw_day = row.get("Giorno")
        day = _parse_day(raw_day)
        label = _cell_text(row.get("Turno/Colonna"))
        if not doctor and not _cell_text(raw_day) and not label:
            continue
        if not doctor or not _cell_text(raw_day) or not label:
            raise ValueError(f"Riga {index}: completa Medico, Giorno e Turno/Colonna.")
        if doctor not in doctors:
            raise ValueError(f"Riga {index}: medico non riconosciuto: {doctor}.")
        if day is None:
            raise ValueError(f"Riga {index}: data non valida: {_cell_text(raw_day)}.")
        if not start <= day <= end:
            raise ValueError(f"Riga {index}: {day:%d/%m/%Y} è fuori dal periodo selezionato.")
        if label not in shift_labels:
            raise ValueError(f"Riga {index}: turno non riconosciuto: {label}.")

        column = shift_labels[label]
        key = (day.isoformat(), column)
        previous = seen.get(key)
        if previous and previous != doctor:
            raise ValueError(
                f"Riga {index}: {day:%d/%m/%Y} {column} è già assegnato a {previous}."
            )
        if previous:
            continue
        seen[key] = doctor
        result.append({"doctor": doctor, "date": day.isoformat(), "column": column})
    return result
