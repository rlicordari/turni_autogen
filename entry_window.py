"""Chiusura automatica dell'inserimento indisponibilità.

Il mese M si chiude il giorno ``close_day`` del mese precedente alle 00:00
(ora italiana): con il default 21, dicembre si chiude il 21 novembre alle 00:00
e l'ultimo giorno utile è il 20 novembre. I mesi successivi restano aperti.
``close_day`` 0 = chiusura automatica disattivata.
"""
from __future__ import annotations

from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

LOCAL_TZ = ZoneInfo("Europe/Rome")
DEFAULT_CLOSE_DAY = 21


def normalize_close_day(value) -> int:
    """0 = disattivata, altrimenti 1..28 (così esiste in ogni mese)."""
    try:
        day = int(value)
    except (TypeError, ValueError):
        return DEFAULT_CLOSE_DAY
    if day <= 0:
        return 0
    return min(day, 28)


def month_deadline(year: int, month: int, close_day) -> datetime | None:
    """Istante (ora italiana) in cui si chiude il mese ``year/month``."""
    day = normalize_close_day(close_day)
    if not day:
        return None
    prev = date(year, month, 1) - timedelta(days=1)
    return datetime(prev.year, prev.month, day, tzinfo=LOCAL_TZ)


def _local(now: datetime) -> datetime:
    return now.astimezone(LOCAL_TZ) if now.tzinfo else now.replace(tzinfo=LOCAL_TZ)


def month_is_open(year: int, month: int, now: datetime, close_day) -> bool:
    deadline = month_deadline(year, month, close_day)
    return deadline is None or _local(now) < deadline


def first_open_month(now: datetime, close_day) -> tuple[int, int]:
    """Primo mese, dal successivo in poi, ancora modificabile."""
    local = _local(now)
    y, m = local.year, local.month
    for _ in range(3):
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
        if month_is_open(y, m, local, close_day):
            break
    return y, m


def now_local() -> datetime:
    return datetime.now(LOCAL_TZ)
