"""Shared event-date mapping helpers for WZDB generation and matching."""

from __future__ import annotations

import hashlib
import json
from datetime import date
from pathlib import Path
from typing import Any, Iterable


DEFAULT_EVENT_DATES = Path("data/event_dates_2026.json")


def event_mapping_key(season: str, values: Iterable[Any], ordinal: int) -> str:
    """Identify one physical event without relying on mutable row indexes."""
    payload = json.dumps(
        [
            str(season),
            *[" ".join(str(value or "").split()) for value in values],
            ordinal,
        ],
        ensure_ascii=False,
        separators=(",", ":"),
    ).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


def load_event_date_mapping(path: Path | None) -> dict[str, str]:
    """Load and validate the approved mapping from event key to ISO date."""
    if path is None or not path.is_file():
        return {}
    payload = json.loads(path.read_text(encoding="utf-8"))
    raw_events = payload.get("events", {})
    if not isinstance(raw_events, dict):
        raise ValueError(f"Nieprawidłowa mapa dat wydarzeń: {path}")

    mapping: dict[str, str] = {}
    for key, value in raw_events.items():
        if not isinstance(key, str) or len(key) != 64:
            raise ValueError(f"Nieprawidłowy klucz wydarzenia w {path}: {key!r}")
        if not isinstance(value, str):
            raise ValueError(f"Nieprawidłowa data wydarzenia w {path}: {value!r}")
        try:
            parsed = date.fromisoformat(value)
        except ValueError as error:
            raise ValueError(f"Nieprawidłowa data ISO w {path}: {value!r}") from error
        if parsed.isoformat() != value:
            raise ValueError(f"Data nie jest w kanonicznym formacie ISO: {value!r}")
        mapping[key] = value
    return mapping
