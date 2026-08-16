#!/usr/bin/env python3
"""Build a compact Wyniki Zuzlowe v4 database from the source XLSM."""

from __future__ import annotations

import argparse
import gzip
import hashlib
import json
import math
import os
import tempfile
import unicodedata
from datetime import date, datetime, time, timedelta, timezone
from pathlib import Path
from typing import Any, Iterable

from openpyxl import load_workbook


FORMAT_VERSION = 4
FIRST_SEASON = 2010
LAST_SEASON = 2026
PLAYER_SHEET = "Zawodnicy"
DEFAULT_SOURCE_URL = (
    "https://drive.google.com/uc?"
    "id=15p7L2RPKcSMIVzoZXkGDkZZqR5xlBA9K&export=download"
)


class StringTable:
    """Insertion-ordered string table used by records and player metadata."""

    def __init__(self) -> None:
        # Index 0 is the empty string in the format consumed by the app.
        self.values: list[str] = [""]
        self._indices: dict[str, int] = {"": 0}

    def intern(self, value: str) -> int:
        index = self._indices.get(value)
        if index is None:
            index = len(self.values)
            self._indices[value] = index
            self.values.append(value)
        return index


def clean_text(value: Any) -> str:
    """Remove accidental cell padding while retaining the displayed spelling."""
    return " ".join(str(value).split())


def normalize_name(value: Any) -> str:
    """Create the search/matching form used by Wyniki Zuzlowe v4.

    NFD decomposition removes combining accents but intentionally leaves letters
    such as Polish ``ł`` distinct from ``l``.
    """
    folded = clean_text(value).casefold()
    return "".join(
        char
        for char in unicodedata.normalize("NFD", folded)
        if not unicodedata.combining(char)
    )


def birth_date_value(value: Any) -> Any:
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if value is None:
        return None
    if isinstance(value, str):
        return clean_text(value) or None
    return value


def record_value(value: Any, strings: StringTable) -> int | None:
    """Encode a result cell as a string-table reference used by the app."""
    if value is None:
        return None
    if isinstance(value, str):
        text = clean_text(value)
        return None if not text else strings.intern(text)
    if isinstance(value, datetime):
        return strings.intern(value.isoformat())
    if isinstance(value, (date, time)):
        return strings.intern(value.isoformat())
    if isinstance(value, timedelta):
        value = value.total_seconds()
    if isinstance(value, float) and not math.isfinite(value):
        return None
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return strings.intern(str(value))


def event_component(value: Any) -> str:
    """Canonical event-key component resilient to Excel text/number drift."""
    if value is None:
        return ""
    if isinstance(value, str):
        return clean_text(value)
    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, (date, time)):
        return value.isoformat()
    if isinstance(value, timedelta):
        return str(value.total_seconds())
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def atomic_write(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{path.name}.", suffix=".tmp", dir=path.parent
    )
    try:
        with os.fdopen(descriptor, "wb") as output:
            output.write(data)
            output.flush()
            os.fsync(output.fileno())
        os.replace(temporary_name, path)
    except BaseException:
        try:
            os.unlink(temporary_name)
        except FileNotFoundError:
            pass
        raise


def iter_player_rows(sheet: Any) -> Iterable[tuple[Any, Any, Any]]:
    for name, nationality, born in sheet.iter_rows(
        min_row=2, min_col=1, max_col=3, values_only=True
    ):
        if name is not None and clean_text(name):
            yield name, nationality, born


def build_database(source_path: Path, source_url: str, built: str) -> dict[str, Any]:
    workbook = load_workbook(
        source_path,
        read_only=True,
        data_only=True,
        keep_links=False,
    )
    try:
        expected_sheets = [str(year) for year in range(FIRST_SEASON, LAST_SEASON + 1)]
        missing = [
            name for name in [PLAYER_SHEET, *expected_sheets] if name not in workbook.sheetnames
        ]
        if missing:
            raise ValueError(f"Brak wymaganych arkuszy: {', '.join(missing)}")

        strings = StringTable()
        players: list[list[Any]] = []
        player_by_normalized_name: dict[str, int] = {}

        for raw_name, raw_nationality, raw_born in iter_player_rows(
            workbook[PLAYER_SHEET]
        ):
            name = clean_text(raw_name)
            normalized = normalize_name(name)
            nationality = clean_text(raw_nationality) if raw_nationality is not None else ""
            player_id = len(players)
            players.append(
                [name, strings.intern(nationality), birth_date_value(raw_born), normalized]
            )
            player_by_normalized_name.setdefault(normalized, player_id)

        years: dict[str, list[list[Any]]] = {}
        events: dict[str, list[list[int]]] = {}
        total_rows = 0
        total_events = 0

        for year in range(FIRST_SEASON, LAST_SEASON + 1):
            year_key = str(year)
            records: list[list[Any]] = []
            year_events: list[list[int]] = []
            current_event_key: tuple[str, ...] | None = None

            # B:P is 15 cells. The record uses B as the player lookup, C:M,
            # skips N, and then appends O:P.
            for cells in workbook[year_key].iter_rows(
                min_row=4, min_col=2, max_col=16, values_only=True
            ):
                raw_player_name = cells[0]
                if raw_player_name is None or not clean_text(raw_player_name):
                    continue

                player_name = clean_text(raw_player_name)
                normalized = normalize_name(player_name)
                player_id = player_by_normalized_name.get(normalized)
                if player_id is None:
                    player_id = len(players)
                    player_by_normalized_name[normalized] = player_id
                    players.append(
                        [player_name, strings.intern(""), None, normalized]
                    )

                raw_record_values = [*cells[1:12], cells[13], cells[14]]
                record = [player_id]
                record.extend(record_value(value, strings) for value in raw_record_values)
                records.append(record)

                # Record indices 5..12 (inclusive) correspond to source G:M,O.
                raw_event_values = (*cells[5:12], cells[13])
                event_key = tuple(event_component(value) for value in raw_event_values)
                record_index = len(records) - 1
                if event_key == current_event_key:
                    year_events[-1][1] += 1
                else:
                    year_events.append([record_index, 1])
                    current_event_key = event_key

            years[year_key] = records
            events[year_key] = year_events
            total_rows += len(records)
            total_events += len(year_events)

        stats = {
            "rows": total_rows,
            "players": len(players),
            "seasons": len(years),
            "from": FIRST_SEASON,
            "to": LAST_SEASON,
            "events": total_events,
        }
        return {
            "version": FORMAT_VERSION,
            "source": source_url,
            "built": built,
            "strings": strings.values,
            "players": players,
            "years": years,
            "stats": stats,
            "events": events,
        }
    finally:
        workbook.close()


def validate_expectations(stats: dict[str, int], args: argparse.Namespace) -> None:
    expectations = {
        "rows": args.expect_rows,
        "players": args.expect_players,
        "seasons": args.expect_seasons,
        "from": args.expect_from,
        "to": args.expect_to,
        "events": args.expect_events,
    }
    mismatches = [
        f"{key}: otrzymano {stats[key]}, oczekiwano {expected}"
        for key, expected in expectations.items()
        if expected is not None and stats[key] != expected
    ]
    if mismatches:
        raise ValueError("Niezgodne statystyki: " + "; ".join(mismatches))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source_xlsm", type=Path, help="Ścieżka do źródłowego PL2.xlsm")
    parser.add_argument("--output", type=Path, default=Path("db/latest.wzdb"))
    parser.add_argument("--version-file", type=Path, default=Path("db/version.json"))
    parser.add_argument("--source-url", default=DEFAULT_SOURCE_URL)
    parser.add_argument("--expect-rows", type=int)
    parser.add_argument("--expect-players", type=int)
    parser.add_argument("--expect-seasons", type=int)
    parser.add_argument("--expect-from", type=int)
    parser.add_argument("--expect-to", type=int)
    parser.add_argument("--expect-events", type=int)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    source_path = args.source_xlsm.resolve()
    if not source_path.is_file():
        raise FileNotFoundError(f"Nie znaleziono pliku źródłowego: {source_path}")

    built = datetime.now(timezone.utc).isoformat(timespec="seconds").replace(
        "+00:00", "Z"
    )
    source_sha256 = sha256_file(source_path)
    generator_sha256 = sha256_file(Path(__file__).resolve())
    database = build_database(source_path, args.source_url, built)
    validate_expectations(database["stats"], args)

    json_bytes = json.dumps(
        database, ensure_ascii=False, separators=(",", ":")
    ).encode("utf-8")
    wzdb_bytes = gzip.compress(json_bytes, compresslevel=9, mtime=0)
    atomic_write(args.output, wzdb_bytes)
    wzdb_sha256 = hashlib.sha256(wzdb_bytes).hexdigest()

    version = {
        "version": FORMAT_VERSION,
        "source": args.source_url,
        "source_sha256": source_sha256,
        "source_hash": source_sha256[:12],
        "generator_sha256": generator_sha256,
        "wzdb_sha256": wzdb_sha256,
        "built": built,
        "stats": database["stats"],
    }
    atomic_write(
        args.version_file,
        (json.dumps(version, ensure_ascii=False, indent=2) + "\n").encode("utf-8"),
    )

    print(json.dumps(version, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
