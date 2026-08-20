#!/usr/bin/env python3
"""Conservatively match 2026 WZDB events to the supplied official calendar."""

from __future__ import annotations

import argparse
import csv
import gzip
import hashlib
import json
import re
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

try:
    from scripts.event_dates import event_mapping_key
except ModuleNotFoundError:  # Direct execution: python scripts/match_event_dates.py
    from event_dates import event_mapping_key


DEFAULT_CUTOFF = "2026-08-17"
DEFAULT_CALENDAR = Path("data/event_dates_2026_candidates.json")
DEFAULT_MAPPING = Path("data/event_dates_2026.json")
DEFAULT_REPORT_DIR = Path("reports")
DEFAULT_SUPPLEMENT = Path("data/event_dates_2026_supplement_combined_v4.json")
EVENT_FIELDS = (
    "home",
    "away",
    "score",
    "league",
    "track",
    "competition",
    "round",
    "capacity",
)

SPECIAL_TRANSLATION = str.maketrans(
    {
        "ł": "l",
        "Ł": "L",
        "ø": "o",
        "Ø": "O",
        "ß": "ss",
        "æ": "ae",
        "Æ": "AE",
        "ð": "d",
        "Ð": "D",
        "þ": "th",
        "Þ": "Th",
    }
)

VENUE_ALIASES = {
    "gustrow": "guestrow",
    "guestrow": "guestrow",
    "malilla": "malilla",
    "zarnovica": "zarnovica",
    "zarnowica": "zarnovica",
    "praga": "prague",
    "praha": "prague",
    "prague": "prague",
    "debreczyn": "debrecen",
    "debrecen": "debrecen",
    "pilzno": "plzen",
    "plzen": "plzen",
    "ryga": "riga",
    "riga": "riga",
    "ostrow wlkp": "ostrow wielkopolski",
    "ostrow wlkp.": "ostrow wielkopolski",
    "grudziadza": "grudziadz",
    "kings lynn": "kings lynn",
    "lamothe landerron": "lamothe landerron",
    # Stadium/location aliases confirmed by the reviewed 2026 supplement.
    "gorican": "donji kraljevec",
    "donji kraljevec": "donji kraljevec",
    "middlesbrough": "redcar ecco arena",
    "redcar": "redcar ecco arena",
    "ecco arena": "redcar ecco arena",
}

TEAM_CODE_ALIASES = {
    "BYD": ("bydgoszcz", "polonia bydgoszcz"),
    "KRO": ("krosno", "wilki krosno"),
    "OST": ("ostrow", "ostrovia ostrow"),
    "LOD": ("lodz", "orzel lodz"),
    "RZE": ("rzeszow", "stal rzeszow"),
    "RYB": ("rybnik", "row rybnik"),
    "POZ": ("poznan", "psz poznan"),
    "PIL": ("pila", "polonia pila"),
    "KRA": ("krakow", "speedway krakow"),
    "GDA": ("gdansk", "wybrzeze gdansk"),
    "TOR": ("torun",),
    "GRU": ("grudziadz", "gkm grudziadz"),
    "LUB": ("lublin", "motor lublin"),
    "CZE": ("czestochowa", "wlokniarz czestochowa"),
    "ZIE": ("zielona gora", "falubaz zielona gora"),
    "GOR": ("gorzow", "stal gorzow"),
    "LES": ("leszno", "unia leszno"),
    "WRO": ("wroclaw", "sparta wroclaw"),
    "SWI": ("swietochlowice", "slask swietochlowice"),
    "DAU": ("daugavpils", "lokomotiv daugavpils"),
    "LAN": ("landshut", "landshut devils"),
    "OPO": ("opole", "kolejarz opole"),
    "TARC": ("tarnow", "unia tarnow"),
}

GENERIC_TITLE_TOKENS = {
    "zawody",
    "turniej",
    "indywidualny",
    "towarzyski",
    "mecz",
    "liga",
    "puchar",
    "final",
    "runda",
    "eliminacje",
    "kwalifikacje",
    "czesc",
    "zasadnicza",
}


def normalize(value: Any) -> str:
    text = str(value or "").translate(SPECIAL_TRANSLATION).casefold()
    text = "".join(
        char
        for char in unicodedata.normalize("NFKD", text)
        if not unicodedata.combining(char)
    )
    return " ".join(re.sub(r"[^a-z0-9]+", " ", text).split())


def venue_key(value: Any) -> str:
    key = normalize(value)
    return VENUE_ALIASES.get(key, key)


def team_code(value: Any) -> str | None:
    key = normalize(value)
    if not key:
        return None
    upper = key.upper()
    if upper in TEAM_CODE_ALIASES:
        return upper
    for code, aliases in TEAM_CODE_ALIASES.items():
        if any(key == alias or key.endswith(f" {alias}") or alias.endswith(f" {key}") for alias in aliases):
            return code
    return None


def teams_equivalent(left: Any, right: Any) -> bool:
    left_key, right_key = normalize(left), normalize(right)
    if not left_key or not right_key:
        return False
    if left_key == right_key:
        return True
    generic_team_tokens = {"speedway", "club", "klub", "league", "premiership", "championship", "bauhaus"}
    left_compact = " ".join(token for token in left_key.split() if token not in generic_team_tokens)
    right_compact = " ".join(token for token in right_key.split() if token not in generic_team_tokens)
    if left_compact and left_compact == right_compact:
        return True
    left_code, right_code = team_code(left), team_code(right)
    return bool(left_code and left_code == right_code)


def competition_family(*values: Any) -> str:
    text = normalize(" ".join(str(value or "") for value in values))
    patterns = (
        ("sgp2", r"\bsgp ?2\b"),
        ("sgp3", r"\bsgp ?3\b"),
        ("sgp4", r"\bsgp ?4\b"),
        ("son2", r"\bson ?2\b|speedway of nations ?2"),
        ("son", r"\bson\b|speedway of nations"),
        ("sec", r"\bsec\b|speedway euro champ|\bime\b.*(?:cwiercfinal|elimin|kwalifik)"),
        ("gp", r"\bgp\b|\bsgp\b|grand prix|\bims\b"),
        ("world_cup", r"world cup|\bdps\b"),
        ("dme_u24", r"\bdme\b.*u24|u24.*\bdme\b"),
        ("dme", r"\bdme\b"),
        ("dmpj", r"\bdmpj\b"),
        ("u24", r"u24|ekstraliga u24"),
        ("mppk", r"\bmppk\b"),
        ("mppk", r"\bmppk\b|mistrzostw.*par"),
        ("i_liga", r"\bi liga\b|\bm2e\b|2 ekstraliga"),
        ("ii_liga", r"\bii liga\b|\bklz\b"),
        ("ekstraliga", r"\bekstraliga\b|\bpgee\b"),
        ("premiership", r"\bpremiership\b"),
        ("championship", r"\bchampionship\b"),
        ("national_league", r"national lea(?:gue|uge)"),
        ("elitserien", r"\belitserien\b|sveriges speedway liga"),
        ("allsvenskan", r"\ballsvenskan\b"),
        ("denmark", r"dm danii|division [123]|danska liga"),
        ("friendly", r"\bsparing\b|mecz towarzyski"),
    )
    for family, pattern in patterns:
        if re.search(pattern, text):
            return family
    return ""


def round_marker(*values: Any) -> str:
    text = normalize(" ".join(str(value or "") for value in values))
    ordinal = re.search(r"(?:runda|round|final)\s*(\d+)", text)
    if not ordinal:
        ordinal = re.search(r"(\d+)\s*(?:runda|round|final)", text)
    if ordinal:
        return f"ordinal:{int(ordinal.group(1))}"
    for label, pattern in (
        ("semi", r"(?:polfinal|semi final)\s*(\d+)"),
        ("qualifying", r"(?:eliminacje|kwalifikacje|qualifying round)\s*(\d+)"),
    ):
        match = re.search(pattern, text)
        if match:
            return f"{label}:{int(match.group(1))}"
    if re.search(r"\b(?:eliminacje|kwalifikacje|qualifying)\b", text):
        return "qualifying"
    if re.search(r"\b(?:polfinal|semi final)\b", text):
        return "semi"
    if re.search(r"\bfinal\b", text):
        return "final"
    return ""


def distinctive_title_match(left: Any, right: Any) -> tuple[bool, str]:
    left_key, right_key = normalize(left), normalize(right)
    if not left_key or not right_key:
        return False, ""
    if left_key == right_key and len(left_key) >= 5:
        return True, "exact"
    left_tokens, right_tokens = set(left_key.split()), set(right_key.split())
    distinctive_left = left_tokens - GENERIC_TITLE_TOKENS
    distinctive_right = right_tokens - GENERIC_TITLE_TOKENS
    if len(distinctive_left) >= 2 and distinctive_left <= right_tokens:
        return True, "contained"
    if len(distinctive_right) >= 2 and distinctive_right <= left_tokens:
        return True, "contained"
    return False, ""


def split_candidate_teams(candidate: dict[str, Any]) -> list[str]:
    parts = [part.strip() for part in re.split(r"\s+[-–—]\s+", str(candidate.get("event") or ""))]
    parts = [part for part in parts if part]
    while parts and re.search(r"\b(?:runda|round|final|polfinal|semi final)\b", normalize(parts[-1])):
        parts.pop()
    if len(parts) < 2:
        return []
    if len(parts) > 2 and not all(normalize(part).upper() in TEAM_CODE_ALIASES for part in parts):
        return []
    return [re.sub(r"\s+(?:Premiership|Championship|League)\s*$", "", part, flags=re.I) for part in parts]


def competition_phase(*values: Any) -> str:
    text = normalize(" ".join(str(value or "") for value in values))
    if any(token in text for token in ("knock out cup", "bsn series", " puchar")) or text.startswith("puchar"):
        return "cup"
    if "czesc zasadnicza" in text:
        return "league"
    return ""


@dataclass
class PhysicalEvent:
    source_event_index: int
    start: int
    count: int
    values: dict[str, str]
    mapping_key: str


@dataclass
class LogicalEvent:
    logical_event_index: int
    physical: list[PhysicalEvent]
    values: dict[str, str]
    teams: list[dict[str, str]] = field(default_factory=list)

    @property
    def start(self) -> int:
        return self.physical[0].start

    @property
    def count(self) -> int:
        return sum(item.count for item in self.physical)

    @property
    def source_event_index(self) -> int:
        return self.physical[0].source_event_index


def decode_events(database: dict[str, Any], season: str) -> list[LogicalEvent]:
    strings, rows = database["strings"], database["years"][season]

    def value(index: Any) -> str:
        return "" if index is None else str(strings[index] or "")

    occurrence: Counter[tuple[str, ...]] = Counter()
    physical: list[PhysicalEvent] = []
    for source_event_index, event_ref in enumerate(database["events"][season]):
        start, count = event_ref[:2]
        row = rows[start]
        values = dict(zip(EVENT_FIELDS, (value(row[index]) for index in range(5, 13))))
        signature = tuple(values[field] for field in EVENT_FIELDS)
        ordinal = occurrence[signature]
        occurrence[signature] += 1
        physical.append(
            PhysicalEvent(
                source_event_index=source_event_index,
                start=start,
                count=count,
                values=values,
                mapping_key=event_mapping_key(season, signature, ordinal),
            )
        )

    logical: list[LogicalEvent] = []
    index = 0
    while index < len(physical):
        first = physical[index]
        signature = tuple(
            normalize(first.values[field])
            for field in ("league", "track", "competition", "round")
        )
        strong = bool(season and first.values["track"] and first.values["competition"])
        end = index + 1
        while (
            strong
            and end < len(physical)
            and tuple(
                normalize(physical[end].values[field])
                for field in ("league", "track", "competition", "round")
            )
            == signature
            and physical[end - 1].start + physical[end - 1].count == physical[end].start
        ):
            end += 1
        run = physical[index:end]
        teams: list[dict[str, str]] = []
        seen: set[str] = set()
        for item in run:
            home, away, score = (
                item.values["home"],
                item.values["away"],
                item.values["score"],
            )
            if score and bool(home) != bool(away):
                name = home or away
                key = normalize(name)
                if key and key not in seen:
                    seen.add(key)
                    teams.append({"name": name, "score": score})
        merge = strong and len(run) > 1 and len(teams) > 1
        groups = [run] if merge else [[item] for item in run]
        for group in groups:
            logical.append(
                LogicalEvent(
                    logical_event_index=len(logical),
                    physical=group,
                    values=group[0].values,
                    teams=teams if merge else [],
                )
            )
        index = end
    return logical


def event_family(event: LogicalEvent) -> str:
    return competition_family(event.values["competition"], event.values["league"])


def candidate_family(candidate: dict[str, Any]) -> str:
    return competition_family(candidate.get("competition"), candidate.get("event"))


def event_team_names(event: LogicalEvent) -> list[str]:
    if event.teams:
        return [team["name"] for team in event.teams]
    home, away = event.values["home"], event.values["away"]
    return [home, away] if home and away else []


def match_evidence(event: LogicalEvent, candidate: dict[str, Any]) -> dict[str, Any] | None:
    if venue_key(event.values["track"]) != venue_key(candidate.get("venue")):
        return None
    evidence: list[str] = ["same_venue"]
    score = 20
    candidate_teams = split_candidate_teams(candidate)
    event_teams = event_team_names(event)
    match_kind = ""

    if len(event_teams) == 2 and len(candidate_teams) == 2:
        if teams_equivalent(event_teams[0], candidate_teams[0]) and teams_equivalent(event_teams[1], candidate_teams[1]):
            score += 115
            match_kind = "team_ordered"
            evidence.append("ordered_teams")
        elif teams_equivalent(event_teams[0], candidate_teams[1]) and teams_equivalent(event_teams[1], candidate_teams[0]):
            score += 75
            match_kind = "team_unordered"
            evidence.append("unordered_teams")
    elif len(event_teams) >= 3 and len(candidate_teams) >= 3:
        left_codes = {team_code(value) for value in event_teams}
        right_codes = {team_code(value) for value in candidate_teams}
        if None not in left_codes and None not in right_codes and left_codes == right_codes:
            score += 110
            match_kind = "multi_team_exact"
            evidence.append("same_team_set")

    left_family, right_family = event_family(event), candidate_family(candidate)
    if left_family and left_family == right_family:
        score += 55
        evidence.append(f"family:{left_family}")
        if not match_kind:
            match_kind = "family_venue"
    elif left_family and right_family and left_family != right_family:
        score -= 90
        evidence.append(f"family_conflict:{left_family}!={right_family}")

    event_phase = competition_phase(event.values["competition"])
    candidate_phase = competition_phase(candidate.get("competition"))
    if not event_phase and event_teams and event.values["league"]:
        event_phase = "league"
    if event_phase and candidate_phase:
        if event_phase == candidate_phase:
            score += 40
            evidence.append(f"phase:{event_phase}")
        else:
            score -= 80
            evidence.append("phase_conflict")

    event_round = round_marker(event.values["round"], event.values["competition"])
    candidate_round = round_marker(candidate.get("event"), candidate.get("competition"))
    if event_round and candidate_round:
        if event_round == candidate_round:
            score += 35
            evidence.append(f"round:{event_round}")
            if match_kind == "family_venue":
                match_kind = "family_round"
        elif event_round.split(":", 1)[0] == candidate_round.split(":", 1)[0] and ":" not in event_round:
            score += 8
            evidence.append("same_stage")
        else:
            score -= 50
            evidence.append("round_conflict")

    title_match, title_kind = distinctive_title_match(event.values["competition"], candidate.get("event"))
    if title_match:
        score += 85 if title_kind == "exact" else 70
        evidence.append(f"title_{title_kind}")
        if not match_kind or match_kind == "family_venue":
            match_kind = f"title_{title_kind}"

    if not match_kind:
        return None
    return {
        "score": score,
        "kind": match_kind,
        "evidence": evidence,
    }


def category_for(event: LogicalEvent) -> str:
    family = event_family(event)
    league = normalize(event.values["league"])
    if family in {"gp", "sgp2", "sgp3", "sgp4", "son", "son2", "world_cup"}:
        return "FIM"
    if family in {"sec", "dme", "dme_u24"} or normalize(event.values["competition"]) in {
        "ime",
        "imej",
        "mep",
        "dme",
        "dme u24",
    }:
        return "FIM Europe"
    if league in {"ekstraliga", "i liga", "ii liga", "polska"}:
        return "Polska"
    if league in {"premiership", "championship", "wielka brytania", "national league", "national leauge"}:
        return "Wielka Brytania"
    if league in {"elitserien", "allsvenskan", "szwecja"}:
        return "Szwecja"
    if league == "dania":
        return "Dania"
    if league == "niemcy":
        return "Niemcy"
    return "pozostałe"


def event_report_base(event: LogicalEvent) -> dict[str, Any]:
    values = event.values
    return {
        "event_index": event.logical_event_index,
        "source_event_index": event.source_event_index,
        "physical_event_indexes": [item.source_event_index for item in event.physical],
        "first_record_index": event.start,
        "last_record_index": event.start + event.count - 1,
        "track": values["track"],
        "league": values["league"],
        "competition": values["competition"],
        "round": values["round"],
        "home": values["home"],
        "away": values["away"],
        "score": values["score"],
        "record_count": event.count,
        "category": category_for(event),
        "teams": event.teams,
        "mapping_keys": [item.mapping_key for item in event.physical],
    }


def candidate_report(candidate_index: int, candidate: dict[str, Any], evidence: dict[str, Any]) -> dict[str, Any]:
    return {
        "candidate_index": candidate_index,
        "date": candidate.get("date", ""),
        "venue": candidate.get("venue", ""),
        "event": candidate.get("event", ""),
        "competition": candidate.get("competition", ""),
        "source": candidate.get("source", ""),
        "source_detail": candidate.get("source_detail", ""),
        "source_page": candidate.get("source_page", ""),
        "calendar_confidence": candidate.get("confidence", ""),
        "match_key": candidate.get("match_key", ""),
        "score": evidence["score"],
        "match_kind": evidence["kind"],
        "evidence": evidence["evidence"],
    }


def match_events(
    events: list[LogicalEvent],
    candidates: list[dict[str, Any]],
    cutoff: str = DEFAULT_CUTOFF,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    candidates_by_venue: defaultdict[str, list[tuple[int, dict[str, Any]]]] = defaultdict(list)
    for candidate_index, candidate in enumerate(candidates):
        candidates_by_venue[venue_key(candidate.get("venue"))].append((candidate_index, candidate))

    event_group_counts = Counter((venue_key(event.values["track"]), event_family(event)) for event in events)
    candidate_group_counts = Counter((venue_key(candidate.get("venue")), candidate_family(candidate)) for candidate in candidates)
    provisional: list[dict[str, Any]] = []

    for event in events:
        reasonable: list[dict[str, Any]] = []
        for candidate_index, candidate in candidates_by_venue.get(venue_key(event.values["track"]), []):
            evidence = match_evidence(event, candidate)
            if evidence and evidence["score"] >= 55:
                reasonable.append(candidate_report(candidate_index, candidate, evidence))
        reasonable.sort(key=lambda item: (-item["score"], item["date"], item["candidate_index"]))
        base = event_report_base(event)
        decision = "UNMATCHED"
        reason = "no credible calendar candidate"
        selected: dict[str, Any] | None = None
        if reasonable:
            top = reasonable[0]
            same_strength = [item for item in reasonable if item["score"] >= top["score"] - 5]
            identity_conflict = any(
                value == "round_conflict"
                or value == "phase_conflict"
                or value.startswith("family_conflict:")
                for value in top["evidence"]
            )
            strong_kind = top["match_kind"] in {
                "team_ordered",
                "multi_team_exact",
                "family_round",
                "title_exact",
                "title_contained",
            } and not identity_conflict
            unique_family_venue = (
                top["match_kind"] == "family_venue"
                and not identity_conflict
                and event_family(event)
                and event_group_counts[(venue_key(event.values["track"]), event_family(event))] == 1
                and candidate_group_counts[(venue_key(event.values["track"]), event_family(event))] == 1
            )
            if len(same_strength) > 1:
                decision = "MEDIUM"
                reason = "multiple similarly strong calendar candidates"
            elif top["date"] > cutoff:
                decision = "MEDIUM"
                reason = "only credible candidate is after the analysis cutoff"
            elif top["calendar_confidence"] != "HIGH":
                decision = "MEDIUM"
                reason = "candidate calendar confidence is not HIGH"
            elif strong_kind or unique_family_venue:
                decision = "HIGH"
                reason = ", ".join(top["evidence"])
                selected = top
            else:
                decision = "MEDIUM"
                reason = "venue agrees but event identity is not unique enough"
        provisional.append(
            {
                **base,
                "confidence": decision,
                "reason": reason,
                "selected": selected,
                "candidates": reasonable,
            }
        )

    candidate_to_high: defaultdict[int, list[dict[str, Any]]] = defaultdict(list)
    for item in provisional:
        if item["confidence"] == "HIGH" and item["selected"]:
            candidate_to_high[item["selected"]["candidate_index"]].append(item)
    for conflicts in candidate_to_high.values():
        if len(conflicts) <= 1:
            continue
        for item in conflicts:
            item["confidence"] = "MEDIUM"
            item["reason"] = "one calendar row matches multiple WZDB events"
            item["selected"] = None

    matched: list[dict[str, Any]] = []
    ambiguous: list[dict[str, Any]] = []
    unmatched: list[dict[str, Any]] = []
    for item in provisional:
        if item["confidence"] == "HIGH":
            selected = item.pop("selected")
            item["event_date"] = selected["date"]
            item["date_source"] = selected["source"]
            item["calendar_candidate"] = selected
            item.pop("candidates", None)
            matched.append(item)
        elif item["confidence"] == "MEDIUM":
            item.pop("selected", None)
            item["event_date"] = ""
            item["date_source"] = ""
            ambiguous.append(item)
        else:
            item.pop("selected", None)
            item.pop("candidates", None)
            item["event_date"] = ""
            item["date_source"] = ""
            unmatched.append(item)
    return matched, ambiguous, unmatched


def apply_reviewed_supplement(
    events: list[LogicalEvent],
    matched: list[dict[str, Any]],
    ambiguous: list[dict[str, Any]],
    unmatched: list[dict[str, Any]],
    supplement: dict[str, Any],
    cutoff: str = DEFAULT_CUTOFF,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    """Apply manually reviewed HIGH decisions after validating immutable identity."""
    event_by_id = {event.logical_event_index: event for event in events}
    buckets = {
        "matched": {item["event_index"]: item for item in matched},
        "ambiguous": {item["event_index"]: item for item in ambiguous},
        "unmatched": {item["event_index"]: item for item in unmatched},
    }
    seen: set[int] = set()
    for decision in supplement.get("events", []):
        event_id = int(decision["event_id"])
        if event_id in seen:
            raise ValueError(f"Powtórzony event_id w uzupełnieniu: {event_id}")
        seen.add(event_id)
        event = event_by_id.get(event_id)
        if event is None:
            raise ValueError(f"Uzupełnienie wskazuje nieistniejący event_id: {event_id}")
        if decision.get("status") != "HIGH":
            raise ValueError(f"Do mapy można zastosować tylko HIGH: event_id {event_id}")
        event_date = str(decision.get("date") or "")
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", event_date) or event_date > cutoff:
            raise ValueError(f"Nieprawidłowa data uzupełnienia dla event_id {event_id}")

        expected_identity = decision.get("identity", {})
        actual_identity = {field: event.values[field] for field in EVENT_FIELDS}
        if expected_identity != actual_identity:
            raise ValueError(f"Zmieniła się tożsamość WZDB dla event_id {event_id}")
        expected_keys = decision.get("mapping_keys", [])
        actual_keys = [fragment.mapping_key for fragment in event.physical]
        if expected_keys != actual_keys:
            raise ValueError(f"Zmieniły się klucze fragmentów dla event_id {event_id}")

        previous = buckets["matched"].get(event_id)
        if previous and previous.get("event_date") != event_date:
            raise ValueError(
                f"Konflikt automatycznej i ręcznej daty dla event_id {event_id}: "
                f"{previous.get('event_date')} != {event_date}"
            )
        for bucket in buckets.values():
            bucket.pop(event_id, None)
        base = event_report_base(event)
        aliases = decision.get("aliases", {})
        reviewed_candidate = {
            "candidate_index": f"supplement:{event_id}",
            "date": event_date,
            "venue": next(iter(aliases.values()), event.values["track"]),
            "event": f"reviewed supplement #{event_id}",
            "competition": event.values["competition"],
            "source": decision.get("basis", "reviewed supplement"),
            "source_detail": decision.get("review", "pełna tożsamość zgodna z WZDB"),
            "source_page": "",
            "source_url": decision.get("source_url", ""),
            "calendar_confidence": "HIGH",
            "match_key": f"supplement:{event_id}:{event_date}",
            "score": 1000,
            "match_kind": "reviewed_supplement",
            "evidence": ["event_id", "exact_identity", "mapping_keys", "reviewed_roster"],
        }
        buckets["matched"][event_id] = {
            **base,
            "confidence": "HIGH",
            "reason": f"reviewed supplement: {decision.get('review', 'pełna zgodność')}",
            "event_date": event_date,
            "date_source": decision.get("basis", "reviewed supplement"),
            "calendar_candidate": reviewed_candidate,
        }

    return tuple(
        sorted(bucket.values(), key=lambda item: item["event_index"])
        for bucket in (
            buckets["matched"],
            buckets["ambiguous"],
            buckets["unmatched"],
        )
    )


def coverage_summary(
    events: list[LogicalEvent],
    matched: list[dict[str, Any]],
    ambiguous: list[dict[str, Any]],
    unmatched: list[dict[str, Any]],
) -> dict[str, Any]:
    total_records = sum(event.count for event in events)
    matched_records = sum(item["record_count"] for item in matched)
    categories: dict[str, dict[str, Any]] = {}
    for category in sorted({category_for(event) for event in events}):
        category_events = [event for event in events if category_for(event) == category]
        high = [item for item in matched if item["category"] == category]
        medium = [item for item in ambiguous if item["category"] == category]
        no_match = [item for item in unmatched if item["category"] == category]
        records = sum(event.count for event in category_events)
        high_records = sum(item["record_count"] for item in high)
        categories[category] = {
            "events": len(category_events),
            "high": len(high),
            "ambiguous": len(medium),
            "unmatched": len(no_match),
            "event_coverage_pct": round(100 * len(high) / len(category_events), 2) if category_events else 0,
            "records": records,
            "matched_records": high_records,
            "record_coverage_pct": round(100 * high_records / records, 2) if records else 0,
        }
    return {
        "events": len(events),
        "high": len(matched),
        "ambiguous": len(ambiguous),
        "unmatched": len(unmatched),
        "event_coverage_pct": round(100 * len(matched) / len(events), 2) if events else 0,
        "records": total_records,
        "matched_records": matched_records,
        "record_coverage_pct": round(100 * matched_records / total_records, 2) if total_records else 0,
        "categories": categories,
    }


def write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def write_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    columns = [
        "event_index",
        "source_event_index",
        "physical_event_indexes",
        "first_record_index",
        "last_record_index",
        "track",
        "league",
        "competition",
        "round",
        "home",
        "away",
        "score",
        "record_count",
        "category",
        "event_date",
        "date_source",
        "confidence",
        "reason",
        "teams",
        "calendar_candidate",
        "candidates",
        "mapping_keys",
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as output:
        writer = csv.DictWriter(output, fieldnames=columns, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            serialized = dict(row)
            for field_name in (
                "physical_event_indexes",
                "teams",
                "calendar_candidate",
                "candidates",
                "mapping_keys",
            ):
                if field_name in serialized:
                    serialized[field_name] = json.dumps(serialized[field_name], ensure_ascii=False, separators=(",", ":"))
            writer.writerow(serialized)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("wzdb", type=Path, help="Current db/latest.wzdb")
    parser.add_argument("--calendar", type=Path, default=DEFAULT_CALENDAR)
    parser.add_argument("--season", default="2026")
    parser.add_argument("--cutoff", default=DEFAULT_CUTOFF)
    parser.add_argument("--mapping", type=Path, default=DEFAULT_MAPPING)
    parser.add_argument("--report-dir", type=Path, default=DEFAULT_REPORT_DIR)
    parser.add_argument("--supplement", type=Path, default=DEFAULT_SUPPLEMENT)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    raw_wzdb = args.wzdb.read_bytes()
    database = json.loads(gzip.decompress(raw_wzdb).decode("utf-8"))
    calendar = json.loads(args.calendar.read_text(encoding="utf-8"))
    candidates = calendar.get("events", calendar)
    events = decode_events(database, str(args.season))
    matched, ambiguous, unmatched = match_events(events, candidates, args.cutoff)
    supplement = (
        json.loads(args.supplement.read_text(encoding="utf-8"))
        if args.supplement.is_file()
        else {"events": []}
    )
    matched, ambiguous, unmatched = apply_reviewed_supplement(
        events, matched, ambiguous, unmatched, supplement, args.cutoff
    )
    summary = coverage_summary(events, matched, ambiguous, unmatched)
    report_metadata = {
        "season": int(args.season),
        "cutoff": args.cutoff,
        "wzdb_sha256": hashlib.sha256(raw_wzdb).hexdigest(),
        "calendar_file": args.calendar.name,
        "calendar_events": len(candidates),
        "supplement_file": args.supplement.name if args.supplement.is_file() else "",
        "supplement_events": len(supplement.get("events", [])),
        "summary": summary,
    }

    for name, rows in (
        ("matched", matched),
        ("ambiguous", ambiguous),
        ("unmatched", unmatched),
    ):
        write_json(
            args.report_dir / f"event_dates_2026_{name}.json",
            {**report_metadata, "status": name, "events": rows},
        )
        write_csv(args.report_dir / f"event_dates_2026_{name}.csv", rows)

    write_json(
        args.report_dir / "event_dates_2026_remaining_undated.json",
        {**report_metadata, "status": "remaining_undated", "events": unmatched},
    )

    approved = {
        "version": 1,
        "season": int(args.season),
        "cutoff": args.cutoff,
        "source_wzdb_sha256": report_metadata["wzdb_sha256"],
        "events": {
            key: item["event_date"]
            for item in matched
            for key in item["mapping_keys"]
        },
    }
    write_json(args.mapping, approved)
    write_json(args.report_dir / "event_dates_2026_summary.json", report_metadata)
    print(json.dumps(report_metadata, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
