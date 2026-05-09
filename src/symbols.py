from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable


NORMALIZE_RE = re.compile(r"[^A-Z0-9]")


@dataclass(frozen=True)
class SymbolMatch:
    label: str
    symbol: str | None
    reason: str


def normalize(value: str) -> str:
    value = value.upper()
    value = value.replace("&", "AND")
    value = re.sub(r"\b(LTD|LIMITED|BE|SM|ST)\b", "", value)
    return NORMALIZE_RE.sub("", value)


def load_overrides(path: Path) -> dict[str, str]:
    if not path.exists():
        return {}
    payload = json.loads(path.read_text())
    return {str(key).strip(): str(value).strip().upper() for key, value in payload.items()}


def build_symbol_map(
    labels: Iterable[str],
    instruments: Iterable[dict],
    overrides: dict[str, str],
) -> dict[str, SymbolMatch]:
    by_symbol = {str(item.get("tradingsymbol", "")).upper(): item for item in instruments}
    normalized_symbol: dict[str, str] = {}
    normalized_name: dict[str, str] = {}
    for item in instruments:
        symbol = str(item.get("tradingsymbol", "")).upper()
        name = str(item.get("name", "")).upper()
        if symbol:
            normalized_symbol.setdefault(normalize(symbol), symbol)
        if name:
            normalized_name.setdefault(normalize(name), symbol)

    result: dict[str, SymbolMatch] = {}
    for label in labels:
        clean_label = label.strip()
        override = overrides.get(clean_label)
        if override:
            if override in by_symbol:
                result[clean_label] = SymbolMatch(clean_label, override, "override")
            else:
                    result[clean_label] = SymbolMatch(clean_label, None, f"override '{override}' not found in instruments")
            continue

        candidates = _label_candidates(clean_label)
        match = None
        reason = "unmatched"
        for candidate in candidates:
            normalized = normalize(candidate)
            if normalized in normalized_symbol:
                match = normalized_symbol[normalized]
                reason = "normalized tradingsymbol"
                break
            if normalized in normalized_name:
                match = normalized_name[normalized]
                reason = "normalized instrument name"
                break
        result[clean_label] = SymbolMatch(clean_label, match, reason)
    return result


def _label_candidates(label: str) -> list[str]:
    parts = re.split(r"[/\\]", label)
    candidates = [label, *parts]
    cleaned = []
    for item in candidates:
        item = item.strip()
        if item and item not in cleaned:
            cleaned.append(item)
    return cleaned
