"""Text cleaning and FM value normalization for OCR output."""

from __future__ import annotations

import re
from datetime import datetime
from typing import Any

_EMPTY_TOKENS = frozenset({"", "-", "—", "–", "n/a", "na", "none", "null"})

_CURRENCY_RE = re.compile(
    r"^[\s£$€¥]?"
    r"([\d,]+(?:\.\d+)?)"
    r"\s*([kmb])?"
    r"[\s£$€¥]?"
    r"\s*$",
    re.IGNORECASE,
)

_WAGE_RE = re.compile(
    r"^[\s£$€¥]?"
    r"([\d,]+(?:\.\d+)?)"
    r"\s*([kmb])?"
    r"\s*(?:p/?\s*w|per\s*week|p/?\s*a|per\s*year|p/?\s*m|per\s*month)?"
    r"[\s£$€¥]?"
    r"\s*$",
    re.IGNORECASE,
)

_PERCENT_RE = re.compile(r"^([\d.,]+)\s*%?\s*$")

_STAR_RE = re.compile(r"[★☆]")
_HALF_STAR = "½"

_DATE_PATTERNS = (
    "%d/%m/%Y",
    "%m/%d/%Y",
    "%d-%m-%Y",
    "%Y-%m-%d",
    "%d %b %Y",
    "%d %B %Y",
    "%b %Y",
    "%B %Y",
)

_FM_POSITIONS = {
    "gk": "GK",
    "goalkeeper": "GK",
    "dc": "DC",
    "defender": "DC",
    "dl": "DL",
    "dr": "DR",
    "wbl": "WBL",
    "wbr": "WBR",
    "dm": "DM",
    "mc": "MC",
    "ml": "ML",
    "mr": "MR",
    "aml": "AML",
    "amr": "AMR",
    "amc": "AMC",
    "st": "ST",
    "sc": "ST",
}

_FOOTEDNESS = {
    "left": "Left",
    "right": "Right",
    "either": "Either",
    "both": "Either",
}


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    text = re.sub(r"\s+", " ", text)
    return text


def is_empty_value(value: Any) -> bool:
    if value is None:
        return True
    text = clean_text(value)
    return text.lower() in _EMPTY_TOKENS


def slugify_column(name: str) -> str:
    text = str(name).strip().lower()
    text = re.sub(r"[^\w\s]", "", text)
    text = re.sub(r"\s+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text or "column"


def parse_int(value: Any) -> int | None:
    if is_empty_value(value):
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value)

    text = clean_text(value).replace(",", "")
    try:
        return int(float(text))
    except ValueError:
        return None


def parse_float(value: Any) -> float | None:
    if is_empty_value(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = clean_text(value).replace(",", "")
    try:
        return float(text)
    except ValueError:
        return None


def parse_percentage(value: Any) -> float | None:
    if is_empty_value(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = clean_text(value)
    match = _PERCENT_RE.match(text)
    if not match:
        return parse_float(text)

    raw = match.group(1).replace(",", ".")
    try:
        return float(raw)
    except ValueError:
        return None


def parse_currency(value: Any) -> float | None:
    if is_empty_value(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = clean_text(value).replace(",", "")
    match = _CURRENCY_RE.match(text)
    if not match:
        return None

    amount = float(match.group(1))
    suffix = (match.group(2) or "").lower()
    multiplier = {"k": 1_000, "m": 1_000_000, "b": 1_000_000_000}.get(suffix, 1)
    return amount * multiplier


def parse_wage(value: Any) -> float | None:
    if is_empty_value(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = clean_text(value).replace(",", "")
    match = _WAGE_RE.match(text)
    if not match:
        return parse_currency(value)

    amount = float(match.group(1))
    suffix = (match.group(2) or "").lower()
    multiplier = {"k": 1_000, "m": 1_000_000, "b": 1_000_000_000}.get(suffix, 1)
    return amount * multiplier


def parse_date(value: Any) -> str | None:
    if is_empty_value(value):
        return None

    text = clean_text(value)
    for pattern in _DATE_PATTERNS:
        try:
            parsed = datetime.strptime(text, pattern)
            return parsed.date().isoformat()
        except ValueError:
            continue
    return text


def normalize_position(value: Any) -> str | None:
    if is_empty_value(value):
        return None
    text = clean_text(value)
    parts = re.split(r"[,/|\s]+", text.lower())
    normalized = []
    for part in parts:
        part = part.strip()
        if not part:
            continue
        normalized.append(_FM_POSITIONS.get(part, part.upper()))
    return ", ".join(normalized) if normalized else text


def normalize_footedness(value: Any) -> str | None:
    if is_empty_value(value):
        return None
    text = clean_text(value).lower()
    for key, label in _FOOTEDNESS.items():
        if key in text:
            return label
    return clean_text(value)


def normalize_star_rating(value: Any) -> float | None:
    if is_empty_value(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = clean_text(value)
    if _STAR_RE.search(text):
        full = text.count("★")
        half = text.count(_HALF_STAR)
        empty = text.count("☆")
        if full + empty + half == 0:
            return None
        return full + (half * 0.5)

    try:
        return float(text)
    except ValueError:
        return None


def normalize_name(value: Any) -> str | None:
    text = clean_text(value)
    return text or None


def normalize_nationality(value: Any) -> str | None:
    return normalize_name(value)


def normalize_club(value: Any) -> str | None:
    return normalize_name(value)


_NORMALIZERS = {
    "age": parse_int,
    "number": parse_float,
    "percentage": parse_percentage,
    "currency": parse_currency,
    "wage": parse_wage,
    "date": parse_date,
    "star_rating": normalize_star_rating,
    "position": normalize_position,
    "nationality": normalize_nationality,
    "club": normalize_club,
    "name": normalize_name,
}


def apply_normalizer(normalizer: str, value: Any) -> Any:
    fn = _NORMALIZERS.get(normalizer)
    if fn is None:
        return clean_text(value) or None
    return fn(value)
