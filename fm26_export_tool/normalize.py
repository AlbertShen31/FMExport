"""Column and cell-value normalization for FM26 HTML exports."""

from __future__ import annotations

import re
from typing import Any

import pandas as pd

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


def slugify_column(name: str) -> str:
    """Convert a column header to a stable snake_case identifier."""
    text = str(name).strip().lower()
    text = re.sub(r"[^\w\s]", "", text)
    text = re.sub(r"\s+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text or "column"


def deduplicate_columns(names: list[str]) -> list[str]:
    """Ensure column identifiers are unique by appending numeric suffixes."""
    seen: dict[str, int] = {}
    result: list[str] = []
    for name in names:
        base = name or "column"
        count = seen.get(base, 0) + 1
        seen[base] = count
        result.append(base if count == 1 else f"{base}_{count}")
    return result


def build_column_map(original_columns: list[str]) -> dict[str, str]:
    """Map normalized column names to their original FM headers."""
    slugs = [slugify_column(col) for col in original_columns]
    normalized = deduplicate_columns(slugs)
    return dict(zip(normalized, original_columns, strict=True))


def is_empty_value(value: Any) -> bool:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return True
    text = str(value).strip()
    return text.lower() in _EMPTY_TOKENS


def parse_currency(value: Any) -> float | None:
    """Parse FM currency strings like '£12.5M' or '$500K' into a float."""
    if is_empty_value(value):
        return None
    if isinstance(value, (int, float)) and not pd.isna(value):
        return float(value)

    text = str(value).strip().replace(",", "")
    match = _CURRENCY_RE.match(text)
    if not match:
        return None

    amount = float(match.group(1))
    suffix = (match.group(2) or "").lower()
    multiplier = {"k": 1_000, "m": 1_000_000, "b": 1_000_000_000}.get(suffix, 1)
    return amount * multiplier


def parse_wage(value: Any) -> float | None:
    """Parse wage strings like '£50 p/w' or '200k p/a'."""
    if is_empty_value(value):
        return None
    if isinstance(value, (int, float)) and not pd.isna(value):
        return float(value)

    text = str(value).strip().replace(",", "")
    match = _WAGE_RE.match(text)
    if not match:
        return parse_currency(value)

    amount = float(match.group(1))
    suffix = (match.group(2) or "").lower()
    multiplier = {"k": 1_000, "m": 1_000_000, "b": 1_000_000_000}.get(suffix, 1)
    return amount * multiplier


def parse_star_rating(value: Any) -> float | None:
    """Parse FM star ratings from unicode stars or numeric strings."""
    if is_empty_value(value):
        return None
    if isinstance(value, (int, float)) and not pd.isna(value):
        return float(value)

    text = str(value).strip()
    if _STAR_RE.search(text):
        full = text.count("★")
        empty = text.count("☆")
        half = text.count(_HALF_STAR)
        if full + empty + half == 0:
            return None
        return full + (half * 0.5)

    try:
        return float(text)
    except ValueError:
        return None


def parse_date(value: Any) -> str | None:
    """Parse common FM date strings to ISO format (YYYY-MM-DD) when possible."""
    if is_empty_value(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.date().isoformat()

    text = str(value).strip()
    for pattern in _DATE_PATTERNS:
        try:
            parsed = pd.to_datetime(text, format=pattern, errors="raise")
            return parsed.date().isoformat()
        except (ValueError, TypeError):
            continue

    try:
        parsed = pd.to_datetime(text, errors="raise", dayfirst=True)
        return parsed.date().isoformat()
    except (ValueError, TypeError):
        return text


def parse_integer(value: Any) -> int | None:
    if is_empty_value(value):
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float) and not pd.isna(value):
        return int(value)

    text = str(value).strip().replace(",", "")
    try:
        return int(float(text))
    except ValueError:
        return None


_CURRENCY_COLUMNS = frozenset(
    {
        "value",
        "asking_price",
        "release_clause",
        "transfer_value",
        "estimated_value",
    }
)
_WAGE_COLUMNS = frozenset({"wage", "wages", "salary"})
_STAR_COLUMNS = frozenset(
    {
        "ability",
        "potential",
        "current_ability",
        "potential_ability",
        "ca",
        "pa",
    }
)
_DATE_COLUMNS = frozenset({"contract", "contract_expiry", "expires", "expiry", "dob", "born"})
_INT_COLUMNS = frozenset({"age", "apps", "appearances", "goals", "assists", "height", "weight"})


def normalize_cell(column: str, value: Any) -> Any:
    """Normalize a single cell based on its column name and raw value."""
    if is_empty_value(value):
        return None

    col = slugify_column(column)

    if col in _CURRENCY_COLUMNS:
        return parse_currency(value)
    if col in _WAGE_COLUMNS:
        return parse_wage(value)
    if col in _STAR_COLUMNS:
        return parse_star_rating(value)
    if col in _DATE_COLUMNS:
        return parse_date(value)
    if col in _INT_COLUMNS:
        return parse_integer(value)

    text = str(value).strip()
    return text if text else None


def normalize_dataframe(df: pd.DataFrame, column_map: dict[str, str]) -> pd.DataFrame:
    """Return a copy of df with normalized cell values."""
    normalized = df.astype(object).copy()
    for norm_col, orig_col in column_map.items():
        if norm_col in normalized.columns:
            normalized[norm_col] = pd.array(
                [normalize_cell(orig_col, v) for v in normalized[norm_col].tolist()],
                dtype=object,
            )
    return normalized.astype(object)
