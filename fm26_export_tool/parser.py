"""Parse Football Manager Web Page HTML exports into structured tables."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from io import StringIO
from pathlib import Path
from typing import Any

import pandas as pd
from bs4 import BeautifulSoup

from fm26_export_tool.normalize import build_column_map, normalize_dataframe


@dataclass
class ParsedExport:
    """Result of parsing an FM HTML export."""

    source_path: Path
    dataframe: pd.DataFrame
    column_map: dict[str, str]
    table_index: int = 0
    metadata: dict[str, Any] = field(default_factory=dict)

    @property
    def original_columns(self) -> list[str]:
        return list(self.column_map.values())

    @property
    def normalized_columns(self) -> list[str]:
        return list(self.column_map.keys())

    def to_records(self) -> list[dict[str, Any]]:
        return self.dataframe.to_dict(orient="records")


def _clean_header(text: str) -> str:
    return " ".join(str(text).split())


def _extract_tables_from_soup(soup: BeautifulSoup) -> list[pd.DataFrame]:
    """Extract HTML tables using pandas, with BeautifulSoup preprocessing."""
    for table in soup.find_all("table"):
        for cell in table.find_all(["th", "td"]):
            cell.string = _clean_header(cell.get_text())

    html = str(soup)
    try:
        tables = pd.read_html(StringIO(html), flavor="lxml")
    except ValueError as exc:
        if "No tables found" in str(exc):
            return []
        raise
    return tables


def _select_player_table(tables: list[pd.DataFrame]) -> tuple[int, pd.DataFrame]:
    """Pick the table most likely to be an FM player list."""
    if not tables:
        raise ValueError("No HTML tables found in export.")

    best_idx = 0
    best_score = -1
    player_keywords = {
        "name",
        "age",
        "position",
        "club",
        "nationality",
        "value",
        "wage",
        "ability",
        "potential",
    }

    for idx, table in enumerate(tables):
        headers = {str(c).strip().lower() for c in table.columns}
        score = len(headers & player_keywords)
        score += len(table) * 0.01
        if score > best_score:
            best_score = score
            best_idx = idx

    return best_idx, tables[best_idx]


def _pandas_original_column(name: str) -> str:
    """Recover FM header text from pandas duplicate suffixes (e.g. Value.1)."""
    text = str(name)
    if re.match(r"^.+\.\d+$", text):
        return text.rsplit(".", 1)[0]
    return text


def _rename_columns(df: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, str]]:
    original_columns = [_pandas_original_column(c) for c in df.columns]
    column_map = build_column_map(original_columns)
    renamed = df.copy()
    renamed.columns = list(column_map.keys())
    return renamed, column_map


def parse_html(
    path: str | Path,
    *,
    table_index: int | None = None,
    normalize: bool = True,
) -> ParsedExport:
    """Parse an FM-exported HTML file into a ParsedExport."""
    source = Path(path)
    if not source.exists():
        raise FileNotFoundError(f"HTML file not found: {source}")

    html = source.read_text(encoding="utf-8", errors="replace")
    soup = BeautifulSoup(html, "lxml")
    tables = _extract_tables_from_soup(soup)

    if table_index is not None:
        if table_index < 0 or table_index >= len(tables):
            raise ValueError(
                f"table_index {table_index} out of range (found {len(tables)} tables)"
            )
        idx, raw_df = table_index, tables[table_index]
    else:
        idx, raw_df = _select_player_table(tables)

    renamed, column_map = _rename_columns(raw_df)
    df = normalize_dataframe(renamed, column_map) if normalize else renamed

    title_tag = soup.find("title")
    metadata: dict[str, Any] = {
        "title": title_tag.get_text(strip=True) if title_tag else None,
        "table_count": len(tables),
        "row_count": len(df),
    }

    return ParsedExport(
        source_path=source,
        dataframe=df,
        column_map=column_map,
        table_index=idx,
        metadata=metadata,
    )


def parse_html_string(
    html: str,
    *,
    table_index: int | None = None,
    normalize: bool = True,
    source_name: str = "<string>",
) -> ParsedExport:
    """Parse FM HTML content from a string (useful for tests)."""
    soup = BeautifulSoup(html, "lxml")
    tables = _extract_tables_from_soup(soup)

    if table_index is not None:
        if table_index < 0 or table_index >= len(tables):
            raise ValueError(
                f"table_index {table_index} out of range (found {len(tables)} tables)"
            )
        idx, raw_df = table_index, tables[table_index]
    else:
        idx, raw_df = _select_player_table(tables)

    renamed, column_map = _rename_columns(raw_df)
    df = normalize_dataframe(renamed, column_map) if normalize else renamed

    return ParsedExport(
        source_path=Path(source_name),
        dataframe=df,
        column_map=column_map,
        table_index=idx,
        metadata={"row_count": len(df), "table_count": len(tables)},
    )
