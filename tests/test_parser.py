"""Tests for FM HTML parsing and exporters."""

from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
import pytest

from fm26_export_tool.exporters import export_csv, export_json, export_xlsx
from fm26_export_tool.parser import parse_html_string


class TestParseHtml:
    def test_parses_player_table(self, sample_fm_html: str):
        parsed = parse_html_string(sample_fm_html)
        assert len(parsed.dataframe) == 3
        assert "name" in parsed.column_map
        assert parsed.column_map["name"] == "Name"

    def test_duplicate_columns_renamed(self, sample_fm_html: str):
        parsed = parse_html_string(sample_fm_html)
        assert "value" in parsed.column_map
        assert "value_2" in parsed.column_map

    def test_normalizes_currency(self, sample_fm_html: str):
        parsed = parse_html_string(sample_fm_html)
        john = parsed.dataframe.iloc[0]
        assert john["value"] == 12_500_000.0

    def test_normalizes_wage(self, sample_fm_html: str):
        parsed = parse_html_string(sample_fm_html)
        john = parsed.dataframe.iloc[0]
        assert john["wage"] == 50_000.0

    def test_normalizes_stars(self, sample_fm_html: str):
        parsed = parse_html_string(sample_fm_html)
        john = parsed.dataframe.iloc[0]
        assert john["ability"] == 4.0
        assert john["potential"] == 5.0

    def test_empty_values_become_none(self, sample_fm_html: str):
        parsed = parse_html_string(sample_fm_html)
        jane = parsed.dataframe.iloc[1]
        assert jane["age"] is None
        assert jane["contract"] is None

    def test_no_tables_raises(self):
        with pytest.raises(ValueError, match="No HTML tables"):
            parse_html_string("<html><body><p>No table</p></body></html>")


class TestExporters:
    def test_export_csv(self, sample_fm_html: str, tmp_path: Path):
        parsed = parse_html_string(sample_fm_html)
        out = export_csv(parsed, tmp_path / "players.csv")
        text = out.read_text(encoding="utf-8")
        assert "name" in text
        assert "John Smith" in text

    def test_export_xlsx(self, sample_fm_html: str, tmp_path: Path):
        parsed = parse_html_string(sample_fm_html)
        out = export_xlsx(parsed, tmp_path / "players.xlsx")
        assert out.exists()
        df = pd.read_excel(out)
        assert len(df) == 3

    def test_export_json(self, sample_fm_html: str, tmp_path: Path):
        parsed = parse_html_string(sample_fm_html)
        out = export_json(parsed, tmp_path / "players.json")
        payload = json.loads(out.read_text(encoding="utf-8"))
        assert len(payload["records"]) == 3
        assert "columns" in payload
