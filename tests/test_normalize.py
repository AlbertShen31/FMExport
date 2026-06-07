"""Tests for column and value normalization."""

from __future__ import annotations

import pytest

from fm26_export_tool.normalize import (
    build_column_map,
    deduplicate_columns,
    normalize_cell,
    parse_currency,
    parse_date,
    parse_star_rating,
    parse_wage,
    slugify_column,
)


class TestSlugifyColumn:
    def test_basic(self):
        assert slugify_column("Player Name") == "player_name"

    def test_special_chars(self):
        assert slugify_column("CA/PA") == "capa"

    def test_empty_fallback(self):
        assert slugify_column("   ") == "column"


class TestDeduplicateColumns:
    def test_unique_unchanged(self):
        assert deduplicate_columns(["name", "age"]) == ["name", "age"]

    def test_duplicates_get_suffix(self):
        assert deduplicate_columns(["value", "value", "value"]) == [
            "value",
            "value_2",
            "value_3",
        ]


class TestBuildColumnMap:
    def test_preserves_originals(self):
        originals = ["Name", "Value", "Value"]
        mapping = build_column_map(originals)
        assert mapping == {
            "name": "Name",
            "value": "Value",
            "value_2": "Value",
        }


class TestParseCurrency:
    @pytest.mark.parametrize(
        ("raw", "expected"),
        [
            ("£12.5M", 12_500_000.0),
            ("$500K", 500_000.0),
            ("€2.3M", 2_300_000.0),
            ("-", None),
            ("N/A", None),
        ],
    )
    def test_currency_values(self, raw, expected):
        assert parse_currency(raw) == expected


class TestParseWage:
    def test_weekly_wage(self):
        assert parse_wage("£50K p/w") == 50_000.0

    def test_small_weekly(self):
        assert parse_wage("£200 p/w") == 200.0


class TestParseStarRating:
    def test_unicode_stars(self):
        assert parse_star_rating("★★★★☆") == 4.0

    def test_numeric(self):
        assert parse_star_rating("3.5") == 3.5

    def test_empty(self):
        assert parse_star_rating("-") is None


class TestParseDate:
    def test_dmy(self):
        assert parse_date("30/06/2028") == "2028-06-30"

    def test_empty(self):
        assert parse_date("N/A") is None


class TestNormalizeCell:
    def test_age_integer(self):
        assert normalize_cell("Age", "24") == 24

    def test_empty_dash(self):
        assert normalize_cell("Name", "-") is None
