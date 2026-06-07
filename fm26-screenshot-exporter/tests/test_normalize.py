"""Tests for OCR text normalization."""

from __future__ import annotations

import pytest

from fm26_screenshot_exporter.normalize import (
    clean_text,
    parse_currency,
    parse_percentage,
    parse_wage,
    slugify_column,
)


class TestCleanText:
    def test_strips_and_collapses(self):
        assert clean_text("  Hello   World  ") == "Hello World"


class TestSlugifyColumn:
    def test_basic(self):
        assert slugify_column("Transfer Value") == "transfer_value"


class TestParseCurrency:
    @pytest.mark.parametrize(
        ("raw", "expected"),
        [
            ("£12.5M", 12_500_000.0),
            ("$500K", 500_000.0),
            ("€2.3M", 2_300_000.0),
            ("-", None),
        ],
    )
    def test_values(self, raw, expected):
        assert parse_currency(raw) == expected


class TestParseWage:
    def test_weekly(self):
        assert parse_wage("£50K p/w") == 50_000.0

    def test_plain_currency_fallback(self):
        assert parse_wage("£200") == 200.0


class TestParsePercentage:
    @pytest.mark.parametrize(
        ("raw", "expected"),
        [
            ("85%", 85.0),
            ("72.5", 72.5),
            ("-", None),
        ],
    )
    def test_values(self, raw, expected):
        assert parse_percentage(raw) == expected
