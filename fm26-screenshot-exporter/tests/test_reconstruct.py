"""Tests for row/column reconstruction and exporters."""

from __future__ import annotations

import pytest

from fm26_screenshot_exporter.config import FixedColumn, OcrToken, Profile, TableRegions
from fm26_screenshot_exporter.exporters import escape_csv_cell
from fm26_screenshot_exporter.ocr import filter_tokens_by_confidence
from fm26_screenshot_exporter.row_column_reconstruction import (
    assign_token_to_column,
    cluster_tokens_by_y,
    reconstruct_table,
)


def _token(text: str, x: int, y: int, conf: float = 90.0) -> OcrToken:
    return OcrToken(
        text=text,
        confidence=conf,
        x=x,
        y=y,
        width=40,
        height=14,
        line_num=1,
        block_num=1,
        par_num=1,
    )


class TestClusterTokensByY:
    def test_groups_nearby_y(self):
        tokens = [
            _token("Alice", 10, 100),
            _token("Smith", 60, 102),
            _token("Bob", 10, 140),
        ]
        clusters = cluster_tokens_by_y(tokens, tolerance=12)
        assert len(clusters) == 2
        assert len(clusters[0]) == 2
        assert len(clusters[1]) == 1

    def test_separates_distant_y(self):
        tokens = [
            _token("A", 10, 50),
            _token("B", 10, 200),
        ]
        clusters = cluster_tokens_by_y(tokens, tolerance=10)
        assert len(clusters) == 2


class TestAssignTokenToColumn:
    def test_assigns_by_x_center(self):
        columns = [
            FixedColumn(name="Name", x_start=0, x_end=100),
            FixedColumn(name="Age", x_start=100, x_end=150),
        ]
        assert assign_token_to_column(_token("24", 110, 50), columns) == "Age"
        assert assign_token_to_column(_token("John", 20, 50), columns) == "Name"

    def test_outside_returns_none(self):
        columns = [FixedColumn(name="Name", x_start=0, x_end=100)]
        assert assign_token_to_column(_token("X", 200, 50), columns) is None


class TestReconstructTable:
    def test_builds_rows_from_fixed_columns(self):
        profile = Profile(
            name="test",
            manual_crop=None,
            column_strategy="fixed_boundaries",
            fixed_columns=[
                FixedColumn(name="Name", x_start=0, x_end=120),
                FixedColumn(name="Age", x_start=120, x_end=180),
            ],
            header_y_range=None,
            row_y_tolerance=12,
            min_confidence=0,
        )
        regions = TableRegions(table_bbox=(0, 0, 200, 200))
        tokens = [
            _token("Erling", 10, 80),
            _token("Haaland", 55, 82),
            _token("24", 130, 80),
            _token("Kevin", 10, 120),
            _token("De", 45, 122),
            _token("Bruyne", 70, 121),
            _token("33", 135, 120),
        ]
        rows, columns = reconstruct_table(tokens, profile, regions)
        assert len(rows) == 2
        assert "Erling" in rows[0].cells["Name"]
        assert rows[0].cells["Age"] == "24"
        assert rows[1].cells["Age"] == "33"


class TestLowConfidenceFiltering:
    def test_filters_below_threshold(self):
        tokens = [
            _token("Good", 10, 10, conf=90),
            _token("Bad", 60, 10, conf=20),
        ]
        filtered = filter_tokens_by_confidence(tokens, min_confidence=50)
        assert len(filtered) == 1
        assert filtered[0].text == "Good"


class TestCsvEscaping:
    def test_quotes_commas(self):
        assert escape_csv_cell('Say "Hi", world') == '"Say ""Hi"", world"'

    def test_plain_text(self):
        assert escape_csv_cell("simple") == "simple"
