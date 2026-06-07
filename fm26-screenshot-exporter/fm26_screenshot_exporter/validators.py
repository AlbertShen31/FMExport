"""Validation warnings for parsed OCR tables."""

from __future__ import annotations

from dataclasses import dataclass, field

from fm26_screenshot_exporter.config import ParseResult, Profile
from fm26_screenshot_exporter.normalize import parse_float, slugify_column


@dataclass
class ValidationWarning:
    code: str
    message: str
    severity: str = "warning"


@dataclass
class ValidationReport:
    warnings: list[ValidationWarning] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not any(w.severity == "error" for w in self.warnings)

    def add(self, code: str, message: str, *, severity: str = "warning") -> None:
        self.warnings.append(ValidationWarning(code=code, message=message, severity=severity))


def validate_parse_result(result: ParseResult, profile: Profile) -> ValidationReport:
    report = ValidationReport()

    present_columns = set(result.columns)
    required = profile.required_columns or profile.expected_columns
    missing = [col for col in required if col not in present_columns]
    if missing:
        report.add(
            "missing_columns",
            f"Required columns missing from output: {', '.join(missing)}",
        )

    if not result.rows:
        report.add("no_rows", "No data rows were extracted.", severity="error")
        return report

    total_cells = len(result.rows) * max(len(result.columns), 1)
    empty_cells = sum(
        1
        for row in result.rows
        for col in result.columns
        if not (row.cells.get(col) or "").strip()
    )
    empty_ratio = empty_cells / total_cells if total_cells else 1.0
    if empty_ratio > profile.max_empty_cell_ratio:
        report.add(
            "too_many_empty_cells",
            f"Empty cell ratio {empty_ratio:.1%} exceeds threshold "
            f"{profile.max_empty_cell_ratio:.1%}.",
        )

    confidences = [r.avg_confidence for r in result.rows if r.avg_confidence is not None]
    if confidences:
        avg_conf = sum(confidences) / len(confidences)
        if avg_conf < profile.min_avg_confidence:
            report.add(
                "low_avg_confidence",
                f"Average row confidence {avg_conf:.1f} is below "
                f"threshold {profile.min_avg_confidence:.1f}.",
            )

    if len(result.rows) < profile.min_rows:
        report.add(
            "low_row_count",
            f"Only {len(result.rows)} rows extracted; expected at least {profile.min_rows}.",
        )

    name_col = _find_name_column(result.columns)
    if name_col:
        names = [row.cells.get(name_col, "").strip().lower() for row in result.rows]
        names = [n for n in names if n]
        duplicates = {n for n in names if names.count(n) > 1}
        if duplicates:
            sample = ", ".join(sorted(duplicates)[:5])
            report.add(
                "duplicate_names",
                f"Duplicate player names detected: {sample}",
            )

    for col in result.columns:
        normalizer = profile.value_normalizers.get(col) or profile.value_normalizers.get(
            slugify_column(col)
        )
        if normalizer not in {"number", "age", "percentage", "currency", "wage", "star_rating"}:
            continue

        bad_values: list[str] = []
        for row in result.rows:
            raw = row.cells.get(col, "")
            if not raw.strip():
                continue
            normalized = row.normalized_cells.get(col)
            if normalized is None and parse_float(raw) is None:
                bad_values.append(raw)
        if bad_values:
            sample = ", ".join(bad_values[:5])
            report.add(
                "non_numeric_values",
                f"Column '{col}' has non-numeric OCR values: {sample}",
            )

    low_conf_rows = [r.row_index for r in result.rows if r.low_confidence]
    if low_conf_rows:
        report.add(
            "low_confidence_rows",
            f"Rows flagged for low OCR confidence: {low_conf_rows[:10]}",
        )

    return report


def _find_name_column(columns: list[str]) -> str | None:
    for col in columns:
        if col.lower() in {"name", "player", "player name"}:
            return col
    for col in columns:
        if "name" in col.lower():
            return col
    return None
