"""Reconstruct table rows and columns from OCR tokens."""

from __future__ import annotations

from dataclasses import dataclass, field

from rapidfuzz import fuzz

from fm26_screenshot_exporter.config import (
    ColumnStrategy,
    FixedColumn,
    HeaderYRange,
    OcrToken,
    ParsedRow,
    Profile,
    TableRegions,
)
from fm26_screenshot_exporter.normalize import apply_normalizer, clean_text, slugify_column


@dataclass
class CellAssignment:
    column_name: str
    tokens: list[OcrToken] = field(default_factory=list)

    @property
    def raw_text(self) -> str:
        ordered = sorted(self.tokens, key=lambda t: t.x)
        return clean_text(" ".join(t.text for t in ordered))

    @property
    def avg_confidence(self) -> float | None:
        if not self.tokens:
            return None
        return sum(t.confidence for t in self.tokens) / len(self.tokens)


def cluster_tokens_by_y(
    tokens: list[OcrToken],
    tolerance: int,
) -> list[list[OcrToken]]:
    """Group OCR tokens into rows by y-center clustering."""
    if not tokens:
        return []

    sorted_tokens = sorted(tokens, key=lambda t: t.y_center)
    clusters: list[list[OcrToken]] = []
    current: list[OcrToken] = [sorted_tokens[0]]
    current_y = sorted_tokens[0].y_center

    for token in sorted_tokens[1:]:
        if abs(token.y_center - current_y) <= tolerance:
            current.append(token)
            current_y = sum(t.y_center for t in current) / len(current)
        else:
            clusters.append(current)
            current = [token]
            current_y = token.y_center
    clusters.append(current)
    return clusters


def assign_token_to_column(
    token: OcrToken,
    columns: list[FixedColumn],
) -> str | None:
    """Assign a token to a column where x_center falls within x_start/x_end."""
    x = token.x_center
    for col in columns:
        if col.x_start <= x < col.x_end:
            return col.name
    return None


def _columns_from_fixed(profile: Profile) -> list[FixedColumn]:
    if profile.fixed_columns:
        return profile.fixed_columns
    raise ValueError(
        "No fixed_columns defined in profile. "
        "Create a profile with fixed column x-boundaries for reliable extraction."
    )


def _infer_columns_from_header(
    tokens: list[OcrToken],
    header_range: HeaderYRange | None,
    expected_columns: list[str],
    tolerance: int = 15,
) -> list[FixedColumn]:
    if not header_range:
        raise ValueError("header_y_range required for from_header_ocr column strategy")

    header_tokens = [
        t
        for t in tokens
        if header_range.y_start <= t.y_center <= header_range.y_end
    ]
    if not header_tokens:
        raise ValueError("No OCR tokens found in header row")

    header_clusters = cluster_tokens_by_y(header_tokens, tolerance)
    header_cells: list[tuple[str, float]] = []
    for cluster in header_clusters:
        ordered = sorted(cluster, key=lambda t: t.x)
        text = clean_text(" ".join(t.text for t in ordered))
        x_center = sum(t.x_center for t in cluster) / len(cluster)
        header_cells.append((text, x_center))

    if not expected_columns:
        expected_columns = [text for text, _ in sorted(header_cells, key=lambda c: c[1])]

    columns: list[FixedColumn] = []
    used: set[int] = set()

    for expected in expected_columns:
        best_idx = -1
        best_score = 0
        for idx, (text, x_center) in enumerate(header_cells):
            if idx in used:
                continue
            score = fuzz.ratio(expected.lower(), text.lower())
            if score > best_score:
                best_score = score
                best_idx = idx

        if best_idx < 0:
            continue
        used.add(best_idx)
        _, x_center = header_cells[best_idx]

        x_start = int(x_center - 40)
        if columns:
            prev = columns[-1]
            midpoint = (prev.x_end + x_start) // 2
            columns[-1] = FixedColumn(name=prev.name, x_start=prev.x_start, x_end=midpoint)
            x_start = midpoint

        columns.append(FixedColumn(name=expected, x_start=x_start, x_end=int(x_center + 40)))

    if columns:
        last = columns[-1]
        max_x = int(max(t.x + t.width for t in header_tokens) + 20)
        columns[-1] = FixedColumn(name=last.name, x_start=last.x_start, x_end=max_x)

    return columns


def _infer_columns_from_alignment(
    tokens: list[OcrToken],
    table_regions: TableRegions,
) -> list[FixedColumn]:
    """Infer column boundaries from vertical alignment of token x-centers."""
    if not tokens:
        raise ValueError("No tokens available for column inference")

    x_centers = sorted({int(t.x_center) for t in tokens})
    if len(x_centers) < 2:
        raise ValueError("Insufficient x-alignment data for column inference")

    gaps: list[tuple[int, int]] = []
    for i in range(len(x_centers) - 1):
        gap = x_centers[i + 1] - x_centers[i]
        gaps.append((gap, i))

    large_gaps = sorted(gaps, reverse=True)
    split_indices = sorted(idx for _, idx in large_gaps[: min(10, len(large_gaps))])

    boundaries = [table_regions.table_bbox[0]]
    for idx in split_indices:
        boundary = (x_centers[idx] + x_centers[idx + 1]) // 2
        if boundary not in boundaries:
            boundaries.append(boundary)
    boundaries.append(table_regions.table_bbox[2])
    boundaries = sorted(set(boundaries))

    columns: list[FixedColumn] = []
    for i in range(len(boundaries) - 1):
        columns.append(
            FixedColumn(
                name=f"column_{i + 1}",
                x_start=boundaries[i],
                x_end=boundaries[i + 1],
            )
        )
    return columns


def resolve_columns(
    profile: Profile,
    tokens: list[OcrToken],
    table_regions: TableRegions,
) -> list[FixedColumn]:
    if profile.column_strategy == ColumnStrategy.FIXED_BOUNDARIES or profile.fixed_columns:
        return _columns_from_fixed(profile)
    if profile.column_strategy == ColumnStrategy.FROM_HEADER_OCR:
        return _infer_columns_from_header(
            tokens,
            profile.header_y_range,
            profile.expected_columns,
            profile.row_y_tolerance,
        )
    if profile.column_strategy == ColumnStrategy.INFER_FROM_VERTICAL_ALIGNMENT:
        return _infer_columns_from_alignment(tokens, table_regions)
    return _columns_from_fixed(profile)


def filter_tokens_in_table(
    tokens: list[OcrToken],
    table_regions: TableRegions,
    *,
    exclude_header: HeaderYRange | None = None,
) -> list[OcrToken]:
    x1, y1, x2, y2 = table_regions.table_bbox
    filtered: list[OcrToken] = []
    for token in tokens:
        if not (x1 <= token.x_center < x2 and y1 <= token.y_center < y2):
            continue
        if exclude_header and exclude_header.y_start <= token.y_center <= exclude_header.y_end:
            continue
        filtered.append(token)
    return filtered


def reconstruct_table(
    tokens: list[OcrToken],
    profile: Profile,
    table_regions: TableRegions,
) -> tuple[list[ParsedRow], list[FixedColumn]]:
    """
    Assign tokens to cells, group into rows, and produce parsed rows with
    raw and normalized cell values.
    """
    columns = resolve_columns(profile, tokens, table_regions)
    data_tokens = filter_tokens_in_table(
        tokens,
        table_regions,
        exclude_header=profile.header_y_range,
    )

    row_clusters = cluster_tokens_by_y(data_tokens, profile.row_y_tolerance)
    parsed_rows: list[ParsedRow] = []

    for row_index, cluster in enumerate(row_clusters):
        cell_map: dict[str, CellAssignment] = {c.name: CellAssignment(c.name) for c in columns}

        for token in cluster:
            col_name = assign_token_to_column(token, columns)
            if col_name and col_name in cell_map:
                cell_map[col_name].tokens.append(token)

        raw_cells = {name: cell.raw_text for name, cell in cell_map.items()}
        if not any(raw_cells.values()):
            continue

        confidences = [
            cell.avg_confidence for cell in cell_map.values() if cell.avg_confidence is not None
        ]
        avg_conf = sum(confidences) / len(confidences) if confidences else None
        low_conf = avg_conf is not None and avg_conf < profile.min_confidence

        normalized: dict[str, object | None] = {}
        for col_name, raw in raw_cells.items():
            slug = slugify_column(col_name)
            normalizer_key = profile.value_normalizers.get(col_name) or profile.value_normalizers.get(
                slug
            )
            if normalizer_key:
                normalized[col_name] = apply_normalizer(normalizer_key, raw)
            else:
                normalized[col_name] = clean_text(raw) or None

        parsed_rows.append(
            ParsedRow(
                row_index=row_index,
                cells=raw_cells,
                normalized_cells=normalized,
                avg_confidence=avg_conf,
                low_confidence=low_conf,
            )
        )

    return parsed_rows, columns
