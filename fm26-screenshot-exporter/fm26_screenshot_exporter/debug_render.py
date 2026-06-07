"""Debug overlay rendering for OCR and table reconstruction."""

from __future__ import annotations

from pathlib import Path

import cv2
import numpy as np

from fm26_screenshot_exporter.config import FixedColumn, OcrToken, ParseResult, Profile, TableRegions
from fm26_screenshot_exporter.image_preprocess import load_image, save_image


def _draw_dashed_rect(
    image: np.ndarray,
    pt1: tuple[int, int],
    pt2: tuple[int, int],
    color: tuple[int, int, int],
    thickness: int = 1,
    dash: int = 8,
) -> None:
    x1, y1 = pt1
    x2, y2 = pt2
    for x in range(x1, x2, dash * 2):
        cv2.line(image, (x, y1), (min(x + dash, x2), y1), color, thickness)
        cv2.line(image, (x, y2), (min(x + dash, x2), y2), color, thickness)
    for y in range(y1, y2, dash * 2):
        cv2.line(image, (x1, y), (x1, min(y + dash, y2)), color, thickness)
        cv2.line(image, (x2, y), (x2, min(y + dash, y2)), color, thickness)


def render_debug_overlay(
    image: np.ndarray,
    *,
    table_regions: TableRegions,
    columns: list[FixedColumn],
    tokens: list[OcrToken],
    parsed_rows: list | None = None,
    profile: Profile | None = None,
    min_confidence: float = 50.0,
) -> np.ndarray:
    """Draw table crop, rows, columns, OCR boxes, and row numbers."""
    canvas = image.copy()
    if len(canvas.shape) == 2:
        canvas = cv2.cvtColor(canvas, cv2.COLOR_GRAY2BGR)

    x1, y1, x2, y2 = table_regions.table_bbox
    cv2.rectangle(canvas, (x1, y1), (x2, y2), (0, 255, 255), 2)
    cv2.putText(
        canvas,
        "table crop",
        (x1 + 4, max(y1 - 6, 12)),
        cv2.FONT_HERSHEY_SIMPLEX,
        0.5,
        (0, 255, 255),
        1,
        cv2.LINE_AA,
    )

    if table_regions.header_bbox:
        hx1, hy1, hx2, hy2 = table_regions.header_bbox
        cv2.rectangle(canvas, (hx1, hy1), (hx2, hy2), (255, 200, 0), 2)
        cv2.putText(
            canvas,
            "header",
            (hx1 + 4, hy1 + 14),
            cv2.FONT_HERSHEY_SIMPLEX,
            0.45,
            (255, 200, 0),
            1,
            cv2.LINE_AA,
        )

    for y in table_regions.row_y_boundaries:
        cv2.line(canvas, (x1, y), (x2, y), (0, 200, 0), 1)

    for col in columns:
        cv2.line(canvas, (col.x_start, y1), (col.x_start, y2), (255, 0, 255), 1)
        cv2.line(canvas, (col.x_end, y1), (col.x_end, y2), (255, 0, 255), 1)
        mid_x = (col.x_start + col.x_end) // 2
        cv2.putText(
            canvas,
            col.name[:12],
            (mid_x - 20, y1 + 16),
            cv2.FONT_HERSHEY_SIMPLEX,
            0.4,
            (255, 0, 255),
            1,
            cv2.LINE_AA,
        )

    threshold = profile.min_confidence if profile else min_confidence
    for token in tokens:
        color = (0, 180, 255) if token.confidence >= threshold else (0, 0, 255)
        tx1, ty1 = token.x, token.y
        tx2, ty2 = token.x + token.width, token.y + token.height
        cv2.rectangle(canvas, (tx1, ty1), (tx2, ty2), color, 1)
        if token.confidence < threshold:
            label = f"{token.confidence:.0f}"
            cv2.putText(
                canvas,
                label,
                (tx1, max(ty1 - 2, 10)),
                cv2.FONT_HERSHEY_SIMPLEX,
                0.35,
                (0, 0, 255),
                1,
                cv2.LINE_AA,
            )

    if parsed_rows:
        for row in parsed_rows:
            row_tokens = [
                t
                for t in tokens
                if any(
                    col_name in row.cells and row.cells[col_name]
                    for col_name in row.cells
                )
            ]
            if not row_tokens:
                continue
            row_y = int(sum(t.y_center for t in row_tokens) / len(row_tokens))
            cv2.putText(
                canvas,
                f"R{row.row_index}",
                (x1 - 40 if x1 > 40 else x1 + 4, row_y),
                cv2.FONT_HERSHEY_SIMPLEX,
                0.4,
                (200, 200, 0),
                1,
                cv2.LINE_AA,
            )

    return canvas


def render_debug_from_result(
    image_path: str | Path,
    result: ParseResult,
    profile: Profile,
    columns: list[FixedColumn],
    output: str | Path,
) -> Path:
    image = load_image(image_path)
    overlay = render_debug_overlay(
        image,
        table_regions=result.table_regions,
        columns=columns,
        tokens=result.ocr_tokens,
        parsed_rows=result.rows,
        profile=profile,
    )
    return save_image(overlay, output)
