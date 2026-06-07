"""Table region detection using OpenCV and profile-based manual crop."""

from __future__ import annotations

import json
from pathlib import Path

import cv2
import numpy as np

from fm26_screenshot_exporter.config import CropRegion, HeaderYRange, Profile, TableRegions
from fm26_screenshot_exporter.image_preprocess import load_image


def _bbox_from_crop(crop: CropRegion) -> tuple[int, int, int, int]:
    return (crop.x, crop.y, crop.x + crop.width, crop.y + crop.height)


def detect_horizontal_lines(gray: np.ndarray, min_length_ratio: float = 0.3) -> list[int]:
    """Detect horizontal line y-positions via morphological operations."""
    h, w = gray.shape
    min_length = int(w * min_length_ratio)
    binary = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 15, 5
    )
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (min_length, 1))
    lines_img = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel, iterations=1)
    contours, _ = cv2.findContours(lines_img, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    y_positions: list[int] = []
    for contour in contours:
        x, y, cw, ch = cv2.boundingRect(contour)
        if cw >= min_length and ch <= 6:
            y_positions.append(y + ch // 2)
    return sorted(set(y_positions))


def detect_vertical_lines(gray: np.ndarray, min_length_ratio: float = 0.15) -> list[int]:
    """Detect vertical line x-positions via morphological operations."""
    h, w = gray.shape
    min_length = int(h * min_length_ratio)
    binary = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 15, 5
    )
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, min_length))
    lines_img = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel, iterations=1)
    contours, _ = cv2.findContours(lines_img, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    x_positions: list[int] = []
    for contour in contours:
        x, y, cw, ch = cv2.boundingRect(contour)
        if ch >= min_length and cw <= 6:
            x_positions.append(x + cw // 2)
    return sorted(set(x_positions))


def detect_table_regions(
    image: np.ndarray,
    profile: Profile | None = None,
) -> TableRegions:
    """
    Detect likely table area using profile manual crop and/or OpenCV line detection.
    Fixed column boundaries from profile take precedence for column_x_boundaries.
    """
    h, w = image.shape[:2]

    if profile and profile.manual_crop:
        table_bbox = _bbox_from_crop(profile.manual_crop)
    else:
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY) if len(image.shape) == 3 else image
        row_ys = detect_horizontal_lines(gray)
        col_xs = detect_vertical_lines(gray)

        if len(row_ys) >= 2 and len(col_xs) >= 2:
            table_bbox = (col_xs[0], row_ys[0], col_xs[-1], row_ys[-1])
        else:
            table_bbox = (0, 0, w, h)

    header_bbox: tuple[int, int, int, int] | None = None
    if profile and profile.header_y_range:
        hx1, hy1, hx2, hy2 = table_bbox
        header_bbox = (
            hx1,
            profile.header_y_range.y_start,
            hx2,
            profile.header_y_range.y_end,
        )

    row_y_boundaries: list[int] = []
    if profile and profile.row_height and profile.header_y_range:
        y = profile.header_y_range.y_end
        _, _, _, ty2 = table_bbox
        while y < ty2:
            row_y_boundaries.append(y)
            y += profile.row_height
        row_y_boundaries.append(ty2)
    else:
        crop = image
        if profile and profile.manual_crop:
            x1, y1, x2, y2 = table_bbox
            crop = image[y1:y2, x1:x2]
        gray = cv2.cvtColor(crop, cv2.COLOR_BGR2GRAY) if len(crop.shape) == 3 else crop
        detected = detect_horizontal_lines(gray)
        if detected:
            offset = table_bbox[1] if profile and profile.manual_crop else 0
            row_y_boundaries = [y + offset for y in detected]

    column_x_boundaries: list[int] = []
    if profile and profile.fixed_columns:
        column_x_boundaries = sorted({c.x_start for c in profile.fixed_columns})
        column_x_boundaries.append(profile.fixed_columns[-1].x_end)
    else:
        crop = image
        if profile and profile.manual_crop:
            x1, y1, x2, y2 = table_bbox
            crop = image[y1:y2, x1:x2]
        gray = cv2.cvtColor(crop, cv2.COLOR_BGR2GRAY) if len(crop.shape) == 3 else crop
        detected = detect_vertical_lines(gray)
        if detected:
            offset = table_bbox[0] if profile and profile.manual_crop else 0
            column_x_boundaries = [x + offset for x in detected]

    return TableRegions(
        table_bbox=table_bbox,
        header_bbox=header_bbox,
        row_y_boundaries=row_y_boundaries,
        column_x_boundaries=column_x_boundaries,
    )


def detect_table_from_path(
    image_path: str | Path,
    profile: Profile | None = None,
) -> TableRegions:
    image = load_image(image_path)
    return detect_table_regions(image, profile)


def save_table_regions(regions: TableRegions, path: str | Path) -> Path:
    out = Path(path)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(regions.model_dump_json(indent=2), encoding="utf-8")
    return out


def load_table_regions(path: str | Path) -> TableRegions:
    return TableRegions.model_validate_json(Path(path).read_text(encoding="utf-8"))
