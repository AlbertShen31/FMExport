"""Full screenshot parse pipeline."""

from __future__ import annotations

from pathlib import Path

from fm26_screenshot_exporter.config import ParseResult, Profile
from fm26_screenshot_exporter.debug_render import render_debug_from_result
from fm26_screenshot_exporter.exporters import export_by_extension
from fm26_screenshot_exporter.image_preprocess import (
    load_image,
    preprocess_for_ocr,
    preprocess_screenshot,
    save_image,
)
from fm26_screenshot_exporter.ocr import filter_tokens_by_confidence, run_ocr
from fm26_screenshot_exporter.profiles import load_profile
from fm26_screenshot_exporter.row_column_reconstruction import reconstruct_table
from fm26_screenshot_exporter.table_detection import detect_table_regions
from fm26_screenshot_exporter.validators import validate_parse_result


def run_preprocess(
    image_path: str | Path,
    output_path: str | Path,
    *,
    profile: Profile | None = None,
) -> Path:
    image = load_image(image_path)
    crop = profile.manual_crop if profile else None
    upscale = profile.upscale_factor if profile else 2.0
    processed = preprocess_screenshot(image, crop=crop, upscale_factor=upscale)
    return save_image(processed, output_path)


def run_ocr_pipeline(
    image_path: str | Path,
    *,
    profile: Profile | None = None,
) -> tuple[list, object]:
    image = load_image(image_path)
    crop = profile.manual_crop if profile else None
    upscale = profile.upscale_factor if profile else 2.0
    ocr_image, color_image = preprocess_for_ocr(image, crop=crop, upscale_factor=upscale)
    tokens = run_ocr(ocr_image)
    return tokens, color_image


def run_detect_table(
    image_path: str | Path,
    profile: Profile | None = None,
):
    image = load_image(image_path)
    crop = profile.manual_crop if profile else None
    upscale = profile.upscale_factor if profile else 2.0
    _, color_image = preprocess_for_ocr(image, crop=crop, upscale_factor=upscale)
    return detect_table_regions(color_image, profile)


def parse_screenshot(
    image_path: str | Path,
    profile: Profile | str,
) -> ParseResult:
    if isinstance(profile, str):
        profile = load_profile(profile)

    path = Path(image_path)
    tokens, color_image = run_ocr_pipeline(path, profile=profile)
    filtered = filter_tokens_by_confidence(tokens, profile.min_confidence)

    table_regions = detect_table_regions(color_image, profile)
    rows, columns = reconstruct_table(filtered, profile, table_regions)

    all_confidences = [t.confidence for t in filtered]
    avg_conf = sum(all_confidences) / len(all_confidences) if all_confidences else 0.0

    result = ParseResult(
        source_path=str(path),
        profile_name=profile.name,
        columns=[c.name for c in columns],
        rows=rows,
        ocr_tokens=filtered,
        table_regions=table_regions,
        metadata={
            "token_count": len(filtered),
            "avg_ocr_confidence": round(avg_conf, 2),
            "row_count": len(rows),
        },
    )

    validation = validate_parse_result(result, profile)
    result.metadata["validation_ok"] = validation.ok
    result.metadata["validation_warnings"] = [
        {"code": w.code, "message": w.message, "severity": w.severity}
        for w in validation.warnings
    ]
    return result


def parse_and_export(
    image_path: str | Path,
    profile: Profile | str,
    output: str | Path,
) -> ParseResult:
    result = parse_screenshot(image_path, profile)
    export_by_extension(result, output)
    return result


def render_debug(
    image_path: str | Path,
    profile: Profile | str,
    output: str | Path,
) -> Path:
    if isinstance(profile, str):
        profile = load_profile(profile)

    path = Path(image_path)
    result = parse_screenshot(path, profile)
    tokens, color_image = run_ocr_pipeline(path, profile=profile)

    from fm26_screenshot_exporter.debug_render import render_debug_overlay
    from fm26_screenshot_exporter.row_column_reconstruction import resolve_columns

    columns = resolve_columns(profile, result.ocr_tokens, result.table_regions)
    overlay = render_debug_overlay(
        color_image,
        table_regions=result.table_regions,
        columns=columns,
        tokens=tokens,
        parsed_rows=result.rows,
        profile=profile,
    )
    return save_image(overlay, output)
