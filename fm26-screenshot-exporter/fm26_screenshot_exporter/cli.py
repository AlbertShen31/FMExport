"""CLI for FM26 screenshot OCR exporter."""

from __future__ import annotations

import platform
import sys
from importlib import metadata
from pathlib import Path

import typer
from rich.console import Console
from rich.table import Table

from fm26_screenshot_exporter import __version__
from fm26_screenshot_exporter.ocr import check_tesseract, save_ocr_json
from fm26_screenshot_exporter.pipeline import (
    parse_and_export,
    parse_screenshot,
    render_debug,
    run_detect_table,
    run_preprocess,
)
from fm26_screenshot_exporter.profiles import list_profiles, load_profile
from fm26_screenshot_exporter.table_detection import save_table_regions
from fm26_screenshot_exporter.watcher import is_image_file, watch_folder

app = typer.Typer(
    name="fm26-ocr",
    help="Convert Football Manager 2026 table screenshots to structured data.",
    no_args_is_help=True,
)
console = Console()


def _echo_validation_warnings(result) -> None:
    warnings = result.metadata.get("validation_warnings", [])
    for item in warnings:
        severity = item.get("severity", "warning")
        style = "yellow" if severity == "warning" else "red"
        console.print(f"[{style}]{item['code']}: {item['message']}[/{style}]")


@app.command("scan-env")
def scan_env() -> None:
    """Check tesseract installation and Python/package info."""
    tess = check_tesseract()

    table = Table(title="FM26 Screenshot OCR Environment")
    table.add_column("Component", style="cyan")
    table.add_column("Status", style="green")

    if tess["installed"]:
        table.add_row("Tesseract", f"OK — {tess['version']}")
        table.add_row("Tesseract path", str(tess["path"]))
    else:
        table.add_row("Tesseract", "[red]NOT FOUND[/red]")
        console.print(
            "[yellow]Install with: brew install tesseract[/yellow]"
        )

    table.add_row("Python", sys.version.split()[0])
    table.add_row("Platform", platform.platform())
    table.add_row("fm26-screenshot-exporter", __version__)

    packages = [
        "typer",
        "rich",
        "pillow",
        "opencv-python",
        "pytesseract",
        "pandas",
        "openpyxl",
        "pydantic",
        "rapidfuzz",
        "watchdog",
        "pyyaml",
    ]
    for pkg in packages:
        try:
            version = metadata.version(pkg)
            table.add_row(pkg, version)
        except metadata.PackageNotFoundError:
            table.add_row(pkg, "[red]not installed[/red]")

    console.print(table)
    profiles = list_profiles()
    if profiles:
        console.print(f"\nAvailable profiles: {', '.join(profiles)}")


@app.command("preprocess")
def preprocess_cmd(
    image_path: Path = typer.Argument(..., help="Input screenshot path"),
    out: Path = typer.Option(..., "--out", "-o", help="Output debug image path"),
    profile: str | None = typer.Option(None, "--profile", "-p", help="Optional profile for crop/upscale"),
) -> None:
    """Apply screenshot preprocessing and save debug image."""
    prof = load_profile(profile) if profile else None
    output = run_preprocess(image_path, out, profile=prof)
    console.print(f"[green]Saved preprocessed image to {output}[/green]")


@app.command("ocr")
def ocr_cmd(
    image_path: Path = typer.Argument(..., help="Input screenshot path"),
    out: Path = typer.Option(..., "--out", "-o", help="Output raw_ocr.json path"),
    profile: str | None = typer.Option(None, "--profile", "-p", help="Optional profile for crop/upscale"),
) -> None:
    """Run OCR with word-level bounding boxes."""
    from fm26_screenshot_exporter.pipeline import run_ocr_pipeline

    prof = load_profile(profile) if profile else None
    tokens, _ = run_ocr_pipeline(image_path, profile=prof)
    output = save_ocr_json(tokens, out)
    console.print(f"[green]Saved {len(tokens)} OCR tokens to {output}[/green]")


@app.command("detect-table")
def detect_table_cmd(
    image_path: Path = typer.Argument(..., help="Input screenshot path"),
    out: Path = typer.Option(..., "--out", "-o", help="Output table_regions.json path"),
    profile: str | None = typer.Option(None, "--profile", "-p", help="Profile for manual crop/columns"),
) -> None:
    """Detect table, header, row, and column regions."""
    prof = load_profile(profile) if profile else None
    regions = run_detect_table(image_path, profile=prof)
    output = save_table_regions(regions, out)
    console.print(f"[green]Saved table regions to {output}[/green]")
    console.print(f"  table_bbox: {regions.table_bbox}")
    console.print(f"  row boundaries: {len(regions.row_y_boundaries)}")
    console.print(f"  column boundaries: {len(regions.column_x_boundaries)}")


@app.command("parse")
def parse_cmd(
    image_path: Path = typer.Argument(..., help="Input screenshot path"),
    profile: str = typer.Option(..., "--profile", "-p", help="Extraction profile name or path"),
    out: Path = typer.Option(..., "--out", "-o", help="Output file (.csv, .xlsx, .json, .parquet)"),
) -> None:
    """Full pipeline: preprocess → OCR → reconstruct → normalize → export."""
    result = parse_and_export(image_path, profile, out)
    console.print(
        f"[green]Parsed {len(result.rows)} rows → {out}[/green] "
        f"(profile: {result.profile_name})"
    )
    _echo_validation_warnings(result)


@app.command("parse-folder")
def parse_folder_cmd(
    input_dir: Path = typer.Argument(..., help="Folder with PNG/JPG screenshots"),
    profile: str = typer.Option(..., "--profile", "-p", help="Extraction profile"),
    out_dir: Path = typer.Option(..., "--out-dir", "-o", help="Output directory"),
    format: str = typer.Option("csv", "--format", "-f", help="Output format: csv, xlsx, json, parquet"),
) -> None:
    """Process all screenshots in a folder."""
    input_dir = Path(input_dir)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    images = sorted(
        p for p in input_dir.iterdir() if p.is_file() and is_image_file(p)
    )
    if not images:
        console.print(f"[yellow]No images found in {input_dir}[/yellow]")
        raise typer.Exit(1)

    suffix = f".{format.lstrip('.')}"
    for image in images:
        output = out_dir / f"{image.stem}{suffix}"
        result = parse_and_export(image, profile, output)
        console.print(f"[green]{image.name} → {output.name} ({len(result.rows)} rows)[/green]")
        _echo_validation_warnings(result)


@app.command("watch-folder")
def watch_folder_cmd(
    input_dir: Path = typer.Argument(..., help="Folder to watch for new screenshots"),
    profile: str = typer.Option(..., "--profile", "-p", help="Extraction profile"),
    out_dir: Path = typer.Option(..., "--out-dir", "-o", help="Output directory"),
    format: str = typer.Option("csv", "--format", "-f", help="Output format"),
) -> None:
    """Watch for new screenshots and auto-parse."""
    out_path = Path(out_dir)
    out_path.mkdir(parents=True, exist_ok=True)
    suffix = f".{format.lstrip('.')}"

    def on_image(path: Path) -> None:
        output = out_path / f"{path.stem}{suffix}"
        console.print(f"[cyan]Processing {path.name}...[/cyan]")
        try:
            result = parse_and_export(path, profile, output)
            console.print(f"[green]→ {output} ({len(result.rows)} rows)[/green]")
            _echo_validation_warnings(result)
        except Exception as exc:
            console.print(f"[red]Failed {path.name}: {exc}[/red]")

    console.print(f"Watching [bold]{input_dir}[/bold] — Ctrl+C to stop")
    watch_folder(input_dir, on_image)


@app.command("debug")
def debug_cmd(
    image_path: Path = typer.Argument(..., help="Input screenshot path"),
    profile: str = typer.Option(..., "--profile", "-p", help="Extraction profile"),
    out: Path = typer.Option(..., "--out", "-o", help="Output debug overlay image"),
) -> None:
    """Render debug overlay with columns, rows, and OCR boxes."""
    output = render_debug(image_path, profile, out)
    result = parse_screenshot(image_path, profile)
    console.print(f"[green]Saved debug overlay to {output}[/green]")
    console.print(f"  rows: {len(result.rows)}, tokens: {len(result.ocr_tokens)}")
    _echo_validation_warnings(result)


if __name__ == "__main__":
    app()
