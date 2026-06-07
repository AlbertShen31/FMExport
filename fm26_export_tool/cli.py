"""CLI entry point for fm26-export-tool."""

from __future__ import annotations

import json
import sys
from pathlib import Path

import click

from fm26_export_tool import __version__
from fm26_export_tool.exporters import (
    export_csv,
    export_json,
    export_parquet,
    export_sqlite,
    export_xlsx,
)
from fm26_export_tool.parser import parse_html
from fm26_export_tool.paths import default_watch_folder, scan_fm26_locations
from fm26_export_tool.watcher import watch_folder


def _load_parsed(input_path: Path, table_index: int | None, no_normalize: bool):
    return parse_html(input_path, table_index=table_index, normalize=not no_normalize)


@click.group()
@click.version_option(__version__, prog_name="fm26-export")
def main() -> None:
    """macOS-first Football Manager 2026 stats export tool."""


@main.command("scan")
def scan_cmd() -> None:
    """Detect likely FM26 folders on macOS."""
    locations = scan_fm26_locations()
    if not locations:
        click.echo("No Football Manager 2026 locations detected.")
        click.echo("Tip: export a player list to HTML, then run scan again.")
        return

    click.echo(f"Found {len(locations)} location(s):\n")
    for loc in locations:
        click.echo(f"  [{loc.confidence}] {loc.kind}: {loc.path}")
        if loc.notes:
            click.echo(f"           {loc.notes}")
        for child in loc.children[:5]:
            click.echo(f"           - {child.name}")
        if len(loc.children) > 5:
            click.echo(f"           ... and {len(loc.children) - 5} more")


@main.command("parse-html")
@click.argument("input_path", type=click.Path(exists=True, path_type=Path))
@click.option("--table-index", type=int, default=None, help="Force a specific table index.")
@click.option("--no-normalize", is_flag=True, help="Skip value normalization.")
@click.option("--output", "-o", type=click.Path(path_type=Path), help="Write JSON summary.")
def parse_html_cmd(
    input_path: Path,
    table_index: int | None,
    no_normalize: bool,
    output: Path | None,
) -> None:
    """Parse one FM-exported HTML file."""
    parsed = _load_parsed(input_path, table_index, no_normalize)
    summary = {
        "source": str(parsed.source_path),
        "table_index": parsed.table_index,
        "rows": len(parsed.dataframe),
        "columns": parsed.column_map,
        "metadata": parsed.metadata,
        "preview": parsed.dataframe.head(5).to_dict(orient="records"),
    }

    if output:
        output.write_text(json.dumps(summary, indent=2, default=str), encoding="utf-8")
        click.echo(f"Wrote summary to {output}")
    else:
        click.echo(json.dumps(summary, indent=2, default=str))


@main.command("export-csv")
@click.argument("input_path", type=click.Path(exists=True, path_type=Path))
@click.option("--output", "-o", type=click.Path(path_type=Path))
@click.option("--table-index", type=int, default=None)
@click.option("--no-normalize", is_flag=True)
@click.option("--include-original-columns", is_flag=True)
def export_csv_cmd(
    input_path: Path,
    output: Path | None,
    table_index: int | None,
    no_normalize: bool,
    include_original_columns: bool,
) -> None:
    """Convert a parsed FM HTML table to CSV (UTF-8)."""
    parsed = _load_parsed(input_path, table_index, no_normalize)
    out = export_csv(parsed, output, include_original_columns=include_original_columns)
    click.echo(f"Wrote {out}")


@main.command("export-xlsx")
@click.argument("input_path", type=click.Path(exists=True, path_type=Path))
@click.option("--output", "-o", type=click.Path(path_type=Path))
@click.option("--table-index", type=int, default=None)
@click.option("--no-normalize", is_flag=True)
@click.option("--include-original-columns", is_flag=True)
def export_xlsx_cmd(
    input_path: Path,
    output: Path | None,
    table_index: int | None,
    no_normalize: bool,
    include_original_columns: bool,
) -> None:
    """Convert a parsed FM HTML table to Excel."""
    parsed = _load_parsed(input_path, table_index, no_normalize)
    out = export_xlsx(parsed, output, include_original_columns=include_original_columns)
    click.echo(f"Wrote {out}")


@main.command("export-json")
@click.argument("input_path", type=click.Path(exists=True, path_type=Path))
@click.option("--output", "-o", type=click.Path(path_type=Path))
@click.option("--table-index", type=int, default=None)
@click.option("--no-normalize", is_flag=True)
@click.option("--no-metadata", is_flag=True)
def export_json_cmd(
    input_path: Path,
    output: Path | None,
    table_index: int | None,
    no_normalize: bool,
    no_metadata: bool,
) -> None:
    """Convert a parsed FM HTML table to JSON records."""
    parsed = _load_parsed(input_path, table_index, no_normalize)
    out = export_json(parsed, output, include_metadata=not no_metadata)
    click.echo(f"Wrote {out}")


@main.command("export-parquet")
@click.argument("input_path", type=click.Path(exists=True, path_type=Path))
@click.option("--output", "-o", type=click.Path(path_type=Path))
@click.option("--table-index", type=int, default=None)
@click.option("--no-normalize", is_flag=True)
def export_parquet_cmd(
    input_path: Path,
    output: Path | None,
    table_index: int | None,
    no_normalize: bool,
) -> None:
    """Convert a parsed FM HTML table to Parquet (requires pyarrow)."""
    parsed = _load_parsed(input_path, table_index, no_normalize)
    try:
        out = export_parquet(parsed, output)
    except ImportError as exc:
        click.echo(str(exc), err=True)
        sys.exit(1)
    click.echo(f"Wrote {out}")


@main.command("export-sqlite")
@click.argument("input_path", type=click.Path(exists=True, path_type=Path))
@click.option("--output", "-o", type=click.Path(path_type=Path))
@click.option("--table-index", type=int, default=None)
@click.option("--no-normalize", is_flag=True)
@click.option("--table-name", default="players")
def export_sqlite_cmd(
    input_path: Path,
    output: Path | None,
    table_index: int | None,
    no_normalize: bool,
    table_name: str,
) -> None:
    """Convert a parsed FM HTML table to SQLite."""
    parsed = _load_parsed(input_path, table_index, no_normalize)
    out = export_sqlite(parsed, output, table_name=table_name)
    click.echo(f"Wrote {out}")


@main.command("inspect-columns")
@click.argument("input_path", type=click.Path(exists=True, path_type=Path))
@click.option("--table-index", type=int, default=None)
@click.option("--no-normalize", is_flag=True)
@click.option("--sample-rows", default=3, show_default=True)
def inspect_columns_cmd(
    input_path: Path,
    table_index: int | None,
    no_normalize: bool,
    sample_rows: int,
) -> None:
    """Print detected columns and sample values."""
    parsed = _load_parsed(input_path, table_index, no_normalize)
    click.echo(f"Source: {parsed.source_path}")
    click.echo(f"Table index: {parsed.table_index}  |  Rows: {len(parsed.dataframe)}\n")

    for norm, orig in parsed.column_map.items():
        series = parsed.dataframe[norm]
        samples = [v for v in series.head(sample_rows).tolist() if v is not None]
        sample_text = ", ".join(str(s) for s in samples) if samples else "(empty)"
        click.echo(f"  {norm}")
        click.echo(f"    original: {orig}")
        click.echo(f"    samples:  {sample_text}")
        click.echo("")


@main.command("watch-folder")
@click.argument("folder", type=click.Path(path_type=Path), required=False)
@click.option(
    "--formats",
    default="csv,xlsx,json",
    show_default=True,
    help="Comma-separated export formats.",
)
@click.option("--output-dir", "-o", type=click.Path(path_type=Path))
def watch_folder_cmd(folder: Path | None, formats: str, output_dir: Path | None) -> None:
    """Watch a folder for new FM exported HTML files and auto-convert them."""
    watch_path = folder or default_watch_folder()
    if watch_path is None:
        click.echo("No watch folder specified and no default found.", err=True)
        sys.exit(1)

    fmt_list = [f.strip() for f in formats.split(",") if f.strip()]
    watch_folder(watch_path, formats=fmt_list, output_dir=output_dir)


if __name__ == "__main__":
    main()
