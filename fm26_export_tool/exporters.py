"""Export parsed FM tables to CSV, Excel, JSON, Parquet, and SQLite."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl.utils import get_column_letter

from fm26_export_tool.parser import ParsedExport


def _resolve_output_path(output: str | Path | None, source: Path, suffix: str) -> Path:
    if output is not None:
        return Path(output)
    return source.with_suffix(suffix)


def _prepare_export_df(parsed: ParsedExport, include_original_columns: bool) -> pd.DataFrame:
    df = parsed.dataframe.copy()
    if include_original_columns:
        for norm, orig in parsed.column_map.items():
            if norm in df.columns:
                df[f"_original_{norm}"] = orig
    return df


def export_csv(
    parsed: ParsedExport,
    output: str | Path | None = None,
    *,
    include_original_columns: bool = False,
) -> Path:
    out = _resolve_output_path(output, parsed.source_path, ".csv")
    df = _prepare_export_df(parsed, include_original_columns)
    df.to_csv(out, index=False, encoding="utf-8")
    return out


def export_xlsx(
    parsed: ParsedExport,
    output: str | Path | None = None,
    *,
    include_original_columns: bool = False,
) -> Path:
    out = _resolve_output_path(output, parsed.source_path, ".xlsx")
    df = _prepare_export_df(parsed, include_original_columns)

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        sheet_name = "Players"
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.sheets[sheet_name]

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for idx, column in enumerate(df.columns, start=1):
            letter = get_column_letter(idx)
            lengths = [len(str(column))]
            lengths.extend(
                len(str(v)) if v is not None else 0
                for v in df[column].head(50)
            )
            ws.column_dimensions[letter].width = min(max(lengths) + 2, 50)

    return out


def export_json(
    parsed: ParsedExport,
    output: str | Path | None = None,
    *,
    include_metadata: bool = True,
    indent: int = 2,
) -> Path:
    out = _resolve_output_path(output, parsed.source_path, ".json")
    payload: dict[str, Any] = {
        "source": str(parsed.source_path),
        "columns": parsed.column_map,
        "records": parsed.to_records(),
    }
    if include_metadata:
        payload["metadata"] = parsed.metadata

    out.write_text(json.dumps(payload, indent=indent, default=str), encoding="utf-8")
    return out


def export_parquet(
    parsed: ParsedExport,
    output: str | Path | None = None,
) -> Path:
    try:
        import pyarrow  # noqa: F401
    except ImportError as exc:
        raise ImportError(
            "pyarrow is required for Parquet export. "
            "Install with: pip install 'fm26-export-tool[parquet]'"
        ) from exc

    out = _resolve_output_path(output, parsed.source_path, ".parquet")
    parsed.dataframe.to_parquet(out, index=False)
    return out


def export_sqlite(
    parsed: ParsedExport,
    output: str | Path | None = None,
    *,
    table_name: str = "players",
) -> Path:
    out = _resolve_output_path(output, parsed.source_path, ".sqlite")
    if out.exists():
        out.unlink()
    parsed.dataframe.to_sql(table_name, f"sqlite:///{out}", index=False, if_exists="replace")
    return out


def export_all(
    parsed: ParsedExport,
    output_dir: str | Path | None = None,
    *,
    formats: list[str] | None = None,
) -> dict[str, Path]:
    """Export parsed data to multiple formats at once."""
    base_dir = Path(output_dir) if output_dir else parsed.source_path.parent
    stem = parsed.source_path.stem
    selected = formats or ["csv", "xlsx", "json"]

    results: dict[str, Path] = {}
    for fmt in selected:
        if fmt == "csv":
            results["csv"] = export_csv(parsed, base_dir / f"{stem}.csv")
        elif fmt == "xlsx":
            results["xlsx"] = export_xlsx(parsed, base_dir / f"{stem}.xlsx")
        elif fmt == "json":
            results["json"] = export_json(parsed, base_dir / f"{stem}.json")
        elif fmt == "parquet":
            results["parquet"] = export_parquet(parsed, base_dir / f"{stem}.parquet")
        elif fmt == "sqlite":
            results["sqlite"] = export_sqlite(parsed, base_dir / f"{stem}.sqlite")
        else:
            raise ValueError(f"Unknown export format: {fmt}")

    return results
