#!/usr/bin/env python3
"""Stage 3 helper: verify or scaffold the FM26ExportProbe plugin source tree."""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
PROBE_ROOT = SCRIPT_DIR.parent
SRC_ROOT = PROBE_ROOT / "src" / "FM26ExportProbe"

REQUIRED_FILES = (
    "FM26ExportProbe.csproj",
    "Plugin.cs",
    "Exporter.cs",
    "UiScanner.cs",
    "CsvWriter.cs",
)


def verify_skeleton() -> list[str]:
    missing = [name for name in REQUIRED_FILES if not (SRC_ROOT / name).exists()]
    return missing


def print_status() -> int:
    print("=" * 60)
    print("FM26ExportProbe Skeleton Check (Stage 3)")
    print("=" * 60)
    print(f"Source root: {SRC_ROOT}")
    print()
    missing = verify_skeleton()
    if missing:
        print("Missing files:")
        for name in missing:
            print(f"  - {name}")
        print()
        print("The repository should already include these files.")
        print("If missing, restore from version control or re-clone the repo.")
        return 1

    print("All required plugin source files are present:")
    for name in REQUIRED_FILES:
        path = SRC_ROOT / name
        print(f"  [ok] {name} ({path.stat().st_size:,} bytes)")
    print()
    print("Next steps:")
    print("  1. Confirm BepInEx Stage 2 detection passes.")
    print("  2. Run: python3 scripts/package_plugin.py")
    print("  3. Copy dist/FM26ExportProbe.dll into BepInEx/plugins/FM26ExportProbe/")
    return 0


def copy_to_staging(staging: Path) -> None:
    if staging.exists():
        shutil.rmtree(staging)
    shutil.copytree(SRC_ROOT, staging)
    print(f"Copied plugin sources to staging: {staging}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Verify FM26ExportProbe plugin skeleton exists."
    )
    parser.add_argument(
        "--copy-to",
        type=Path,
        help="Optional path to copy the skeleton for external editing.",
    )
    args = parser.parse_args(argv)

    if args.copy_to:
        copy_to_staging(args.copy_to.expanduser().resolve())

    return print_status()


if __name__ == "__main__":
    raise SystemExit(main())
