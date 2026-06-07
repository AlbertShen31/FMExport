#!/usr/bin/env python3
"""Stage 2: Detect whether BepInEx is installed and logging for FM26 on macOS."""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass, field
from pathlib import Path

LOG_MARKERS = (
    "Chainloader startup complete",
    "Loading plugin",
    "FM26ExportProbe",
    "error",
    "exception",
)


@dataclass
class BepInExReport:
    search_roots: list[Path]
    installed: bool
    found_paths: dict[str, Path | None] = field(default_factory=dict)
    log_tail: list[str] = field(default_factory=list)
    marker_hits: dict[str, list[str]] = field(default_factory=dict)
    notes: list[str] = field(default_factory=list)

    def print_report(self) -> None:
        print("=" * 60)
        print("BepInEx Install Detection (Stage 2)")
        print("=" * 60)
        print(f"Installed:      {'YES' if self.installed else 'NO'}")
        print()
        print("Search roots:")
        for root in self.search_roots:
            print(f"  - {root}")
        print()
        print("Expected paths:")
        for name, path in self.found_paths.items():
            status = str(path) if path else "(missing)"
            print(f"  {name}: {status}")
        print()

        if self.notes:
            print("Notes:")
            for note in self.notes:
                print(f"  - {note}")
            print()

        if self.marker_hits:
            print("Log markers:")
            for marker, lines in self.marker_hits.items():
                print(f"  [{marker}] {len(lines)} hit(s)")
                for line in lines[:3]:
                    print(f"    {line.strip()}")
            print()

        if self.log_tail:
            print(f"Last {len(self.log_tail)} lines of LogOutput.log:")
            print("-" * 60)
            for line in self.log_tail:
                print(line.rstrip())
            print("-" * 60)
        else:
            print("LogOutput.log: not found or empty.")
            print()
            print(
                "Stop condition: If LogOutput.log is never created after launching FM26, "
                "BepInEx did not boot. Fall back to manual HTML export or OCR."
            )


def _resolve_app_path(raw: str) -> Path:
    path = Path(raw).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"App path does not exist: {path}")
    if path.suffix == ".app":
        return path
    if path.is_dir() and (path / "Contents").is_dir():
        return path

    preferred_names = ("fm.app", "Football Manager 26.app", "Football Manager 2026.app")
    for name in preferred_names:
        candidate = path / name
        if candidate.exists():
            return candidate.resolve()

    app_candidates = sorted(path.glob("*.app"))
    if len(app_candidates) == 1:
        return app_candidates[0].resolve()
    if app_candidates:
        for candidate in app_candidates:
            if "football" in candidate.name.lower() or candidate.name.lower() == "fm.app":
                return candidate.resolve()
        return app_candidates[0].resolve()

    raise FileNotFoundError(f"Could not resolve .app bundle from: {path}")


def _search_roots(app_path: Path) -> list[Path]:
    contents = app_path / "Contents"
    roots = [
        contents / "MacOS",
        app_path.parent,
        app_path,
    ]
    unique: list[Path] = []
    seen: set[Path] = set()
    for root in roots:
        resolved = root.resolve()
        if resolved not in seen:
            seen.add(resolved)
            unique.append(resolved)
    return unique


def _find_first(roots: list[Path], relative: str) -> Path | None:
    for root in roots:
        candidate = root / relative
        if candidate.exists():
            return candidate
    return None


def _read_log_tail(log_path: Path, lines: int = 100) -> list[str]:
    try:
        content = log_path.read_text(encoding="utf-8", errors="replace").splitlines()
    except OSError:
        return []
    return content[-lines:]


def _scan_markers(lines: list[str]) -> dict[str, list[str]]:
    hits: dict[str, list[str]] = {marker: [] for marker in LOG_MARKERS}
    for line in lines:
        lower = line.lower()
        for marker in LOG_MARKERS:
            if marker.lower() in lower:
                hits[marker].append(line)
    return {k: v for k, v in hits.items() if v}


def inspect_bepinex(app_path: Path) -> BepInExReport:
    roots = _search_roots(app_path)
    targets = {
        "BepInEx": "BepInEx",
        "BepInEx/plugins": "BepInEx/plugins",
        "BepInEx/LogOutput.log": "BepInEx/LogOutput.log",
        "run_bepinex.sh": "run_bepinex.sh",
        "doorstop_config.ini": "doorstop_config.ini",
    }

    found: dict[str, Path | None] = {}
    for label, rel in targets.items():
        found[label] = _find_first(roots, rel)

    log_path = found.get("BepInEx/LogOutput.log")
    log_tail = _read_log_tail(log_path) if log_path else []
    marker_hits = _scan_markers(log_tail)

    installed = found["BepInEx"] is not None
    notes: list[str] = []

    if installed and not log_path:
        notes.append(
            "BepInEx folder exists but LogOutput.log is missing. "
            "Launch FM26 once via BepInEx and re-run this script."
        )
    if log_path and not marker_hits.get("Chainloader startup complete"):
        notes.append(
            "Log exists but 'Chainloader startup complete' not found in last 100 lines. "
            "BepInEx may have failed during startup."
        )
    if marker_hits.get("FM26ExportProbe"):
        notes.append("FM26ExportProbe plugin messages found in log — plugin may have loaded.")
    if marker_hits.get("error") or marker_hits.get("exception"):
        notes.append("Errors/exceptions found in log — review before continuing.")

    return BepInExReport(
        search_roots=roots,
        installed=installed,
        found_paths=found,
        log_tail=log_tail,
        marker_hits=marker_hits,
        notes=notes,
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Detect BepInEx installation and log status for FM26 on macOS."
    )
    parser.add_argument(
        "--app-path",
        required=True,
        help="Path to Football Manager 26.app or install directory.",
    )
    args = parser.parse_args(argv)

    try:
        app_path = _resolve_app_path(args.app_path)
        report = inspect_bepinex(app_path)
        report.print_report()
        return 0 if report.installed else 1
    except FileNotFoundError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
