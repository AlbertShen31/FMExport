#!/usr/bin/env python3
"""Build and package the FM26ExportProbe BepInEx plugin DLL."""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
PROBE_ROOT = SCRIPT_DIR.parent
SRC_ROOT = PROBE_ROOT / "src" / "FM26ExportProbe"
DIST_ROOT = PROBE_ROOT / "dist" / "FM26ExportProbe"


def _find_bepinex_libs(explicit: Path | None) -> Path | None:
    if explicit:
        path = explicit.expanduser().resolve()
        if path.is_dir():
            return path
        raise FileNotFoundError(f"BepInEx libs path not found: {path}")

    env = os.environ.get("BEPINEX_LIBS")
    if env:
        path = Path(env).expanduser().resolve()
        if path.is_dir():
            return path

    # Common locations after manual BepInEx install inside FM26.app
    steam_root = (
        Path.home()
        / "Library/Application Support/Steam/steamapps/common/Football Manager 26"
    )
    candidates = [
        steam_root / "fm.app/Contents/MacOS/BepInEx/core",
        steam_root / "Football Manager 26.app/Contents/MacOS/BepInEx/core",
        Path("/Applications/fm.app/Contents/MacOS/BepInEx/core"),
        Path("/Applications/Football Manager 26.app/Contents/MacOS/BepInEx/core"),
    ]
    for candidate in candidates:
        if candidate.is_dir() and list(candidate.glob("BepInEx*.dll")):
            return candidate
    return None


def build(configuration: str, bepinex_libs: Path | None) -> Path:
    project = SRC_ROOT / "FM26ExportProbe.csproj"
    if not project.exists():
        raise FileNotFoundError(f"Project file not found: {project}")

    env = os.environ.copy()
    if bepinex_libs:
        env["BEPINEX_LIBS"] = str(bepinex_libs)

    print("$", "dotnet restore", str(project))
    subprocess.run(["dotnet", "restore", str(project)], cwd=SRC_ROOT, check=True, env=env)

    cmd = ["dotnet", "build", str(project), "-c", configuration, "--no-restore"]
    print("$", " ".join(cmd))
    subprocess.run(cmd, cwd=SRC_ROOT, check=True, env=env)

    output_dir = SRC_ROOT / "bin" / configuration / "net6.0"
    dll = output_dir / "FM26ExportProbe.dll"
    if not dll.exists():
        raise FileNotFoundError(f"Build succeeded but DLL not found: {dll}")
    return dll


def package(dll: Path, target: Path) -> None:
    target.mkdir(parents=True, exist_ok=True)
    shutil.copy2(dll, target / dll.name)
    readme = target / "INSTALL.txt"
    readme.write_text(
        "\n".join(
            [
                "FM26 Export Probe — install instructions",
                "",
                "1. Confirm BepInEx Stage 2 detection passes.",
                "2. Copy FM26ExportProbe.dll into:",
                "   <FM26.app>/Contents/MacOS/BepInEx/plugins/FM26ExportProbe/",
                "3. Launch FM26 once.",
                "4. Check ~/Documents/FM26Exports/probe_loaded.txt",
                "5. Press F8 in-game and inspect fm26_probe_diagnostic.txt",
                "",
                "Do not overwrite game files without a backup.",
            ]
        ),
        encoding="utf-8",
    )
    print(f"Packaged plugin to: {target}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Build FM26ExportProbe plugin DLL.")
    parser.add_argument(
        "--configuration",
        default="Release",
        help="dotnet build configuration (default: Release).",
    )
    parser.add_argument(
        "--bepinex-libs",
        type=Path,
        help="Path to BepInEx core DLLs (or set BEPINEX_LIBS env var).",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=DIST_ROOT,
        help="Package output directory.",
    )
    args = parser.parse_args(argv)

    try:
        libs = _find_bepinex_libs(args.bepinex_libs)
        if libs:
            print(f"Using BepInEx libs: {libs}")
        else:
            print(
                "Warning: BepInEx core DLLs not found. Build may fail unless "
                "NuGet-resolved references are sufficient.",
            )
            print("Set --bepinex-libs or BEPINEX_LIBS to your FM26 BepInEx/core folder.")

        dll = build(args.configuration, libs)
        package(dll, args.output.expanduser().resolve())
        return 0
    except FileNotFoundError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2
    except subprocess.CalledProcessError as exc:
        print(f"Build failed with exit code {exc.returncode}", file=sys.stderr)
        return exc.returncode or 3


if __name__ == "__main__":
    raise SystemExit(main())
