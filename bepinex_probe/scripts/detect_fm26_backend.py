#!/usr/bin/env python3
"""Stage 1: Detect whether FM26 on macOS uses Unity Mono or IL2CPP."""

from __future__ import annotations

import argparse
import plistlib
import sys
from dataclasses import dataclass
from pathlib import Path


@dataclass
class BackendReport:
    app_path: Path
    backend: str  # "il2cpp", "mono", "unknown"
    executable: Path | None
    evidence: list[str]
    warnings: list[str]

    def print_report(self) -> None:
        print("=" * 60)
        print("FM26 Unity Backend Detection (Stage 1)")
        print("=" * 60)
        print(f"App bundle:     {self.app_path}")
        print(f"Backend:        {self.backend.upper()}")
        if self.executable:
            print(f"Executable:     {self.executable}")
        else:
            print("Executable:     (not found)")
        print()
        if self.evidence:
            print("Evidence:")
            for item in self.evidence:
                print(f"  - {item}")
        if self.warnings:
            print()
            print("Warnings:")
            for item in self.warnings:
                print(f"  ! {item}")
        print()
        if self.backend == "il2cpp":
            print(
                "Note: IL2CPP builds require BepInEx 6 IL2CPP builds on macOS. "
                "Compatibility is uncertain — treat as experimental."
            )
        elif self.backend == "unknown":
            print(
                "Note: Backend could not be determined. "
                "Do not install BepInEx until the Unity backend is confirmed."
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

    raise FileNotFoundError(
        f"Expected an .app bundle or install directory containing one: {path}"
    )


def _inspect_paths(app_path: Path) -> BackendReport:
    contents = app_path / "Contents"
    resources_data = contents / "Resources" / "Data"
    managed = resources_data / "Managed"
    il2cpp_data = resources_data / "il2cpp_data"
    frameworks = contents / "Frameworks"
    macos = contents / "MacOS"

    evidence: list[str] = []
    warnings: list[str] = []

    executable: Path | None = None
    info_plist = contents / "Info.plist"
    if info_plist.exists():
        try:
            with info_plist.open("rb") as fh:
                info = plistlib.load(fh)
            exe_name = info.get("CFBundleExecutable")
            if exe_name:
                candidate = macos / exe_name
                if candidate.exists():
                    executable = candidate
                    evidence.append(f"Info.plist executable: {exe_name}")
        except (plistlib.InvalidFileException, OSError) as exc:
            warnings.append(f"Could not read Info.plist: {exc}")

    if executable is None and macos.is_dir():
        binaries = [
            p
            for p in macos.iterdir()
            if p.is_file() and not p.name.startswith(".")
        ]
        if len(binaries) == 1:
            executable = binaries[0]
            evidence.append(f"Single binary in MacOS: {executable.name}")
        elif binaries:
            preferred = [p for p in binaries if "football" in p.name.lower() or "fm" in p.name.lower()]
            executable = preferred[0] if preferred else binaries[0]
            evidence.append(f"Selected MacOS binary: {executable.name}")

    il2cpp_score = 0
    mono_score = 0

    if il2cpp_data.is_dir():
        il2cpp_score += 2
        evidence.append(f"Found il2cpp_data: {il2cpp_data}")
        metadata = il2cpp_data / "Metadata" / "global-metadata.dat"
        if metadata.exists():
            il2cpp_score += 2
            evidence.append(f"Found global-metadata.dat ({metadata.stat().st_size:,} bytes)")

    if frameworks.is_dir():
        for pattern in ("GameAssembly", "libil2cpp", "UnityPlayer"):
            matches = list(frameworks.glob(f"*{pattern}*"))
            for match in matches:
                il2cpp_score += 1
                evidence.append(f"Framework: {match.name}")

    if managed.is_dir():
        dlls = list(managed.glob("*.dll"))
        if dlls:
            mono_score += 2
            evidence.append(f"Managed DLLs: {len(dlls)} in {managed}")
            game_dlls = [d for d in dlls if "assembly" in d.name.lower() or "game" in d.name.lower()]
            if game_dlls:
                mono_score += 1
                evidence.append(f"Likely game assemblies: {', '.join(d.name for d in game_dlls[:5])}")
    elif (resources_data / "Managed").exists():
        evidence.append("Managed path exists but contains no DLLs (likely IL2CPP)")

    scripting_backend = resources_data / "ScriptingAssemblies.json"
    if scripting_backend.exists():
        evidence.append(f"Found ScriptingAssemblies.json (common in IL2CPP builds)")

    boot_config = resources_data / "boot.config"
    if boot_config.exists():
        try:
            text = boot_config.read_text(encoding="utf-8", errors="replace")
            if "il2cpp" in text.lower():
                il2cpp_score += 1
                evidence.append("boot.config mentions il2cpp")
        except OSError:
            pass

    if il2cpp_score > mono_score:
        backend = "il2cpp"
    elif mono_score > il2cpp_score:
        backend = "mono"
    else:
        backend = "unknown"
        warnings.append(
            "Could not confidently determine Unity backend. "
            "Inspect Contents/Resources/Data manually before installing BepInEx."
        )

    if backend == "mono":
        warnings.append(
            "Mono backend detected. BepInEx 5.x Mono may apply instead of IL2CPP builds. "
            "Verify against the BepInEx FM26 macOS documentation before proceeding."
        )

    return BackendReport(
        app_path=app_path,
        backend=backend,
        executable=executable,
        evidence=evidence,
        warnings=warnings,
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Detect FM26 Unity backend (Mono vs IL2CPP) on macOS."
    )
    parser.add_argument(
        "--app-path",
        required=True,
        help="Path to Football Manager 26.app or its parent install directory.",
    )
    args = parser.parse_args(argv)

    try:
        app_path = _resolve_app_path(args.app_path)
        report = _inspect_paths(app_path)
        report.print_report()
        return 0 if report.backend != "unknown" else 1
    except FileNotFoundError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
