"""macOS path discovery for Football Manager 2026 installations and exports."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class FM26Location:
    """A detected Football Manager-related folder on macOS."""

    path: Path
    kind: str
    confidence: str
    notes: str = ""
    children: list[Path] = field(default_factory=list)


def _expand(path: str | Path) -> Path:
    return Path(path).expanduser().resolve()


def _steam_fm_paths() -> list[FM26Location]:
    locations: list[FM26Location] = []
    steam_root = _expand("~/Library/Application Support/Steam")
    if not steam_root.exists():
        return locations

    for app_manifest in steam_root.glob("steamapps/appmanifest_*.acf"):
        try:
            text = app_manifest.read_text(encoding="utf-8", errors="replace")
        except OSError:
            continue
        if "football manager" not in text.lower() and "fm26" not in text.lower():
            continue

        installdir = None
        for line in text.splitlines():
            if '"installdir"' in line:
                installdir = line.split('"')[3]
                break
        if not installdir:
            continue

        game_dir = steam_root / "steamapps/common" / installdir
        if game_dir.exists():
            locations.append(
                FM26Location(
                    path=game_dir,
                    kind="steam_install",
                    confidence="high",
                    notes=f"Detected via {app_manifest.name}",
                )
            )

    return locations


def _documents_exports() -> list[FM26Location]:
    locations: list[FM26Location] = []
    docs = _expand("~/Documents")
    if not docs.exists():
        return locations

    candidates = [
        docs / "Sports Interactive" / "Football Manager 2026",
        docs / "Sports Interactive" / "Football Manager 26",
        docs / "Football Manager 2026",
        docs / "Football Manager 26",
    ]
    for path in candidates:
        if path.exists():
            html_files = list(path.rglob("*.html")) + list(path.rglob("*.htm"))
            locations.append(
                FM26Location(
                    path=path,
                    kind="documents",
                    confidence="medium",
                    notes="Sports Interactive documents folder",
                    children=html_files[:10],
                )
            )

    return locations


def _downloads_exports() -> list[FM26Location]:
    locations: list[FM26Location] = []
    downloads = _expand("~/Downloads")
    if not downloads.exists():
        return locations

    html_files = [
        p
        for p in downloads.glob("*.html")
        if any(k in p.name.lower() for k in ("player", "squad", "fm", "football"))
    ]
    if html_files:
        locations.append(
            FM26Location(
                path=downloads,
                kind="downloads",
                confidence="low",
                notes=f"Found {len(html_files)} likely FM HTML export(s)",
                children=html_files[:10],
            )
        )
    return locations


def _applications_search() -> list[FM26Location]:
    locations: list[FM26Location] = []
    apps = _expand("/Applications")
    if not apps.exists():
        return locations

    for app in apps.glob("Football Manager*.app"):
        locations.append(
            FM26Location(
                path=app,
                kind="application",
                confidence="high",
                notes="FM application bundle",
            )
        )
    return locations


def _container_saves() -> list[FM26Location]:
    """Check sandboxed container paths used by some Mac App Store builds."""
    locations: list[FM26Location] = []
    containers = _expand("~/Library/Containers")
    if not containers.exists():
        return locations

    for container in containers.glob("*Football*Manager*"):
        data = container / "Data" / "Documents"
        if data.exists():
            locations.append(
                FM26Location(
                    path=data,
                    kind="container",
                    confidence="medium",
                    notes=f"App container: {container.name}",
                )
            )
    return locations


def scan_fm26_locations() -> list[FM26Location]:
    """Scan common macOS locations for FM26 installs and export folders."""
    seen: set[Path] = set()
    results: list[FM26Location] = []

    scanners = (
        _applications_search,
        _steam_fm_paths,
        _documents_exports,
        _downloads_exports,
        _container_saves,
    )
    for scanner in scanners:
        for loc in scanner():
            resolved = loc.path.resolve()
            if resolved not in seen:
                seen.add(resolved)
                results.append(loc)

    return sorted(results, key=lambda loc: (loc.confidence != "high", loc.path))


def default_watch_folder() -> Path | None:
    """Return the best default folder to watch for FM HTML exports."""
    for loc in scan_fm26_locations():
        if loc.kind in {"documents", "downloads"}:
            return loc.path
    docs = _expand("~/Documents/Sports Interactive/Football Manager 2026")
    return docs if docs.parent.exists() else _expand("~/Downloads")
