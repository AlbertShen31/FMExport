"""Load and resolve YAML extraction profiles."""

from __future__ import annotations

from pathlib import Path

import yaml

from fm26_screenshot_exporter.config import Profile

_PACKAGE_ROOT = Path(__file__).resolve().parent.parent
_DEFAULT_PROFILES_DIR = _PACKAGE_ROOT / "profiles"


def profiles_dir() -> Path:
    return _DEFAULT_PROFILES_DIR


def list_profiles(directory: Path | None = None) -> list[str]:
    root = directory or profiles_dir()
    if not root.exists():
        return []
    return sorted(p.stem for p in root.glob("*.yaml"))


def load_profile(name_or_path: str, *, profiles_root: Path | None = None) -> Profile:
    """Load a profile by name (without extension) or explicit file path."""
    path = Path(name_or_path)
    if path.suffix in {".yaml", ".yml"} and path.exists():
        profile_path = path
    else:
        root = profiles_root or profiles_dir()
        candidates = [
            root / f"{name_or_path}.yaml",
            root / f"{name_or_path}.yml",
            Path(name_or_path),
        ]
        profile_path = next((c for c in candidates if c.exists()), None)
        if profile_path is None:
            available = ", ".join(list_profiles(root)) or "(none)"
            raise FileNotFoundError(
                f"Profile '{name_or_path}' not found. Available: {available}"
            )

    data = yaml.safe_load(profile_path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError(f"Invalid profile format in {profile_path}")

    profile = Profile.model_validate(data)
    if not profile.name:
        profile = profile.model_copy(update={"name": profile_path.stem})
    return profile
