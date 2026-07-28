"""Versionnage de fichiers par nom logique (ex. par client) : cherche la version
la plus haute déjà présente dans un dossier et incrémente automatiquement
(mineure par défaut, ex. v2.0 -> v2.1). Partagé entre la génération de
catalogues (pptx_generator.py, generate_catalogs.py) et tout autre outil ayant
besoin de la même convention AAAA-MM-JJ_Catalogue_[Label]_v[X.Y].<ext>.
"""

from __future__ import annotations

import re
from pathlib import Path

ILLEGAL_CHARS = re.compile(r'[<>:"/\\|?*]')


def sanitize_name(name: str) -> str:
    name = ILLEGAL_CHARS.sub("_", name).strip()
    return re.sub(r"\s+", " ", name)


def find_latest_version(directory: Path, ext: str, name_filter: str | None = None) -> tuple[int, int] | None:
    """Cherche récursivement dans directory la version (major, minor) la plus haute
    pour des fichiers finissant par _v<major>.<minor>.<ext>. Si name_filter est
    fourni, ne considère que les fichiers dont le nom le contient (ex. le code client)."""
    if not directory.exists():
        return None
    version_re = re.compile(rf"_v(?P<major>\d+)\.(?P<minor>\d+)\.{re.escape(ext)}$", re.IGNORECASE)
    best: tuple[int, int] | None = None
    for file_path in directory.rglob(f"*.{ext}"):
        if name_filter and name_filter not in file_path.name:
            continue
        match = version_re.search(file_path.name)
        if not match:
            continue
        version = (int(match.group("major")), int(match.group("minor")))
        if best is None or version > best:
            best = version
    return best


def next_version(latest: tuple[int, int] | None, bump_major: bool = False) -> tuple[int, int]:
    if latest is None:
        return (1, 0)
    major, minor = latest
    if bump_major:
        return (major + 1, 0)
    return (major, minor + 1)


def build_versioned_filename(date_str: str, label: str, version: tuple[int, int], ext: str) -> str:
    major, minor = version
    return f"{date_str}_Catalogue_{sanitize_name(label)}_v{major}.{minor}.{ext}"


def next_available_path(
    directory: Path, date_str: str, label: str, ext: str, bump_major: bool = False
) -> tuple[Path, tuple[int, int]]:
    """Calcule le prochain chemin de fichier versionné libre (sans collision) pour ce label."""
    latest = find_latest_version(directory, ext, name_filter=sanitize_name(label))
    major, minor = next_version(latest, bump_major)
    for _attempt in range(1000):
        version = (major, minor)
        path = directory / build_versioned_filename(date_str, label, version, ext)
        if not path.exists():
            return path, version
        major, minor = next_version((major, minor), bump_major=False)
    raise RuntimeError(f"Impossible de trouver un nom de fichier libre pour '{label}' après 1000 tentatives.")