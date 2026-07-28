#!/usr/bin/env python
"""Renomme et archive des fichiers de catalogues téléchargés selon la convention
AAAA-MM-JJ_Catalogue_[NomGroupe]_v[X.Y].pdf.

Cherche dans un dossier source les fichiers nommés selon l'ancien format
catalogue-<nom-de-groupe>-<numero-JCP>.<ext>, et les copie (ou déplace) vers :

    <dossier-archive>/<NomGroupe>/<Année>/AAAA-MM-JJ_Catalogue_<NomGroupe>_v<X.Y>.<ext>

Le numéro de version est déterminé automatiquement par groupe : le script
cherche la version la plus haute déjà archivée pour ce groupe (toutes années
confondues) et incrémente la version mineure (v2.0 -> v2.1). Si aucune
version n'existe encore, démarre à v1.0.

Usage :
    python archive_catalogues.py --dossier-source "C:\\Users\\<vous>\\Downloads" --dossier-archive "C:\\...\\Catalogues"
"""

from __future__ import annotations

import argparse
import csv
import logging
import re
import shutil
import sys
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

logger = logging.getLogger("archive_catalogues")

SOURCE_PATTERN = re.compile(r"^catalogue-(?P<groupe>.+)-(?P<jcp>[A-Za-z0-9]+)$", re.IGNORECASE)
VERSION_PATTERN_TEMPLATE = r"_v(?P<major>\d+)\.(?P<minor>\d+)\.{ext}$"
ILLEGAL_CHARS = re.compile(r'[<>:"/\\|?*]')

HISTORY_CSV_NAME = "historique_versions.csv"
HISTORY_FIELDS = ["date", "version", "nom_fichier", "fichier_source_original"]


@dataclass
class ArchiveResult:
    source: Path
    status: str  # "ok" | "ignore" | "erreur"
    detail: str
    target: Path | None = None
    version: str | None = None


def sanitize_group_name(name: str) -> str:
    name = ILLEGAL_CHARS.sub("_", name).strip()
    return re.sub(r"\s+", " ", name)


def parse_source_filename(path: Path) -> tuple[str, str] | None:
    """Retourne (nom_groupe, numero_jcp) si le fichier correspond au pattern source, sinon None."""
    match = SOURCE_PATTERN.match(path.stem)
    if not match:
        return None
    return match.group("groupe"), match.group("jcp")


def find_latest_version(group_dir: Path, ext: str) -> tuple[int, int] | None:
    """Cherche récursivement dans group_dir la version (major, minor) la plus haute."""
    if not group_dir.exists():
        return None
    version_re = re.compile(VERSION_PATTERN_TEMPLATE.format(ext=re.escape(ext)), re.IGNORECASE)
    best: tuple[int, int] | None = None
    for file_path in group_dir.rglob(f"*.{ext}"):
        match = version_re.search(file_path.name)
        if not match:
            continue
        version = (int(match.group("major")), int(match.group("minor")))
        if best is None or version > best:
            best = version
    return best


def next_version(latest: tuple[int, int] | None, bump_major: bool) -> tuple[int, int]:
    if latest is None:
        return (1, 0)
    major, minor = latest
    if bump_major:
        return (major + 1, 0)
    return (major, minor + 1)


def resolve_date_str(source_path: Path, date_source: str) -> str:
    if date_source == "mtime":
        return datetime.fromtimestamp(source_path.stat().st_mtime).strftime("%Y-%m-%d")
    return date.today().strftime("%Y-%m-%d")


def append_history(group_dir: Path, date_str: str, version_str: str, nom_fichier: str, source_original: str) -> None:
    history_path = group_dir / HISTORY_CSV_NAME
    is_new = not history_path.exists()
    with open(history_path, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=HISTORY_FIELDS)
        if is_new:
            writer.writeheader()
        writer.writerow({
            "date": date_str,
            "version": version_str,
            "nom_fichier": nom_fichier,
            "fichier_source_original": source_original,
        })


def process_file(
    source_path: Path,
    archive_root: Path,
    ext: str,
    date_source: str,
    bump_major: bool,
    move: bool,
) -> ArchiveResult:
    parsed = parse_source_filename(source_path)
    if parsed is None:
        return ArchiveResult(source_path, "ignore", "Ne correspond pas au format catalogue-<groupe>-<jcp>.")

    groupe_brut, _jcp = parsed
    groupe = sanitize_group_name(groupe_brut)
    if not groupe:
        return ArchiveResult(source_path, "erreur", "Nom de groupe vide après extraction/nettoyage.")

    date_str = resolve_date_str(source_path, date_source)
    year = date_str[:4]

    group_dir = archive_root / groupe
    target_dir = group_dir / year

    latest = find_latest_version(group_dir, ext)
    major, minor = next_version(latest, bump_major)

    for _attempt in range(1000):
        version_str = f"{major}.{minor}"
        target_name = f"{date_str}_Catalogue_{groupe}_v{version_str}.{ext}"
        target_path = target_dir / target_name
        if not target_path.exists():
            break
        logger.warning("Collision de nom, incrémentation de version : %s existe déjà.", target_path.name)
        major, minor = next_version((major, minor), bump_major=False)
    else:
        return ArchiveResult(source_path, "erreur", "Impossible de trouver un nom de fichier libre après 1000 tentatives.")

    try:
        target_dir.mkdir(parents=True, exist_ok=True)
        if move:
            shutil.move(str(source_path), str(target_path))
        else:
            shutil.copy2(str(source_path), str(target_path))
        append_history(group_dir, date_str, version_str, target_name, source_path.name)
    except OSError as exc:
        return ArchiveResult(source_path, "erreur", f"Échec copie/déplacement ou écriture historique : {exc}")

    return ArchiveResult(source_path, "ok", "Archivé.", target=target_path, version=version_str)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Renomme et archive des catalogues téléchargés selon la convention AAAA-MM-JJ_Catalogue_[NomGroupe]_v[X.Y].pdf.")
    parser.add_argument("--dossier-source", type=Path, default=Path.home() / "Downloads", help="Dossier à scanner (défaut : ~/Downloads).")
    parser.add_argument("--dossier-archive", type=Path, default=Path.cwd() / "Catalogues", help="Racine de l'archive, contiendra <NomGroupe>/<Année>/... (défaut : ./Catalogues).")
    parser.add_argument("--extension", default="pdf", help="Extension des fichiers à traiter, sans le point (défaut : pdf).")
    parser.add_argument("--deplacer", action="store_true", help="Déplace les fichiers au lieu de les copier (par défaut : copie, fichier source conservé).")
    parser.add_argument("--date-source", choices=["aujourdhui", "mtime"], default="aujourdhui", help="Date à utiliser dans le nom cible : date du jour, ou date de dernière modification du fichier source (défaut : aujourdhui).")
    parser.add_argument("--version-majeure", action="store_true", help="Incrémente la version majeure (vX.0) au lieu de la version mineure par défaut.")
    parser.add_argument("--recursif", action="store_true", help="Scanne aussi les sous-dossiers du dossier source.")
    parser.add_argument("-v", "--verbose", action="store_true", help="Active les logs de niveau DEBUG.")
    return parser.parse_args(argv)


def configure_logging(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s  %(levelname)-7s  %(message)s",
        datefmt="%H:%M:%S",
        stream=sys.stdout,
    )


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    configure_logging(args.verbose)

    source_dir: Path = args.dossier_source
    archive_root: Path = args.dossier_archive

    if not source_dir.exists():
        logger.error("Dossier source introuvable : %s", source_dir)
        return 1

    archive_root.mkdir(parents=True, exist_ok=True)

    pattern = f"*.{args.extension}"
    candidates = sorted(source_dir.rglob(pattern) if args.recursif else source_dir.glob(pattern))
    logger.info("Dossier source : %s (%d fichier(s) .%s trouvé(s))", source_dir, len(candidates), args.extension)

    results: list[ArchiveResult] = []
    for source_path in candidates:
        try:
            result = process_file(
                source_path, archive_root, args.extension, args.date_source, args.version_majeure, args.deplacer,
            )
        except Exception as exc:  # sécurité : une erreur inattendue sur un fichier ne doit pas arrêter le lot
            logger.exception("Erreur inattendue sur %s", source_path.name)
            result = ArchiveResult(source_path, "erreur", str(exc))
        results.append(result)

        if result.status == "ok":
            logger.info("OK      %s -> %s (v%s)", source_path.name, result.target, result.version)
        elif result.status == "ignore":
            logger.debug("Ignoré  %s (%s)", source_path.name, result.detail)
        else:
            logger.error("Erreur  %s : %s", source_path.name, result.detail)

    ok_results = [r for r in results if r.status == "ok"]
    ignored_results = [r for r in results if r.status == "ignore"]
    error_results = [r for r in results if r.status == "erreur"]

    print()
    print("=== Résumé ===")
    print(f"Fichiers scannés     : {len(results)}")
    print(f"Archivés avec succès : {len(ok_results)}")
    print(f"Ignorés (hors format): {len(ignored_results)}")
    print(f"Erreurs              : {len(error_results)}")
    if ok_results:
        print("\nVersions attribuées :")
        for r in ok_results:
            print(f"  - {r.source.name} -> {r.target.relative_to(archive_root)} (v{r.version})")
    if error_results:
        print("\nErreurs détaillées :")
        for r in error_results:
            print(f"  - {r.source.name} : {r.detail}")

    return 1 if error_results else 0


if __name__ == "__main__":
    sys.exit(main())