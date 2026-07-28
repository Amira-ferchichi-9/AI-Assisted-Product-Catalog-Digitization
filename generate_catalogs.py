#!/usr/bin/env python
"""Génère un catalogue PPTX par client actif à partir de la base Excel Layla.

Usage :
    python generate_catalogs.py --excel-path "C:\\...\\Layla_Base_Donnees_v1.3_groupes_seuls.xlsx" --output-dir "C:\\...\\output\\catalogues"

Conçu pour être branché sur une tâche planifiée qui relit le même chemin
(ex. un fichier synchronisé OneDrive) à intervalle régulier : code de sortie
non nul si le fichier est introuvable, si la feuille Contrôles signale une
anomalie, ou si une erreur est survenue pendant la génération d'au moins un
client. Les photos manquantes sont journalisées comme avertissements (le
catalogue est quand même généré, avec un encart "Photo à venir") et n'entrent
pas dans le code de sortie.
"""

from __future__ import annotations

import argparse
import datetime as dt
import logging
import sys
from pathlib import Path

SRC_DIR = Path(__file__).resolve().parent / "src"
sys.path.insert(0, str(SRC_DIR))

import config  # noqa: E402
from catalog_version_history import (  # noqa: E402
    construire_nom_fichier,
    enregistrer_version,
    obtenir_prochain_numero_version,
)
from excel_source import ControlesError, load_clients, load_produits, load_relations, validate_controles  # noqa: E402
from pptx_builder import build_catalog_for_client  # noqa: E402

PROJECT_ROOT = Path(__file__).resolve().parent
DEFAULT_LOGO_DIR = PROJECT_ROOT / "logo"
DEFAULT_PHOTOS_DIR = PROJECT_ROOT / "photos"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "output" / "catalogues"
DEFAULT_LAYLA_LOGO = SRC_DIR / "assets" / "layla_logo.png"

logger = logging.getLogger("layla_catalog")


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Génère un catalogue PPTX par client à partir de la base Excel Layla.")
    parser.add_argument("--excel-path", required=True, type=Path, help="Chemin du fichier Excel source (Produits/Clients/Relations/Contrôles).")
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR, help=f"Dossier de sortie des PPTX (défaut : {DEFAULT_OUTPUT_DIR}).")
    parser.add_argument("--photos-dir", type=Path, default=DEFAULT_PHOTOS_DIR, help=f"Dossier des photos produits, nommées exactement comme photo_fichier (défaut : {DEFAULT_PHOTOS_DIR}).")
    parser.add_argument("--logo-dir", type=Path, default=DEFAULT_LOGO_DIR, help=f"Dossier des logos clients (défaut : {DEFAULT_LOGO_DIR}).")
    parser.add_argument("--layla-logo", type=Path, default=DEFAULT_LAYLA_LOGO, help="Chemin du logo Layla.")
    parser.add_argument("--catalogue", type=str, default=None, help="Filtre optionnel sur la colonne 'catalogue' de la feuille Produits (ex. Apéritifs). Par défaut : tous les produits actifs.")
    parser.add_argument("--date", type=str, default=None, help="Date de référence AAAA-MM-JJ pour le calcul des prix négociés (défaut : aujourd'hui).")
    parser.add_argument("--client", type=str, default=None, help="Ne générer que pour ce code_client (défaut : tous les clients actifs).")
    parser.add_argument(
        "--raison", type=str, default="Génération automatique (tâche planifiée)",
        help="Raison enregistrée dans l'historique des versions pour cette exécution.",
    )
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

    excel_path: Path = args.excel_path
    if not excel_path.exists():
        logger.error("Fichier Excel introuvable : %s", excel_path)
        return 1

    on_date = dt.date.today()
    if args.date:
        try:
            on_date = dt.datetime.strptime(args.date, "%Y-%m-%d").date()
        except ValueError:
            logger.error("Format de --date invalide (attendu AAAA-MM-JJ) : %s", args.date)
            return 1

    logger.info("Lecture de la base : %s", excel_path)

    try:
        validate_controles(excel_path)
    except ControlesError as exc:
        logger.error("Contrôles de complétude en échec — génération annulée.")
        for label, count in exc.anomalies:
            logger.error("  - %s : %g", label, count)
        return 1
    logger.info("Contrôles de complétude : OK (aucune anomalie).")

    produits_df, lang = load_produits(excel_path)
    clients_df = load_clients(excel_path)
    relations_df = load_relations(excel_path)
    logger.info(
        "Chargé : %d produits actifs (langue détectée : %s), %d clients actifs, %d lignes de négociation.",
        len(produits_df), lang or "inconnue", len(clients_df), len(relations_df),
    )

    if args.catalogue:
        before = len(produits_df)
        produits_df = produits_df[produits_df["catalogue"] == args.catalogue]
        logger.info("Filtre --catalogue=%s : %d/%d produits retenus.", args.catalogue, len(produits_df), before)

    if produits_df.empty:
        logger.error("Aucun produit actif à traiter — génération annulée.")
        return 1

    if args.client:
        clients_df = clients_df[clients_df["code_client"] == args.client]
        if clients_df.empty:
            logger.error("Client '%s' introuvable ou inactif dans la feuille Clients.", args.client)
            return 1

    args.output_dir.mkdir(parents=True, exist_ok=True)
    if not args.photos_dir.exists():
        logger.warning("Dossier photos introuvable (%s) — toutes les cartes produit afficheront un encart 'Photo à venir'.", args.photos_dir)

    nom_catalogue = args.catalogue or "Catalogue"

    had_errors = False
    total_warnings = 0
    clients_done = 0

    for _, client_row in clients_df.iterrows():
        code_client = client_row["code_client"]
        nom_client = client_row["nom_client"]
        try:
            prs, warnings = build_catalog_for_client(
                client_row, produits_df, relations_df,
                args.photos_dir, args.logo_dir, args.layla_logo, on_date,
            )
            n_version = obtenir_prochain_numero_version(config.VERSION_HISTORY_PATH, nom_catalogue, client=nom_client)
            output_path = args.output_dir / construire_nom_fichier(nom_catalogue, nom_client, n_version)
            prs.save(output_path)
            enregistrer_version(
                config.VERSION_HISTORY_PATH, nom_catalogue, nom_client, n_version, "dynamique", args.raison
            )
        except Exception:
            logger.exception("Échec de la génération pour le client '%s'.", code_client)
            had_errors = True
            continue

        clients_done += 1
        total_warnings += len(warnings)
        logger.info("Client %s -> %s (%d slide(s), %d avertissement(s)).", code_client, output_path.name, len(prs.slides), len(warnings))
        for w in warnings:
            logger.warning("  [%s] %s", code_client, w)

    logger.info(
        "Terminé : %d/%d client(s) généré(s), %d avertissement(s) au total, dossier de sortie : %s",
        clients_done, len(clients_df), total_warnings, args.output_dir,
    )

    return 1 if had_errors else 0


if __name__ == "__main__":
    sys.exit(main())