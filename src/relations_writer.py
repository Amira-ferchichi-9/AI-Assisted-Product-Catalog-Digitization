"""Écriture des prix négociés dans la feuille Relations de la base Excel Layla.

Les 200 lignes de données (6 à 205) existent déjà dans le classeur avec la formule
d'écart pré-remplie en colonne G : ces fonctions se contentent de remplir ou vider
les colonnes A-F de la ligne correspondante, sans jamais recréer cette formule.
"""

from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path

import openpyxl
import pandas as pd

from data_access import load_clients, load_relations

RELATIONS_SHEET = "Relations"
RELATIONS_FIRST_ROW = 6
BACKUP_DIR_NAME = "backups"


class RelationsWriteError(Exception):
    """Erreur explicite lors de l'écriture dans la feuille Relations."""


class FileChangedError(RelationsWriteError):
    """Le fichier Excel a été modifié depuis le chargement de la page."""


def backup_excel(excel_path: Path) -> Path:
    backup_dir = excel_path.parent / BACKUP_DIR_NAME
    backup_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup_path = backup_dir / f"{excel_path.stem}_{timestamp}{excel_path.suffix}"
    shutil.copy2(excel_path, backup_path)
    return backup_path


def get_negotiated_price(excel_path: Path, code_client: str, code_article: str) -> float | None:
    relations_df = load_relations(excel_path)
    match = relations_df[
        (relations_df["code_client"] == code_client) & (relations_df["code_article"] == code_article)
    ]
    if match.empty:
        return None
    value = match.iloc[0]["prix_negocie_kg"]
    if pd.isna(value):
        return None
    return float(value)


def _check_expected_mtime(excel_path: Path, expected_mtime: float | None) -> None:
    if expected_mtime is None:
        return
    current_mtime = excel_path.stat().st_mtime
    if current_mtime != expected_mtime:
        raise FileChangedError(
            "Le fichier Excel a été modifié depuis le chargement de la page : "
            "rechargez la page avant d'enregistrer pour ne pas écraser un changement récent."
        )


def _check_client_is_groupe(excel_path: Path, code_client: str) -> None:
    clients_df = load_clients(excel_path)
    match = clients_df[clients_df["code_client"] == code_client]
    if match.empty:
        raise RelationsWriteError(f"Client '{code_client}' introuvable dans la feuille Clients.")
    type_client = match.iloc[0].get("type_client")
    if type_client is not None and not pd.isna(type_client) and str(type_client).strip() != "Groupe":
        raise RelationsWriteError(
            f"La négociation doit être saisie au niveau du groupe, pas pour '{code_client}' "
            f"(type_client = '{type_client}')."
        )


def _find_relations_row(ws, code_client: str, code_article: str) -> int | None:
    for row_idx in range(RELATIONS_FIRST_ROW, ws.max_row + 1):
        if (
            ws.cell(row=row_idx, column=1).value == code_client
            and ws.cell(row=row_idx, column=2).value == code_article
        ):
            return row_idx
    return None


def _find_empty_relations_row(ws) -> int | None:
    for row_idx in range(RELATIONS_FIRST_ROW, ws.max_row + 1):
        if ws.cell(row=row_idx, column=1).value in (None, ""):
            return row_idx
    return None


def upsert_negotiated_price(
    excel_path: Path,
    code_client: str,
    code_article: str,
    prix_negocie_kg: float,
    expected_mtime: float | None = None,
) -> None:
    """Crée ou met à jour la ligne Relations pour ce couple client/produit."""
    _check_expected_mtime(excel_path, expected_mtime)
    _check_client_is_groupe(excel_path, code_client)

    backup_excel(excel_path)

    wb = openpyxl.load_workbook(excel_path)
    ws = wb[RELATIONS_SHEET]

    row_idx = _find_relations_row(ws, code_client, code_article)
    if row_idx is None:
        row_idx = _find_empty_relations_row(ws)
        if row_idx is None:
            raise RelationsWriteError(
                "Plus aucune ligne disponible dans la feuille Relations (limite de "
                f"{ws.max_row - RELATIONS_FIRST_ROW + 1} lignes de négociation atteinte) — "
                "contactez le support pour étendre le classeur."
            )
        ws.cell(row=row_idx, column=1, value=code_client)
        ws.cell(row=row_idx, column=2, value=code_article)

    ws.cell(row=row_idx, column=3, value=float(prix_negocie_kg))

    wb.save(excel_path)


def clear_negotiated_price(
    excel_path: Path,
    code_client: str,
    code_article: str,
    expected_mtime: float | None = None,
) -> None:
    """Supprime la ligne Relations pour ce couple client/produit (retour au prix standard)."""
    _check_expected_mtime(excel_path, expected_mtime)

    wb = openpyxl.load_workbook(excel_path)
    ws = wb[RELATIONS_SHEET]

    row_idx = _find_relations_row(ws, code_client, code_article)
    if row_idx is None:
        return

    backup_excel(excel_path)
    for column in range(1, 7):  # A-F
        ws.cell(row=row_idx, column=column).value = None

    wb.save(excel_path)
