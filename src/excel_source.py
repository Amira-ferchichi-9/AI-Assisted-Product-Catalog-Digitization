"""Lecture et validation de la base Layla_Base_Donnees_v1.3_groupes_seuls.xlsx.

Les feuilles Produits / Clients / Relations partagent le même format : 4 lignes
de titre/instructions, une ligne d'en-tête (ligne 5), puis les données à partir
de la ligne 6 (voir feuille 'Mode d'emploi').

La feuille Contrôles contient des compteurs d'anomalies en colonne C : le
classeur ne doit jamais servir à générer un catalogue si l'une d'elles est
différente de 0 — c'est la garde appliquée par validate_controles().
"""

from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass
from pathlib import Path

import openpyxl
import pandas as pd

HEADER_SKIPROWS = 4
CONTROLES_SHEET = "Contrôles"
CONTROLES_FIRST_ROW = 6


class ControlesError(Exception):
    """Levée quand la feuille Contrôles signale au moins une anomalie."""

    def __init__(self, anomalies: list[tuple[str, float]]):
        self.anomalies = anomalies
        message = "Le fichier Excel contient des anomalies non résolues :\n" + "\n".join(
            f"  - {label} : {count:g}" for label, count in anomalies
        )
        super().__init__(message)


def get_controle_anomalies(excel_path: Path) -> list[tuple[str, float]]:
    """Retourne la liste (libellé, nombre d'anomalies) pour chaque compteur != 0."""
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb[CONTROLES_SHEET]

    anomalies: list[tuple[str, float]] = []
    for row in ws.iter_rows(min_row=CONTROLES_FIRST_ROW, max_row=ws.max_row):
        label_cell, _total_cell, anomaly_cell = row[0], row[1], row[2]
        anomaly_value = anomaly_cell.value
        if anomaly_value is None:
            continue
        try:
            anomaly_count = float(anomaly_value)
        except (TypeError, ValueError):
            continue
        if anomaly_count != 0:
            label = str(label_cell.value).strip() if label_cell.value else f"ligne {label_cell.row}"
            anomalies.append((label, anomaly_count))
    return anomalies


def validate_controles(excel_path: Path) -> None:
    """Lève ControlesError si la feuille Contrôles signale une anomalie."""
    anomalies = get_controle_anomalies(excel_path)
    if anomalies:
        raise ControlesError(anomalies)


def _load_sheet(excel_path: Path, sheet_name: str, required_col: str) -> pd.DataFrame:
    df = pd.read_excel(excel_path, sheet_name=sheet_name, skiprows=HEADER_SKIPROWS)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(subset=[required_col])
    return df.reset_index(drop=True)


_DESC_COL_PATTERN = re.compile(r"^description_(\w+)\s*\*?$", re.IGNORECASE)
_ARG_COL_PATTERN = re.compile(r"^argumentaire_(\w+)\s*\*?$", re.IGNORECASE)


def _detect_language_columns(columns: list[str]) -> tuple[str | None, str | None, str | None]:
    """Détecte la colonne description_XX présente (FR, DE, EN, ...) et sa langue."""
    desc_col = None
    lang = None
    for col in columns:
        match = _DESC_COL_PATTERN.match(col)
        if match:
            desc_col = col
            lang = match.group(1).upper()
            break
    arg_col = None
    for col in columns:
        match = _ARG_COL_PATTERN.match(col)
        if match:
            arg_col = col
            break
    return desc_col, arg_col, lang


def load_produits(excel_path: Path) -> tuple[pd.DataFrame, str | None]:
    """Charge la feuille Produits. Retourne (dataframe, langue détectée).

    Le dataframe expose toujours des colonnes normalisées 'description' et
    'argumentaire', quelle que soit la langue réelle de la colonne source
    (description_fr, description_de, ...).
    """
    df = _load_sheet(excel_path, "Produits", "code_article *")
    df = df.rename(columns=lambda c: c.replace(" *", "").strip())

    desc_col, arg_col, lang = _detect_language_columns(list(df.columns))
    df["description"] = df[desc_col] if desc_col else ""
    df["argumentaire"] = df[arg_col] if arg_col else ""

    if "actif" in df.columns:
        df = df[df["actif"].astype(str).str.upper() != "N"]
    return df, lang


def load_clients(excel_path: Path) -> pd.DataFrame:
    df = _load_sheet(excel_path, "Clients", "code_client *")
    df = df.rename(columns=lambda c: c.replace(" *", "").strip())
    if "actif" in df.columns:
        df = df[df["actif"].astype(str).str.upper() != "N"]
    return df


def load_relations(excel_path: Path) -> pd.DataFrame:
    df = _load_sheet(excel_path, "Relations", "code_client *")
    df = df.rename(columns=lambda c: c.replace(" *", "").strip())
    return df


def _parse_date(value) -> dt.date | None:
    if value is None:
        return None
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None
        for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
            try:
                return dt.datetime.strptime(text, fmt).date()
            except ValueError:
                continue
    return None


@dataclass
class PriceResult:
    price_per_kg: float
    negotiated: bool


def compute_final_price(
    code_article: str,
    code_client: str,
    produits_df: pd.DataFrame,
    relations_df: pd.DataFrame,
    on_date: dt.date,
) -> PriceResult | None:
    """Prix négocié valide à on_date si trouvé, sinon prix standard.

    Retourne None si le code_article n'existe pas dans Produits.
    """
    candidates = relations_df[
        (relations_df["code_client"] == code_client) & (relations_df["code_article"] == code_article)
    ]
    for _, row in candidates.iterrows():
        date_debut = _parse_date(row.get("date_debut"))
        date_fin = _parse_date(row.get("date_fin"))
        if date_debut and on_date < date_debut:
            continue
        if date_fin and on_date > date_fin:
            continue
        return PriceResult(price_per_kg=float(row["prix_negocie_kg"]), negotiated=True)

    produit = produits_df[produits_df["code_article"] == code_article]
    if produit.empty:
        return None
    return PriceResult(price_per_kg=float(produit.iloc[0]["prix_standard_kg"]), negotiated=False)
