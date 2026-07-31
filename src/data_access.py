"""Lecture des feuilles Produits / Clients / Relations de la base Excel Layla.

Les 3 feuilles partagent le même format : 4 lignes de titre/instructions,
puis la ligne d'en-tête, puis les données (voir feuille 'Mode d'emploi').
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

HEADER_SKIPROWS = 4


def _load_sheet(excel_path: Path, sheet_name: str, required_col: str) -> pd.DataFrame:
    df = pd.read_excel(excel_path, sheet_name=sheet_name, skiprows=HEADER_SKIPROWS)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(subset=[required_col])
    return df.reset_index(drop=True)


def load_produits(excel_path: Path) -> pd.DataFrame:
    df = _load_sheet(excel_path, "Produits", "code_article *")
    df = df.rename(columns=lambda c: c.replace(" *", "").strip())
    if "actif" in df.columns:
        df = df[df["actif"].astype(str).str.upper() != "N"]
    return df


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


GROUPE_COLUMN = "Catalogue/ Groupe"
ASTORE_CODE_CLIENT = "ASTORE"


def _select_produit_row(produits_df: pd.DataFrame, code_article: str, code_client: str | None) -> pd.Series | None:
    """Choisit la bonne ligne Produits pour ce code_article quand plusieurs lignes
    existent (une par variante de catalogue/groupe, ex. "Deutschland" standard vs
    "ASTORE Deutschland") : la ligne du groupe du client s'il en fait partie,
    sinon la ligne standard (hors groupe)."""
    matches = produits_df[produits_df["code_article"] == code_article]
    if matches.empty:
        return None
    if GROUPE_COLUMN in matches.columns and len(matches) > 1:
        est_astore = (code_client or "").strip().upper() == ASTORE_CODE_CLIENT
        groupe = matches[GROUPE_COLUMN].astype(str).str.upper()
        astore_rows = matches[groupe.str.contains(ASTORE_CODE_CLIENT)]
        standard_rows = matches[~groupe.str.contains(ASTORE_CODE_CLIENT)]
        if est_astore and not astore_rows.empty:
            return astore_rows.iloc[0]
        if not est_astore and not standard_rows.empty:
            return standard_rows.iloc[0]
    return matches.iloc[0]


def get_price(
    code_article: str,
    code_client: str | None,
    produits_df: pd.DataFrame,
    relations_df: pd.DataFrame,
) -> tuple[float | None, bool]:
    """Retourne (prix, est_negocie). prix=None si code_article introuvable."""
    if code_client:
        match = relations_df[
            (relations_df["code_client"] == code_client)
            & (relations_df["code_article"] == code_article)
        ]
        if not match.empty:
            return float(match.iloc[0]["prix_negocie_kg"]), True

    produit = _select_produit_row(produits_df, code_article, code_client)
    if produit is None:
        return None, False
    return float(produit["prix_standard_kg"]), False


def get_client_logo_filename(code_client: str, clients_df: pd.DataFrame) -> str | None:
    match = clients_df[clients_df["code_client"] == code_client]
    if match.empty:
        return None
    value = match.iloc[0].get("logo_client_fichier")
    if pd.isna(value) or not str(value).strip():
        return None
    return str(value).strip()
