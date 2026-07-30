"""Chemins centraux du projet. Le dossier racine contient les fichiers de travail
réels (Excel, PPTX, logos) tels que déposés par l'utilisateur — on ne les déplace pas."""

import os
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent

# Le chemin de la base Excel peut être surchargé par la variable d'environnement
# LAYLA_EXCEL_PATH, sans toucher au code — utile pour pointer vers un dossier
# OneDrive synchronisé (ex. C:\Users\<vous>\OneDrive\...\Layla_Base_Donnees.xlsx)
# une fois celui-ci mis en place, ou vers tout autre emplacement plus tard.
EXCEL_PATH = Path(os.environ.get("LAYLA_EXCEL_PATH", PROJECT_ROOT / "Layla_Base_Donnees.xlsx"))
TEMPLATE_PATH = PROJECT_ROOT / "Layla_Apero_Catalogue_FR.pptx"
LOGO_DIR = PROJECT_ROOT / "logo"
OUTPUT_DIR = PROJECT_ROOT / "output"
PRICE_MAP_PATH = PROJECT_ROOT / "config" / "gamme_price_map.json"

SAMPLE_EXCEL_PATH = PROJECT_ROOT / "data" / "sample_produits.xlsx"

COVER_SLIDE_INDEX = 0

# Nom de catalogue utilisé pour la convention de nommage et les journaux
# (voir catalog_version_history.py). Un seul catalogue existe pour l'instant.
CATALOGUE_NAME = "Aperitifs"
VERSION_HISTORY_PATH = PROJECT_ROOT / "historique_versions.csv"
DOWNLOAD_HISTORY_PATH = PROJECT_ROOT / "historique_telechargements.csv"
