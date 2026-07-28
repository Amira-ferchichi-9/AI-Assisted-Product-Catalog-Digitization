"""Chemins centraux du projet. Le dossier racine contient les fichiers de travail
réels (Excel, PPTX, logos) tels que déposés par l'utilisateur — on ne les déplace pas."""

import os
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent

# Le fichier maître vit sur OneDrive (synchronisé localement à ce chemin fixe) :
# le directeur commercial l'édite dans Excel, et l'app lit toujours sa dernière
# version enregistrée, sans étape manuelle. Reste surchargeable par la variable
# d'environnement LAYLA_EXCEL_PATH pour un autre poste/déploiement futur — mais
# NE PAS compter dessus par défaut : une variable d'environnement ne survit pas
# à la fermeture du terminal, contrairement à ce chemin en dur.
DEFAULT_EXCEL_PATH = Path.home() / "OneDrive" / "Layla_Base_Donnees_v1.3_groupes_seuls.xlsx"
EXCEL_PATH = Path(os.environ.get("LAYLA_EXCEL_PATH", DEFAULT_EXCEL_PATH))
TEMPLATE_PATH = PROJECT_ROOT / "Layla_Apero_Catalogue_FR.pptx"
LOGO_DIR = PROJECT_ROOT / "logo"
OUTPUT_DIR = PROJECT_ROOT / "output"
PRICE_MAP_PATH = PROJECT_ROOT / "config" / "gamme_price_map.json"

SAMPLE_EXCEL_PATH = PROJECT_ROOT / "data" / "sample_produits.xlsx"

COVER_SLIDE_INDEX = 0

# Nom de catalogue utilisé pour la convention de nommage et le journal de versions
# (voir catalog_version_history.py). Un seul catalogue existe pour l'instant.
CATALOGUE_NAME = "Aperitifs"
VERSION_HISTORY_PATH = PROJECT_ROOT / "historique_versions.csv"
