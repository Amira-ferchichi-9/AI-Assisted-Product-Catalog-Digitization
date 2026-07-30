"""Historique de téléchargement et historique de versions des catalogues.

Les prix ne sont pas gérés ici : tout changement de prix se fait directement
dans la base Excel. Ce module gère uniquement deux journaux distincts :

  1. Historique de téléchargement : une ligne à CHAQUE génération/téléchargement
     d'un catalogue (Standard ou personnalisé pour un groupe/client).
  2. Historique de versions : une ligne à chaque changement de VERSION du
     catalogue (nouveaux articles, photos/descriptions optimisées, logo ou
     photo dynamique modifiés, ou révision annuelle des prix).

Format du numéro de version : AAMM.NN (ex. 2604.01 = version 01, année 2026,
mois 04).
  - Le préfixe AAMM ne change que lors d'un changement de type "prix" (une
    fois par an) : il prend l'année+mois courants et NN repart à 01.
  - Le suffixe NN s'incrémente de 1 pour un changement de type "contenu"
    (nouveaux articles, photos/descriptions, logos...), sans toucher à AAMM.
"""

from __future__ import annotations

import csv
import datetime as dt
from dataclasses import asdict, dataclass
from pathlib import Path

TYPES_CHANGEMENT = ("prix", "contenu")

CHAMPS_VERSION = ["numero_version", "date", "raison"]
CHAMPS_TELECHARGEMENT = ["version", "groupe_standard", "raison", "langue", "pays", "date", "nom_fichier"]


class RaisonManquanteError(ValueError):
    """La raison (obligatoire) est vide ou absente."""


@dataclass
class EntreeVersion:
    """Une ligne de l'historique de versions."""

    numero_version: str
    date: str
    raison: str


@dataclass
class EntreeTelechargement:
    """Une ligne de l'historique de téléchargement."""

    version: str
    groupe_standard: str
    raison: str
    langue: str
    pays: str
    date: str
    nom_fichier: str


def _lire_csv(chemin: Path, champs: list[str]) -> list[dict]:
    """Lit un journal CSV existant ; renvoie une liste vide si le fichier n'existe pas encore."""
    if not chemin.exists():
        return []
    with chemin.open("r", encoding="utf-8", newline="") as f:
        return list(csv.DictReader(f))


def _ajouter_ligne_csv(chemin: Path, champs: list[str], ligne: dict) -> None:
    """Ajoute une ligne à la fin d'un journal CSV, en créant le fichier et l'en-tête si besoin."""
    fichier_existe = chemin.exists()
    chemin.parent.mkdir(parents=True, exist_ok=True)
    with chemin.open("a", encoding="utf-8", newline="") as f:
        ecrivain = csv.DictWriter(f, fieldnames=champs)
        if not fichier_existe:
            ecrivain.writeheader()
        ecrivain.writerow(ligne)


def _lire_historique_versions(chemin_historique_versions: Path) -> list[EntreeVersion]:
    return [EntreeVersion(**ligne) for ligne in _lire_csv(chemin_historique_versions, CHAMPS_VERSION)]


def _lire_historique_telechargements(chemin_historique_telechargements: Path) -> list[EntreeTelechargement]:
    return [
        EntreeTelechargement(**ligne)
        for ligne in _lire_csv(chemin_historique_telechargements, CHAMPS_TELECHARGEMENT)
    ]


def _parser_version(numero_version: str) -> tuple[str, int]:
    """Décompose "AAMM.NN" en (AAMM, NN)."""
    aamm, nn = numero_version.split(".")
    return aamm, int(nn)


def _formater_version(aamm: str, nn: int) -> str:
    """Recompose (AAMM, NN) en "AAMM.NN", NN toujours sur 2 chiffres."""
    return f"{aamm}.{nn:02d}"


def obtenir_version_courante(chemin_historique_versions: Path) -> str | None:
    """Renvoie le numéro de la dernière version enregistrée, ou None si aucune n'existe encore."""
    historique = _lire_historique_versions(chemin_historique_versions)
    return historique[-1].numero_version if historique else None


def calculer_prochaine_version(
    chemin_historique_versions: Path, type_changement: str, date: dt.date | None = None
) -> str:
    """Calcule le prochain numéro de version selon le type de changement.

    - "prix" : nouveau préfixe AAMM (année+mois courants de `date`), NN repart à 01
      (une révision de prix a lieu une fois par an et redémarre la numérotation).
    - "contenu" : garde le préfixe AAMM de la dernière version enregistrée, NN + 1
      (nouveaux articles, photos/descriptions, logos...).
    - S'il n'existe encore aucune version enregistrée, renvoie AAMM.01 avec le
      AAMM courant, quel que soit le type de changement (amorçage du journal).
    """
    if type_changement not in TYPES_CHANGEMENT:
        raise ValueError(f"type_changement invalide : '{type_changement}' (attendu : {TYPES_CHANGEMENT})")

    date = date or dt.date.today()
    aamm_courant = date.strftime("%y%m")

    version_actuelle = obtenir_version_courante(chemin_historique_versions)
    if version_actuelle is None or type_changement == "prix":
        return _formater_version(aamm_courant, 1)

    aamm_precedent, nn_precedent = _parser_version(version_actuelle)
    return _formater_version(aamm_precedent, nn_precedent + 1)


def enregistrer_version(
    chemin_historique_versions: Path,
    numero_version: str,
    raison: str,
    date: dt.date | None = None,
) -> EntreeVersion:
    """Ajoute une ligne à l'historique de versions, après un changement de version réel.

    La raison est obligatoire (lève RaisonManquanteError si vide).
    """
    if not raison or not raison.strip():
        raise RaisonManquanteError("La raison du changement de version est obligatoire.")

    entree = EntreeVersion(
        numero_version=numero_version,
        date=(date or dt.date.today()).isoformat(),
        raison=raison.strip(),
    )
    _ajouter_ligne_csv(chemin_historique_versions, CHAMPS_VERSION, asdict(entree))
    return entree


def enregistrer_telechargement(
    chemin_historique_telechargements: Path,
    version: str,
    groupe_standard: str,
    raison: str,
    langue: str,
    pays: str,
    nom_fichier: str,
    date: dt.datetime | None = None,
) -> EntreeTelechargement:
    """Ajoute une ligne à l'historique de téléchargement, à chaque génération.

    `raison` ("pour quel client / pourquoi cette génération") est un champ
    obligatoire : lève RaisonManquanteError si elle est vide, ce qui doit
    empêcher la génération côté interface (aucune ligne n'est écrite).
    """
    if not raison or not raison.strip():
        raise RaisonManquanteError("La raison de la génération est obligatoire.")

    entree = EntreeTelechargement(
        version=version,
        groupe_standard=groupe_standard or "Standard",
        raison=raison.strip(),
        langue=langue,
        pays=pays,
        date=(date or dt.datetime.now()).isoformat(timespec="seconds"),
        nom_fichier=nom_fichier,
    )
    _ajouter_ligne_csv(chemin_historique_telechargements, CHAMPS_TELECHARGEMENT, asdict(entree))
    return entree


def obtenir_historique_versions(chemin_historique_versions: Path) -> list[EntreeVersion]:
    """Renvoie l'historique de versions, trié par date décroissante."""
    return sorted(_lire_historique_versions(chemin_historique_versions), key=lambda e: e.date, reverse=True)


def obtenir_historique_telechargements(chemin_historique_telechargements: Path) -> list[EntreeTelechargement]:
    """Renvoie l'historique de téléchargement, trié par date décroissante."""
    return sorted(
        _lire_historique_telechargements(chemin_historique_telechargements), key=lambda e: e.date, reverse=True
    )


if __name__ == "__main__":
    # --- Exemple d'utilisation ---
    historique_versions = Path("historique_versions.csv")
    historique_telechargements = Path("historique_telechargements.csv")

    # 1. Version de base (changement de contenu initial)
    v = calculer_prochaine_version(historique_versions, "contenu")
    enregistrer_version(historique_versions, v, "Version de base du catalogue")
    print(f"Version enregistrée : {v}")

    # 2. Génération/téléchargement pour le groupe Marriott, à la version courante
    enregistrer_telechargement(
        historique_telechargements,
        version=obtenir_version_courante(historique_versions),
        groupe_standard="Marriott",
        raison="Envoi au directeur des achats Marriott",
        langue="Français",
        pays="Tunisie",
        nom_fichier=f"Aperitifs_Marriott_v{v}_2026-07-30.pptx",
    )

    # 3. Nouveaux articles ajoutés -> nouvelle sous-version de contenu
    v = calculer_prochaine_version(historique_versions, "contenu")
    enregistrer_version(historique_versions, v, "Ajout de 3 nouveaux articles")
    print(f"Version enregistrée : {v}")

    # 4. Révision annuelle des prix -> nouveau préfixe AAMM, NN repart à 01
    v = calculer_prochaine_version(historique_versions, "prix")
    enregistrer_version(historique_versions, v, "Révision annuelle des prix")
    print(f"Version enregistrée : {v}")

    # 5. Consultation des historiques
    print("\nHistorique de versions :")
    for e in obtenir_historique_versions(historique_versions):
        print(f"  {e.date}  v{e.numero_version}  {e.raison}")

    print("\nHistorique de téléchargements :")
    for e in obtenir_historique_telechargements(historique_telechargements):
        print(f"  {e.date}  v{e.version}  {e.groupe_standard}  ({e.raison})")
