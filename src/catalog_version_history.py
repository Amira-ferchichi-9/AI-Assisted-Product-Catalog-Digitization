"""Historique des versions de catalogues générés.

Gère la numérotation automatique des versions et le journal des générations
(fichier CSV, sans base de données externe), pour tracer chaque catalogue
produit : standard (sans client) ou personnalisé pour un client donné, avec
la raison du changement.

Convention de nommage des fichiers :
    {Catalogue}_{Client}_v{N}_{AAAA-MM-JJ}.pptx
    ex. Snacks_Marriott_v3_2026-07-28.pptx

Identifiant interne de version :
    {Catalogue}-{Client}-v{N}
"""

from __future__ import annotations

import csv
import datetime as dt
from dataclasses import asdict, dataclass
from pathlib import Path

# Colonnes du journal CSV, dans l'ordre d'écriture.
CHAMPS = ["id_version", "catalogue", "client", "n_version", "date", "type", "raison"]

TYPES_VALIDES = ("statique", "dynamique")


@dataclass
class EntreeVersion:
    """Une ligne du journal : une génération de catalogue."""

    id_version: str
    catalogue: str
    client: str
    n_version: int
    date: str  # horodatage ISO 8601
    type: str  # "statique" ou "dynamique"
    raison: str


def _nom_client(client: str | None) -> str:
    """Normalise le nom du client : absence de client -> "Standard"."""
    return client.strip() if client and client.strip() else "Standard"


def _lire_journal(chemin_journal: Path) -> list[EntreeVersion]:
    """Lit le journal existant ; renvoie une liste vide si le fichier n'existe pas encore."""
    if not chemin_journal.exists():
        return []
    with chemin_journal.open("r", encoding="utf-8", newline="") as f:
        lecteur = csv.DictReader(f)
        return [
            EntreeVersion(
                id_version=ligne["id_version"],
                catalogue=ligne["catalogue"],
                client=ligne["client"],
                n_version=int(ligne["n_version"]),
                date=ligne["date"],
                type=ligne["type"],
                raison=ligne["raison"],
            )
            for ligne in lecteur
        ]


def obtenir_prochain_numero_version(
    chemin_journal: Path, catalogue: str, client: str | None = None
) -> int:
    """Calcule le prochain numéro de version pour ce couple (catalogue, client).

    Cherche la version la plus élevée déjà enregistrée pour ce couple dans le
    journal et renvoie ce numéro + 1 (ou 1 si aucune version n'existe encore).
    À appeler avant chaque génération.
    """
    client_normalise = _nom_client(client)
    versions_existantes = [
        entree.n_version
        for entree in _lire_journal(chemin_journal)
        if entree.catalogue == catalogue and entree.client == client_normalise
    ]
    return max(versions_existantes, default=0) + 1


def construire_id_version(catalogue: str, client: str | None, n_version: int) -> str:
    """Construit l'identifiant interne de version : {Catalogue}-{Client}-v{N}."""
    return f"{catalogue}-{_nom_client(client)}-v{n_version}"


def construire_nom_fichier(
    catalogue: str, client: str | None, n_version: int, date: dt.date | None = None
) -> str:
    """Construit le nom de fichier : {Catalogue}_{Client}_v{N}_{AAAA-MM-JJ}.pptx."""
    date = date or dt.date.today()
    return f"{catalogue}_{_nom_client(client)}_v{n_version}_{date.isoformat()}.pptx"


def enregistrer_version(
    chemin_journal: Path,
    catalogue: str,
    client: str | None,
    n_version: int,
    type_version: str,
    raison: str,
    date: dt.datetime | None = None,
) -> EntreeVersion:
    """Ajoute une entrée au journal après la génération d'un catalogue.

    - type_version doit être "statique" (changement générique, ex. mise à jour
      du template ou des prix standard) ou "dynamique" (personnalisation propre
      à un client — dans ce cas, `client` est obligatoire).
    - Pour un changement "statique", seuls date + raison + n_version comptent
      vraiment (client reste "Standard" si non précisé).
    - Pour une modification "dynamique", le nom du client concerné est en plus
      obligatoire.
    """
    if type_version not in TYPES_VALIDES:
        raise ValueError(f"type_version invalide : '{type_version}' (attendu : {TYPES_VALIDES})")

    client_normalise = _nom_client(client)
    if type_version == "dynamique" and client_normalise == "Standard":
        raise ValueError("Une modification 'dynamique' doit préciser le client concerné.")

    entree = EntreeVersion(
        id_version=construire_id_version(catalogue, client, n_version),
        catalogue=catalogue,
        client=client_normalise,
        n_version=n_version,
        date=(date or dt.datetime.now()).isoformat(timespec="seconds"),
        type=type_version,
        raison=raison,
    )

    fichier_existe = chemin_journal.exists()
    chemin_journal.parent.mkdir(parents=True, exist_ok=True)
    with chemin_journal.open("a", encoding="utf-8", newline="") as f:
        ecrivain = csv.DictWriter(f, fieldnames=CHAMPS)
        if not fichier_existe:
            ecrivain.writeheader()
        ecrivain.writerow(asdict(entree))

    return entree


def obtenir_historique(
    chemin_journal: Path, catalogue: str, client: str | None = None
) -> list[EntreeVersion]:
    """Renvoie l'historique d'un catalogue (et, si précisé, d'un client), trié par date décroissante.

    Si `client` n'est pas précisé, renvoie l'historique toutes versions
    confondues (standard + tous les clients) pour ce catalogue.
    """
    client_normalise = _nom_client(client) if client is not None else None
    entrees = [
        entree
        for entree in _lire_journal(chemin_journal)
        if entree.catalogue == catalogue and (client_normalise is None or entree.client == client_normalise)
    ]
    return sorted(entrees, key=lambda e: e.date, reverse=True)


if __name__ == "__main__":
    # --- Exemple d'utilisation ---
    journal = Path("historique_versions.csv")

    # 1. Génération d'un catalogue standard (changement générique du template)
    n = obtenir_prochain_numero_version(journal, "Snacks", client=None)
    nom_fichier = construire_nom_fichier("Snacks", None, n)
    enregistrer_version(journal, "Snacks", None, n, "statique", "Mise à jour des prix du template")
    print(f"Généré : {nom_fichier}")

    # 2. Génération d'un catalogue personnalisé pour un client
    n = obtenir_prochain_numero_version(journal, "Snacks", client="Marriott")
    nom_fichier = construire_nom_fichier("Snacks", "Marriott", n)
    enregistrer_version(journal, "Snacks", "Marriott", n, "dynamique", "Ajout du prix négocié Marriott")
    print(f"Généré : {nom_fichier}")

    # 3. Consultation de l'historique du client Marriott pour ce catalogue
    print("\nHistorique Snacks / Marriott :")
    for entree in obtenir_historique(journal, "Snacks", client="Marriott"):
        print(f"  {entree.date}  {entree.id_version}  ({entree.type})  {entree.raison}")
