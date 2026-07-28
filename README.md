# Générateur de catalogue — Apéritifs (Layla)

Personnalise automatiquement le deck `Layla_Apero_Catalogue_FR.pptx` pour un client donné :

1. **Logo de couverture** : si le client a un logo renseigné (feuille `Clients`,
   colonne `logo_client_fichier`), il est ajouté à côté du logo Layla sur la slide 1.
   Sinon, la couverture reste standard (logo Layla seul).
2. **Prix** : les 3 pages "gamme Perle croustillante" (slides 4, 5, 6) affichent
   le prix négocié du client (feuille `Relations`) s'il existe, sinon le prix
   standard (feuille `Produits`).

Le reste du deck (9 slides) n'est jamais modifié : la mise en page d'origine est conservée telle quelle.

## Comment ça marche

- **Base Excel réelle** : `C:\Users\<vous>\OneDrive\Layla_Base_Donnees_v1.3_groupes_seuls.xlsx`
  (feuilles `Produits`, `Clients`, `Relations`). C'est désormais le chemin **par
  défaut** codé dans `src/config.py` (`DEFAULT_EXCEL_PATH`) — pas besoin de
  variable d'environnement pour l'usage normal. Le fichier étant synchronisé par
  OneDrive, l'app lit toujours sa dernière version enregistrée automatiquement.

  **Piège déjà rencontré** : avant, ce chemin dépendait d'une variable
  d'environnement (`LAYLA_EXCEL_PATH`) à redéfinir à chaque ouverture de
  terminal — si on oubliait, l'app retombait silencieusement sur une ancienne
  copie locale quasi vide, et les changements de prix semblaient "ne rien
  faire" (ils étaient bien écrits, mais dans le mauvais fichier). Le chemin
  par défaut pointe maintenant directement sur le vrai fichier OneDrive pour
  éviter que ça se reproduise ; `LAYLA_EXCEL_PATH` reste disponible pour
  surcharger ponctuellement (autre poste, tests), mais n'est plus nécessaire
  au quotidien.
- `Layla_Apero_Catalogue_FR.pptx` : le template PowerPoint, jamais modifié en place
  (chaque génération produit un nouveau fichier dans `output/`).
- `logo/` : logos clients (nom de fichier = valeur de `logo_client_fichier`).
- `config/gamme_price_map.json` : fait le lien entre une zone de prix du PPTX
  (numéro de slide + shape) et un `code_article` de la feuille `Produits`.
- `data/sample_produits.xlsx` : copie de démonstration de la base, avec les
  3 `code_article` attendus déjà renseignés, pour tester l'outil sans toucher
  au fichier réel.

## Installation

```powershell
pip install python-pptx openpyxl pandas streamlit pillow
```

## Lancer l'interface

```powershell
streamlit run app.py
```

Un navigateur s'ouvre : choisir un client dans la liste, cliquer sur
**"Générer le catalogue"**, puis télécharger le fichier `.pptx` produit
(également enregistré dans `output/`).

Le fichier généré suit la convention `AAAA-MM-JJ_Catalogue_[CodeClient]_v[X.Y].pptx`
(ex. `2026-07-29_Catalogue_MARRIOTT_v1.0.pptx`) — la version est calculée
automatiquement par client à partir des fichiers déjà présents dans `output/`
(voir `src/versioning.py`, la même logique que `archive_catalogues.py`).
Chaque génération dans la journée pour un même client incrémente la version
mineure (v1.0 -> v1.1 -> ...) ; un fichier existant n'est jamais écrasé.

Pour tester avec les données d'exemple plutôt que le fichier réel, activer le
bouton "Utiliser les données d'exemple" en haut de la page.

## Limites actuelles / À faire avant usage en production

1. **Le fichier `logo_client_fichier` doit correspondre exactement** à un nom
   de fichier présent dans `logo/` (ex. `Logo Marriott Tunis .png`, avec
   l'espace et l'extension exacts) — sinon le logo est silencieusement ignoré
   (avertissement affiché, mais pas d'erreur bloquante).
2. **Seuls 3 emplacements de prix sont automatisés** (le prix "TND/KG" en tête
   des slides 4, 5 et 6, mappés dans `config/gamme_price_map.json`). Les prix
   comparatifs de la slide 7 et le "prix par unité" dans les blocs
   "Conditionnement" restent fixes — leur logique de calcul n'était pas
   évidente à déduire du template et mériterait d'être confirmée avant
   automatisation.
3. Le générateur ne touche à rien d'autre que le logo de couverture et ces
   3 zones de prix : toute autre personnalisation (textes, images produit,
   ajout/suppression de slides) est hors périmètre de ce MVP.
4. **Effet de bord connu de l'écriture des prix négociés** (`src/relations_writer.py`) :
   comme elle sauvegarde le classeur avec openpyxl (qui ne recalcule pas les
   formules), les valeurs mises en cache des formules ailleurs dans le
   classeur — colonne "écart vs standard" de `Relations`, compteurs de la
   feuille `Contrôles` — peuvent apparaître vides jusqu'à la prochaine
   ouverture/sauvegarde du fichier dans Excel (qui les recalcule). Cela ne
   touche pas les prix eux-mêmes (valeurs saisies directement, pas des
   formules), mais peut fausser temporairement la lecture visuelle de ces
   deux colonnes calculées dans Excel.

## Prochaines étapes possibles

- Étendre `config/gamme_price_map.json` à d'autres catalogues/gammes.
- Générer un export PDF en plus du PPTX (nécessite LibreOffice installé pour
  la conversion en ligne de commande).
- Génération en lot (tous les clients actifs d'un coup).
