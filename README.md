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

- `Layla_Base_Donnees.xlsx` : base de données réelle (feuilles `Produits`,
  `Clients`, `Relations`). Actuellement, seule la ligne d'exemple de chaque feuille
  est remplie — voir la section **Limites actuelles** ci-dessous.
  **Important** : gardez toujours ce nom de fichier exact, au même endroit. Le
  script relit ce fichier à chaque génération (jamais de cache) : modifiez-le et
  sauvegardez (Ctrl+S) directement dedans, sans jamais faire "Enregistrer sous"
  vers une copie — sinon l'app continuera de lire l'ancienne version.

  **Base partagée via OneDrive** : si le fichier maître est en fait édité par
  le directeur commercial sur OneDrive (dossier synchronisé), installez le
  client OneDrive sur la machine qui exécute cette app, puis pointez dessus
  sans toucher au code, via une variable d'environnement :

  ```powershell
  $env:LAYLA_EXCEL_PATH = "C:\Users\<vous>\OneDrive\...\Layla_Base_Donnees.xlsx"
  streamlit run app.py
  ```

  Le fichier restera alors toujours à jour automatiquement (synchro OneDrive),
  sans téléchargement manuel ni copie périmée. Si l'app est un jour déployée
  sur un serveur sans client OneDrive installé, remplacer cette lecture directe
  par un appel Microsoft Graph API (`GET /drives/{id}/items/{id}/content`) est
  le chemin de migration naturel — même fonction `load_produits()` /
  `load_clients()` / `load_relations()` en aval, seule la source du fichier change.
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

Pour tester avec les données d'exemple plutôt que le fichier réel, activer le
bouton "Utiliser les données d'exemple" en haut de la page.

## Limites actuelles / À faire avant usage en production

1. **La base Excel réelle est vide** (structure + 1 ligne d'exemple par feuille,
   comme indiqué dans l'onglet `Mode d'emploi`). Pour que la génération à partir
   du fichier réel fonctionne, il faut :
   - Ajouter dans `Produits` une ligne par `code_article` référencé dans
     `config/gamme_price_map.json` (`PC-AMANDE-500`, `PC-CACAHUETE-500`,
     `PC-SESAME-COURGE-500` par défaut — à renommer si vos vrais codes articles
     sont différents, en gardant `config/gamme_price_map.json` synchronisé).
   - Remplir `Clients` avec vos vrais clients, en vérifiant que
     `logo_client_fichier` correspond exactement à un nom de fichier présent
     dans `logo/` (ex. la ligne d'exemple actuelle indique `marriott.png`,
     alors que le fichier réel s'appelle `Logo Marriott Tunis .png`).
   - Remplir `Relations` pour les clients ayant un prix négocié.
2. **Seuls 3 emplacements de prix sont automatisés** (le prix "TND/KG" en tête
   des slides 4, 5 et 6). Les prix comparatifs de la slide 7 et le "prix par
   unité" dans les blocs "Conditionnement" restent fixes — leur logique de
   calcul n'était pas évidente à déduire du template et mériterait d'être
   confirmée avant automatisation.
3. Le générateur ne touche à rien d'autre que le logo de couverture et ces
   3 zones de prix : toute autre personnalisation (textes, images produit,
   ajout/suppression de slides) est hors périmètre de ce MVP.

## Prochaines étapes possibles

- Étendre `config/gamme_price_map.json` à d'autres catalogues/gammes.
- Générer un export PDF en plus du PPTX (nécessite LibreOffice installé pour
  la conversion en ligne de commande).
- Génération en lot (tous les clients actifs d'un coup).
