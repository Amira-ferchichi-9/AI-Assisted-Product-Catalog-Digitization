# generate_catalogs.py — Génération en lot, un PPTX par client

Outil en ligne de commande, pensé pour être branché sur une tâche planifiée :
il relit la base Excel à chaque exécution et régénère un catalogue PPTX par
client actif. Contrairement à `app.py` (Streamlit, un client à la fois, sur le
deck `Layla_Apero_Catalogue_FR.pptx`), ce script construit les slides de zéro
à partir des données — nécessaire car le nombre de produits par gamme varie.

## Ce qu'il fait

1. **Contrôles obligatoires** : lit la feuille `Contrôles` du classeur. S'il
   existe la moindre valeur non nulle en colonne C (anomalies), le script
   **s'arrête immédiatement** (code de sortie 1) et liste les anomalies —
   aucun catalogue n'est généré tant que le fichier n'est pas propre.
2. **Prix final par client** : pour chaque client actif et chaque produit
   actif, applique le prix négocié (`Relations`) si une ligne existe pour ce
   couple client × produit et que la date du jour est dans la plage
   `date_debut`/`date_fin` (bornes ouvertes si vides) ; sinon le prix
   standard (`Produits.prix_standard_kg`).
3. **Un PPTX par client** (`{code_client}.pptx`) : page de garde (logo client
   à côté du logo Layla si `logo_client_fichier` est renseigné, sinon logo
   Layla seul) puis une page par `gamme` (triée par `ordre_affichage`), avec
   une carte par produit : photo (ou encart "Photo à venir" si introuvable),
   nom, description, prix au kg + prix unitaire (déduit de `poids_net_kg`),
   allergènes si renseignés. Une gamme avec plus de 6 produits actifs se
   poursuit sur une slide "(suite)".
4. **Log** clair de chaque étape (clients traités, photos manquantes, etc.),
   avec code de sortie non nul en cas d'échec (fichier introuvable, contrôles
   en échec, erreur de génération pour un client).

La colonne description utilisée s'adapte automatiquement à la langue présente
dans le classeur (`description_fr`, `description_de`, ...) — le script
détecte et journalise la langue trouvée.

## Installation

```powershell
pip install python-pptx openpyxl pandas pillow
```

## Lancer

```powershell
python generate_catalogs.py --excel-path "C:\Users\amira\OneDrive\Layla_Base_Donnees_v1.3_groupes_seuls.xlsx" --output-dir "C:\Users\amira\Desktop\CATALOGUE DIGITALISEE\output\catalogues"
```

Le fichier OneDrive étant synchronisé localement, ce chemin reflète toujours
la dernière version enregistrée par le directeur commercial — aucune
récupération manuelle n'est nécessaire.

### Options utiles

| Option | Rôle |
|---|---|
| `--photos-dir` | Dossier des photos produits, nommées exactement comme `photo_fichier` (défaut : `photos/`). |
| `--logo-dir` | Dossier des logos clients (défaut : `logo/`). |
| `--catalogue "Apéritifs"` | Ne garder que les produits dont la colonne `catalogue` vaut cette valeur. |
| `--client MARRIOTT` | Ne générer que pour ce client. |
| `--date 2026-09-01` | Date de référence pour évaluer les prix négociés (défaut : aujourd'hui). |
| `-v` | Logs détaillés (DEBUG). |

### Pour une tâche planifiée

```powershell
python generate_catalogs.py --excel-path "C:\Users\amira\OneDrive\Layla_Base_Donnees_v1.3_groupes_seuls.xlsx" --output-dir "C:\...\output\catalogues"
if ($LASTEXITCODE -ne 0) { <# alerter : contrôles en échec ou erreur #> }
```

## État constaté sur le fichier réel actuel

Testé sur `Layla_Base_Donnees_v1.3_groupes_seuls.xlsx` (25/07/2026) :

- **Le fichier est aujourd'hui rejeté par les contrôles** : les 20 produits
  saisis n'ont aucun `photo_fichier` renseigné (`· sans photo : 20`). Tant que
  cette colonne n'est pas remplie, `generate_catalogs.py` refusera de générer
  quoi que ce soit — c'est le comportement voulu, pas un bug.
- Une fois les photos renseignées (et les fichiers correspondants déposés
  dans `--photos-dir`), le reste de la chaîne est validé de bout en bout :
  9 clients actifs, 3 gammes, prix standard corrects (aucune négociation
  saisie pour l'instant dans `Relations`).
- Les descriptions du fichier actuel sont en **allemand**
  (`description_de` / `argumentaire_de`) — le catalogue généré sera donc en
  allemand tant que cette colonne n'est pas remplacée/complétée par une
  version française.
- Remarque data : chaque parfum apparaît sur **deux lignes** `code_article`
  distinctes avec le même nom commercial (ex. deux lignes
  "Knusperkugel Mandel — Paprika"). Le script les traite comme deux produits
  séparés (donc deux cartes identiques) — à vérifier de votre côté si c'est
  intentionnel (deux formats/conditionnements ?) ou une double saisie.
