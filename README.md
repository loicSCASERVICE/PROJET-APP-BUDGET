# Suivi de budget — traitement des relevés mensuels

## Ce que fait le script

`budget_processor.py` lit un relevé bancaire mensuel (export CSV) et met à jour
un fichier Excel **`suivi_budget.xlsx`** avec :

- un onglet **"Récapitulatif"** : une ligne par catégorie, une colonne par mois
  (formules qui se recalculent automatiquement)
- un onglet **par mois** (ex: `2026-08`) avec le détail de chaque transaction
  et sa catégorie assignée

## Installation

```bash
pip install openpyxl flask
```

## Utilisation

```bash
python3 budget_processor.py Relevé_AOUT.csv
```

Par défaut, une page s'ouvre automatiquement dans ton navigateur : un tableau
liste toutes les transactions du mois, chacune avec un menu déroulant de
catégorie. Les lignes marquées **★** ont déjà une suggestion pré-sélectionnée
(basée sur un mot-clé mémorisé) — vérifie-la et corrige si besoin, sinon
choisis la bonne catégorie. Un seul bouton **"Valider et générer le fichier"**
en bas de page termine le traitement ; tu peux ensuite fermer l'onglet du
navigateur, le script continue tout seul dans le terminal.

Le mois est déduit automatiquement du relevé. Tu peux aussi le préciser :

```bash
python3 budget_processor.py Relevé_AOUT.csv --month 2026-08 --output suivi_budget.xlsx
```

Relance la commande chaque mois avec le nouveau relevé : une nouvelle colonne
s'ajoute au récapitulatif (si tu relances pour un mois déjà traité, sa colonne
est simplement mise à jour).

## Autres modes disponibles

```bash
python3 budget_processor.py Relevé_AOUT.csv --cli
```
Catégorisation dans le terminal, transaction par transaction (sans navigateur).

```bash
python3 budget_processor.py Relevé_AOUT.csv --no-interactive
```
Aucune question posée : catégorisation automatique via les mots-clés déjà
connus uniquement (mode batch/script). Les transactions non reconnues restent
"À catégoriser" (à corriger directement dans Excel).

## Découper un relevé annuel en fichiers mensuels

Si ta banque t'a fourni un seul export CSV couvrant plusieurs mois (ex: un
relevé annuel) plutôt qu'un fichier par mois, découpe-le d'abord :

```bash
python3 split_releve_annuel.py Releve_Annuel.csv
```

Cela génère un fichier `Releve_2026-01.csv`, `Releve_2026-02.csv`, etc. (un
par mois trouvé dans le fichier), dans le même dossier que l'original. Utilise
`--output-dir un_dossier` pour choisir un autre emplacement.

Traite ensuite chaque fichier mensuel normalement, l'un après l'autre, avec
`budget_processor.py` (ou via l'icône Lanceur_Budget) — exactement comme si
c'étaient des exports mensuels classiques.

## Comment la catégorisation fonctionne

1. **Suggestion automatique** : si un mot-clé déjà mémorisé correspond au
   libellé de la transaction (ou si le relevé fournit lui-même une catégorie),
   elle est pré-sélectionnée (★ dans le navigateur, "Entrée = ..." en mode
   `--cli`).
2. **Ton choix fait foi** : tu valides ou corriges chaque ligne.
3. Si tu coches "Mémoriser mes corrections" (navigateur) ou réponds "o" (mode
   `--cli`), un mot-clé est extrait automatiquement du libellé et ajouté à
   `categories_config.json` : les transactions similaires seront reconnues
   toutes seules dès le prochain relevé.

## Corriger une transaction après coup

Si une ligne est restée "À catégoriser" (surlignée en rouge), ouvre
`suivi_budget.xlsx`, va dans l'onglet du mois concerné, et modifie directement
la cellule de la colonne **Catégorie**. Les totaux du "Récapitulatif" utilisent
des formules `SUMIF` : ils se recalculent tout seuls à l'ouverture du fichier
dans Excel.

## Ajuster les catégories ou les règles de mapping

Tout se trouve dans `categories_config.json` :
- `categories` : la liste et l'ordre des catégories affichées
- `bank_category_mapping` : correspondance "catégorie banque" → "catégorie cible"
  (utile si ton export bancaire fournit déjà une catégorie)
- `keywords` : mots-clés (sans accents, en minuscules) utilisés pour suggérer
  une catégorie automatiquement — enrichi au fil de l'eau quand tu coches
  "Mémoriser"

Aucune modification du script n'est nécessaire pour ajouter une catégorie ou
un mot-clé.
