#!/usr/bin/env python3
"""
budget_processor.py
--------------------
Traite un relevé bancaire mensuel (export CSV type Crédit Agricole) et met à jour
un fichier Excel récapitulatif (un mois par colonne, une catégorie par ligne).

Usage:
    python3 budget_processor.py RELEVE.csv [--month 2026-08] [--output suivi_budget.xlsx] [--config categories_config.json]

Si --month est omis, le script tente de le déduire automatiquement du contenu du
relevé (ligne "Liste des opérations du compte entre le JJ/MM/AAAA et le JJ/MM/AAAA").

Par défaut, le script ouvre une page dans ton navigateur listant toutes les
transactions avec une catégorie suggérée pré-sélectionnée ⭐ (basée sur les
mots-clés déjà mémorisés) : ajuste ce qu'il faut et clique sur "Valider" pour
générer le fichier Excel. Options :
  --cli             catégorisation dans le terminal, ligne par ligne (sans navigateur)
  --no-interactive  aucune question posée (mode batch, mots-clés connus uniquement)
"""

import argparse
import csv
import json
import re
import shutil
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

UNCATEGORIZED_LABEL = "À catégoriser"
TOTAL_LABEL = "TOTAL / SOLDE"
BUDGETS_SHEET = "Budgets"

HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TOTAL_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
FLAG_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
BUDGET_WARN_FILL = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
CURRENCY_FMT = '#,##0.00 €;-#,##0.00 €;-'
MAX_BUDGET_ROWS = 500  # bornes des plages de mise en forme conditionnelle (large marge)


# --------------------------------------------------------------------------
# Utilitaires texte / config
# --------------------------------------------------------------------------

def normalize(text: str) -> str:
    """Minuscules, sans accents, espaces normalisés."""
    if text is None:
        return ""
    text = str(text).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in text if not unicodedata.combining(c))
    text = re.sub(r"\s+", " ", text)
    return text


def load_config(path: Path) -> dict:
    with open(path, encoding="utf-8") as f:
        cfg = json.load(f)
    cfg["_bank_map_norm"] = {normalize(k): v for k, v in cfg.get("bank_category_mapping", {}).items()}
    cfg["_keywords_norm"] = {
        cat: [normalize(kw) for kw in kws] for cat, kws in cfg.get("keywords", {}).items()
    }
    return cfg


def save_config(cfg: dict, path: Path) -> None:
    """Réécrit le fichier de config (uniquement les clés publiques, pas les index internes)."""
    ordered = {}
    for key in ("_comment", "categories", "bank_category_mapping", "keywords"):
        if key in cfg:
            ordered[key] = cfg[key]
    with open(path, "w", encoding="utf-8") as f:
        json.dump(ordered, f, ensure_ascii=False, indent=2)
        f.write("\n")


# Préfixes génériques de libellés bancaires à ignorer lors de l'extraction
# d'un mot-clé à mémoriser (voir extract_signature).
GENERIC_PREFIXES = [
    "paiement par carte",
    "prelevement",
    "virement emis",
    "virement en votre faveur",
    "virement instantane",
    "retrait",
]


def extract_signature(libelle: str) -> str:
    """Extrait un mot-clé réutilisable d'un libellé (enlève préfixes génériques,
    numéros de carte masqués, dates, longues suites de chiffres)."""
    norm = normalize(libelle)
    for prefix in GENERIC_PREFIXES:
        norm = norm.replace(prefix, " ")
    norm = re.sub(r"\bx\d{3,}\b", " ", norm)
    norm = re.sub(r"\b\d{1,2}/\d{1,2}(?:/\d{2,4})?\b", " ", norm)
    norm = re.sub(r"\b\d{5,}\b", " ", norm)
    norm = re.sub(r"\s+", " ", norm).strip()
    words = [w for w in norm.split(" ") if len(w) > 2]
    return " ".join(words[:4])


# --------------------------------------------------------------------------
# Lecture et parsing du relevé CSV
# --------------------------------------------------------------------------

def parse_amount(value: str) -> float:
    """Convertit un montant au format FR ('1 004,91' ou '1\u202f004,91') en float."""
    if value is None:
        return 0.0
    value = str(value).strip()
    if not value:
        return 0.0
    # supprime tous les espaces (classiques, insécables, insécables fines)
    value = re.sub(r"[\s\u00a0\u202f]", "", value)
    value = value.replace(",", ".")
    try:
        return float(value)
    except ValueError:
        return 0.0


def parse_date(value: str):
    for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d"):
        try:
            return datetime.strptime(value.strip(), fmt)
        except (ValueError, AttributeError):
            continue
    return None


def find_header_row(lines: list[str]) -> int:
    """Trouve l'index de la ligne d'en-tête réelle du tableau (colonne 'Date')."""
    for i, line in enumerate(lines):
        first_cell = line.split(";")[0].strip().strip('"')
        if normalize(first_cell) == "date":
            return i
    raise ValueError(
        "Impossible de trouver la ligne d'en-tête du tableau (colonne 'Date'). "
        "Vérifie le format du fichier."
    )


def detect_month_from_content(raw_text: str) -> str | None:
    m = re.search(r"entre le \d{2}/\d{2}/(\d{4}) et le (\d{2})/(\d{2})/\d{4}", raw_text)
    if m:
        year, _, month = m.groups()
        return f"{year}-{month}"
    return None


def read_statement(csv_path: Path) -> tuple[list[dict], str | None]:
    """Lit le CSV bancaire et retourne (liste de transactions, mois détecté ou None).

    Chaque transaction: {"date": datetime, "libelle": str, "montant": float, "categorie_banque": str}
    montant > 0 = recette, montant < 0 = dépense.
    """
    # Essaye utf-8, puis cp1252 en repli (exports bancaires parfois en Windows-1252)
    raw_text = None
    for enc in ("utf-8", "utf-8-sig", "cp1252"):
        try:
            raw_text = csv_path.read_text(encoding=enc)
            break
        except UnicodeDecodeError:
            continue
    if raw_text is None:
        raise ValueError(f"Impossible de lire l'encodage du fichier {csv_path}")

    lines = raw_text.splitlines()
    header_idx = find_header_row(lines)
    detected_month = detect_month_from_content(raw_text)

    table_text = "\n".join(lines[header_idx:])
    reader = csv.reader(table_text.splitlines(), delimiter=";")
    rows = list(reader)
    header = [normalize(h) for h in rows[0]]

    def col_index(*candidates):
        for cand in candidates:
            for i, h in enumerate(header):
                if cand in h:
                    return i
        return None

    idx_date = col_index("date")
    idx_libelle = col_index("libelle")
    idx_debit = col_index("debit")
    idx_credit = col_index("credit")
    idx_categorie = col_index("categorie")

    if idx_date is None or idx_libelle is None:
        raise ValueError("Colonnes 'Date' et/ou 'Libellé' introuvables dans le relevé.")

    transactions = []
    for row in rows[1:]:
        if not row or not "".join(row).strip():
            continue
        # protège contre des lignes plus courtes que prévu
        def get(i):
            return row[i] if i is not None and i < len(row) else ""

        date = parse_date(get(idx_date))
        libelle = get(idx_libelle).replace("\n", " ").strip()
        debit = parse_amount(get(idx_debit))
        credit = parse_amount(get(idx_credit))
        montant = credit - debit
        categorie_banque = get(idx_categorie).strip()

        if date is None and not libelle:
            continue  # ligne vide / parasite

        transactions.append({
            "date": date,
            "libelle": libelle,
            "montant": montant,
            "categorie_banque": categorie_banque,
        })

    return transactions, detected_month


# --------------------------------------------------------------------------
# Catégorisation
# --------------------------------------------------------------------------

def suggest_category(tx: dict, cfg: dict) -> str | None:
    """Propose une catégorie (mapping banque puis mots-clés). Ne retourne rien
    de définitif : c'est une suggestion soumise à confirmation de l'utilisateur."""
    categories = cfg["categories"]
    bank_map = cfg["_bank_map_norm"]
    keywords = cfg["_keywords_norm"]

    # 1. Catégorie déjà fournie par la banque -> mapping vers catégorie cible
    banque_norm = normalize(tx["categorie_banque"])
    if banque_norm:
        if banque_norm in bank_map:
            return bank_map[banque_norm]
        # peut-être que la banque utilise déjà exactement un nom de catégorie cible
        for cat in categories:
            if normalize(cat) == banque_norm:
                return cat

    # 2. Mots-clés sur le libellé (dont ceux mémorisés lors des sessions précédentes)
    libelle_norm = normalize(tx["libelle"])
    for cat in categories:
        for kw in keywords.get(cat, []):
            if kw and kw in libelle_norm:
                return cat

    # 3. Rien trouvé
    return None


# --------------------------------------------------------------------------
# Sauvegarde de sécurité
# --------------------------------------------------------------------------

def backup_existing_output(output_path: Path) -> Path | None:
    """Copie le classeur existant dans un sous-dossier backups/ avant de le
    réécrire, pour ne jamais perdre l'historique en cas de problème pendant
    la génération. Retourne le chemin de la copie, ou None si le fichier
    n'existait pas encore."""
    if not output_path.exists():
        return None
    backups_dir = output_path.parent / "backups"
    backups_dir.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = backups_dir / f"{output_path.stem}_{timestamp}{output_path.suffix}"
    shutil.copy2(output_path, backup_path)
    return backup_path


# --------------------------------------------------------------------------
# Écriture du fichier Excel maître
# --------------------------------------------------------------------------

def month_label(month_key: str) -> str:
    """'2026-08' -> 'Août 2026'"""
    months_fr = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet",
                 "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
    try:
        year, mo = month_key.split("-")
        return f"{months_fr[int(mo) - 1]} {year}"
    except Exception:
        return month_key


def write_detail_sheet(wb: Workbook, month_key: str, transactions: list[dict], categorized: list[str]):
    sheet_name = month_key  # ex: "2026-08" (unique, trie bien alphabétiquement)
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)

    headers = ["Date", "Libellé", "Montant", "Catégorie"]
    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL

    for row_idx, (tx, cat) in enumerate(zip(transactions, categorized), start=2):
        ws.cell(row=row_idx, column=1, value=tx["date"].strftime("%d/%m/%Y") if tx["date"] else "")
        ws.cell(row=row_idx, column=2, value=tx["libelle"])
        c_montant = ws.cell(row=row_idx, column=3, value=round(tx["montant"], 2))
        c_montant.number_format = CURRENCY_FMT
        c_cat = ws.cell(row=row_idx, column=4, value=cat)
        if cat == UNCATEGORIZED_LABEL:
            c_cat.fill = FLAG_FILL

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 55
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 22
    ws.freeze_panes = "A2"
    return sheet_name


def ensure_budgets_sheet(wb: Workbook, cfg: dict) -> None:
    """Crée (une seule fois) la feuille 'Budgets', où tu peux saisir directement
    dans Excel un budget mensuel prévisionnel par catégorie. Les valeurs déjà
    présentes ne sont jamais écrasées lors des exécutions suivantes : seules
    les catégories qui n'y figurent pas encore sont ajoutées en bas."""
    budgets_seed = cfg.get("budgets", {})

    if BUDGETS_SHEET not in wb.sheetnames:
        ws = wb.create_sheet(BUDGETS_SHEET)
        for col, h in enumerate(["Catégorie", "Budget mensuel"], start=1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
        ws.column_dimensions["A"].width = 26
        ws.column_dimensions["B"].width = 16
        for r, cat in enumerate(cfg["categories"], start=2):
            ws.cell(row=r, column=1, value=cat)
            cell = ws.cell(row=r, column=2, value=budgets_seed.get(cat))
            cell.number_format = CURRENCY_FMT
        return

    ws = wb[BUDGETS_SHEET]
    existing = {ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)}
    next_row = ws.max_row + 1
    for cat in cfg["categories"]:
        if cat in existing:
            continue
        ws.cell(row=next_row, column=1, value=cat)
        cell = ws.cell(row=next_row, column=2, value=budgets_seed.get(cat))
        cell.number_format = CURRENCY_FMT
        next_row += 1


def update_recap_sheet(wb: Workbook, cfg: dict, month_key: str, sheet_name: str):
    recap_name = "Récapitulatif"
    if recap_name not in wb.sheetnames:
        ws = wb.create_sheet(recap_name, 0)
        ws.cell(row=1, column=1, value="Catégorie").font = HEADER_FONT
        ws.cell(row=1, column=1).fill = HEADER_FILL
        all_cats = cfg["categories"] + [UNCATEGORIZED_LABEL, TOTAL_LABEL]
        for r, cat in enumerate(all_cats, start=2):
            cell = ws.cell(row=r, column=1, value=cat)
            if cat == TOTAL_LABEL:
                cell.font = Font(bold=True)
                cell.fill = TOTAL_FILL
        ws.column_dimensions["A"].width = 26
    else:
        ws = wb[recap_name]

    # Repère la ligne de chaque catégorie
    cat_row = {}
    max_row = ws.max_row
    for r in range(2, max_row + 1):
        val = ws.cell(row=r, column=1).value
        if val:
            cat_row[val] = r

    # Ajoute une ligne pour toute catégorie de la config absente du récapitulatif
    # (ex: catégorie créée à la volée dans l'interface web), juste au-dessus de
    # "À catégoriser" / "TOTAL / SOLDE" pour que le mois en cours l'inclue déjà.
    missing_cats = [c for c in cfg["categories"] if c not in cat_row]
    if missing_cats:
        insert_before_row = cat_row.get(UNCATEGORIZED_LABEL) or cat_row.get(TOTAL_LABEL) or (ws.max_row + 1)
        ws.insert_rows(insert_before_row, amount=len(missing_cats))
        for offset, cat in enumerate(missing_cats):
            ws.cell(row=insert_before_row + offset, column=1, value=cat)
        # Recalcule toutes les lignes après insertion (les positions ont bougé)
        cat_row = {}
        max_row = ws.max_row
        for r in range(2, max_row + 1):
            val = ws.cell(row=r, column=1).value
            if val:
                cat_row[val] = r

    # Trouve ou crée la colonne pour ce mois (repérée via une ligne cachée d'ID mois en ligne 1 -> on stocke le month_key dans un commentaire de cellule, plus simple: on relit l'en-tête affiché "Mois Année" et on garde une correspondance via une ligne technique en dernière ligne)
    # Astuce simple: on stocke le month_key brut dans la ligne 0 (juste au-dessus, non visible) -> openpyxl n'a pas de ligne 0.
    # On utilise donc une ligne technique tout en bas ("_month_key") pour retrouver la colonne d'un mois donné lors des relances.
    key_row_label = "_month_key"
    key_row = cat_row.get(key_row_label)
    if key_row is None:
        key_row = max_row + 1
        ws.cell(row=key_row, column=1, value=key_row_label).font = Font(italic=True, size=8, color="AAAAAA")
        cat_row[key_row_label] = key_row

    month_col = None
    max_col = ws.max_column
    for c in range(2, max_col + 1):
        if ws.cell(row=key_row, column=c).value == month_key:
            month_col = c
            break

    if month_col is None:
        month_col = max_col + 1 if max_col >= 2 else 2
        ws.cell(row=1, column=month_col, value=month_label(month_key)).font = HEADER_FONT
        ws.cell(row=1, column=month_col).fill = HEADER_FILL
        ws.column_dimensions[get_column_letter(month_col)].width = 16
        ws.cell(row=key_row, column=month_col, value=month_key)

        # Met en évidence les catégories qui dépassent leur budget mensuel
        # (feuille "Budgets"), quand un budget a été saisi pour cette catégorie.
        new_col_letter = get_column_letter(month_col)
        warn_range = f"{new_col_letter}2:{new_col_letter}{MAX_BUDGET_ROWS}"
        warn_formula = (
            f'=AND($A2<>"",ISNUMBER(VLOOKUP($A2,{BUDGETS_SHEET}!$A:$B,2,FALSE)),'
            f'ABS({new_col_letter}2)>VLOOKUP($A2,{BUDGETS_SHEET}!$A:$B,2,FALSE))'
        )
        ws.conditional_formatting.add(warn_range, FormulaRule(formula=[warn_formula], fill=BUDGET_WARN_FILL))

    col_letter = get_column_letter(month_col)

    # Formules SUMIF par catégorie, référencées à la feuille détail du mois
    for cat, r in cat_row.items():
        if cat in (key_row_label, TOTAL_LABEL):
            continue
        formula = f"=SUMIF('{sheet_name}'!D:D,A{r},'{sheet_name}'!C:C)"
        cell = ws.cell(row=r, column=month_col, value=formula)
        cell.number_format = CURRENCY_FMT

    total_row = cat_row.get(TOTAL_LABEL)
    if total_row:
        first_cat_row = 2
        last_cat_row = total_row - 1
        formula = f"=SUM({col_letter}{first_cat_row}:{col_letter}{last_cat_row})"
        cell = ws.cell(row=total_row, column=month_col, value=formula)
        cell.number_format = CURRENCY_FMT
        cell.font = Font(bold=True)
        cell.fill = TOTAL_FILL

    # Masquer la ligne technique _month_key
    ws.row_dimensions[key_row].hidden = True

    return recap_name


def reorder_recap_columns(wb: Workbook):
    """Trie les colonnes mensuelles du récapitulatif par ordre chronologique."""
    ws = wb["Récapitulatif"]
    key_row = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(row=r, column=1).value == "_month_key":
            key_row = r
            break
    if key_row is None:
        return

    max_col = ws.max_column
    month_cols = []
    for c in range(2, max_col + 1):
        key = ws.cell(row=key_row, column=c).value
        if key:
            month_cols.append((key, c))
    month_cols.sort()  # tri chronologique sur "YYYY-MM"

    if [c for _, c in month_cols] == list(range(2, 2 + len(month_cols))):
        return  # déjà trié, rien à faire

    # Copie les valeurs/formules dans le bon ordre vers des colonnes temporaires au-delà, puis retasse
    snapshot = {}
    for r in range(1, ws.max_row + 1):
        snapshot[r] = {c: (ws.cell(row=r, column=c).value, ws.cell(row=r, column=c).number_format)
                       for _, c in month_cols}

    for new_idx, (key, old_col) in enumerate(month_cols, start=2):
        for r in range(1, ws.max_row + 1):
            val, fmt = snapshot[r][old_col]
            cell = ws.cell(row=r, column=new_idx, value=val)
            cell.number_format = fmt
    # Remets les styles d'en-tête / largeur de colonnes
    for new_idx, _ in enumerate(month_cols, start=2):
        ws.cell(row=1, column=new_idx).font = HEADER_FONT
        ws.cell(row=1, column=new_idx).fill = HEADER_FILL
        ws.column_dimensions[get_column_letter(new_idx)].width = 16


# --------------------------------------------------------------------------
# Résolution interactive des catégories manquantes
# --------------------------------------------------------------------------

def prompt_for_category(tx: dict, categories: list[str], suggestion: str | None, index: int, total: int) -> str | None:
    """Affiche la transaction et demande à l'utilisateur de confirmer/choisir une
    catégorie. Retourne le nom de la catégorie choisie, ou None si laissée non
    catégorisée. Appuyer sur Entrée seul accepte la suggestion si elle existe."""
    date_str = tx["date"].strftime("%d/%m/%Y") if tx["date"] else "?"
    print("\n" + "-" * 70)
    print(f"[{index}/{total}]")
    print(f"  Date    : {date_str}")
    print(f"  Libellé : {tx['libelle']}")
    print(f"  Montant : {tx['montant']:.2f} €")
    print("\nCatégories :")
    for i, cat in enumerate(categories, start=1):
        marker = "  ⭐" if cat == suggestion else "    "
        print(f"{marker} {i:>2}. {cat}")
    print("     0. Laisser 'À catégoriser' (à corriger plus tard dans Excel)")

    if suggestion:
        prompt_txt = f"Numéro [Entrée = {suggestion}] > "
    else:
        prompt_txt = "Numéro > "

    while True:
        try:
            choice = input(prompt_txt).strip()
        except EOFError:
            return suggestion  # plus d'entrée possible (ex: script non-interactif) -> on garde la suggestion si dispo
        if choice == "":
            return suggestion  # accepte la suggestion (peut être None)
        if choice == "0":
            return None
        if choice.isdigit() and 1 <= int(choice) <= len(categories):
            return categories[int(choice) - 1]
        print("  ⚠️  Entrée invalide, réessaie.")


def categorize_interactively(transactions: list[dict], cfg: dict) -> tuple[list[str], bool]:
    """Demande à l'utilisateur de confirmer/choisir la catégorie de CHAQUE
    transaction (avec suggestion pré-remplie quand une correspondance existe).
    Retourne (liste des catégories, True si la config a été enrichie)."""
    categories = cfg["categories"]
    categorized = []
    config_changed = False
    total = len(transactions)

    for i, tx in enumerate(transactions, start=1):
        suggestion = suggest_category(tx, cfg)
        chosen = prompt_for_category(tx, categories, suggestion, i, total)

        if chosen is None:
            categorized.append(UNCATEGORIZED_LABEL)
            continue

        categorized.append(chosen)

        # Si le choix diffère de la suggestion (ou qu'il n'y avait pas de
        # suggestion), on propose de mémoriser un mot-clé pour la prochaine fois.
        if chosen != suggestion:
            try:
                remember = input(f"   Mémoriser ce type de transaction sous '{chosen}' pour la prochaine fois ? (o/N) ").strip().lower()
            except EOFError:
                remember = "n"
            if remember == "o":
                signature = extract_signature(tx["libelle"])
                if signature:
                    cfg.setdefault("keywords", {}).setdefault(chosen, [])
                    if signature not in cfg["keywords"][chosen]:
                        cfg["keywords"][chosen].append(signature)
                        cfg["_keywords_norm"].setdefault(chosen, []).append(signature)
                        config_changed = True
                        print(f"   → mot-clé '{signature}' associé à '{chosen}' pour les prochains relevés.")

    return categorized, config_changed


# --------------------------------------------------------------------------
# Programme principal
# --------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Traite un relevé bancaire mensuel et met à jour le récapitulatif Excel.")
    parser.add_argument("csv_file", type=Path, help="Chemin du relevé CSV bancaire du mois")
    parser.add_argument("--month", type=str, default=None, help="Mois au format AAAA-MM (déduit automatiquement si omis)")
    parser.add_argument("--output", type=Path, default=Path("suivi_budget.xlsx"), help="Fichier Excel récapitulatif (créé s'il n'existe pas)")
    parser.add_argument("--config", type=Path, default=Path(__file__).parent / "categories_config.json", help="Fichier de configuration des catégories")
    parser.add_argument("--no-interactive", action="store_true",
                         help="Ne pas demander confirmation pour chaque ligne (mode batch : catégorisation automatique via mots-clés uniquement)")
    parser.add_argument("--cli", action="store_true",
                         help="Catégoriser dans le terminal ligne par ligne au lieu d'ouvrir le navigateur")
    args = parser.parse_args()

    cfg = load_config(args.config)
    transactions, detected_month = read_statement(args.csv_file)

    if not transactions:
        print("⚠️  Aucune transaction trouvée dans le fichier.", file=sys.stderr)
        sys.exit(1)

    month_key = args.month or detected_month
    if not month_key:
        # Repli: mois majoritaire des dates de transactions
        dated = [tx["date"] for tx in transactions if tx["date"]]
        if dated:
            from collections import Counter
            counts = Counter(d.strftime("%Y-%m") for d in dated)
            month_key = counts.most_common(1)[0][0]
        else:
            print("❌ Impossible de déterminer le mois. Utilise --month AAAA-MM.", file=sys.stderr)
            sys.exit(1)

    if args.no_interactive:
        categorized = []
        for tx in transactions:
            cat = suggest_category(tx, cfg)
            categorized.append(cat if cat else UNCATEGORIZED_LABEL)
    elif args.cli:
        print(f"\n🔎 {len(transactions)} transaction(s) à passer en revue.")
        categorized, config_changed = categorize_interactively(transactions, cfg)
        if config_changed:
            save_config(cfg, args.config)
            print(f"\n   ✓ Config mise à jour : {args.config}")
    else:
        from webui import run_web_categorizer
        categorized, config_changed = run_web_categorizer(
            transactions, cfg, month_label(month_key), suggest_category, extract_signature
        )
        if config_changed:
            save_config(cfg, args.config)
            print(f"   ✓ Config mise à jour : {args.config}")

    # Sauvegarde de sécurité du classeur existant avant toute modification
    backup_path = backup_existing_output(args.output)

    # Charge ou crée le classeur maître
    if args.output.exists():
        wb = load_workbook(args.output)
    else:
        wb = Workbook()
        wb.remove(wb.active)

    ensure_budgets_sheet(wb, cfg)
    sheet_name = write_detail_sheet(wb, month_key, transactions, categorized)
    update_recap_sheet(wb, cfg, month_key, sheet_name)
    reorder_recap_columns(wb)

    # Rétablit "Récapitulatif" en premier onglet
    wb.move_sheet("Récapitulatif", offset=-len(wb.sheetnames))

    wb.save(args.output)

    # ---------------- Résumé console ----------------
    total_transactions = len(transactions)
    n_uncat = sum(1 for c in categorized if c == UNCATEGORIZED_LABEL)
    total_montant = sum(tx["montant"] for tx in transactions)

    print(f"\n✅ {total_transactions} transactions traitées pour {month_label(month_key)}")
    print(f"   Solde du mois : {total_montant:,.2f} €".replace(",", " ").replace(".", ","))
    print(f"   Fichier mis à jour : {args.output}")
    if backup_path:
        print(f"   💾 Sauvegarde de l'ancienne version : {backup_path}")

    if n_uncat:
        print(f"\n⚠️  {n_uncat} transaction(s) non catégorisée(s) — à corriger directement "
              f"dans l'onglet '{sheet_name}' (colonne Catégorie, en rouge) :")
        for tx, cat in zip(transactions, categorized):
            if cat == UNCATEGORIZED_LABEL:
                d = tx["date"].strftime("%d/%m/%Y") if tx["date"] else "?"
                print(f"   - {d} | {tx['libelle'][:60]:<60} | {tx['montant']:>8.2f} €")
        print("\n   💡 Astuce : édite la colonne 'Catégorie' dans cet onglet, puis "
              "ouvre le fichier dans Excel (les totaux du Récapitulatif se recalculent seuls).")
    else:
        print("   Toutes les transactions ont été catégorisées automatiquement 🎉")


if __name__ == "__main__":
    main()
