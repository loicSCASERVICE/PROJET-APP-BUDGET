#!/usr/bin/env python3
"""
split_releve_annuel.py
------------------------
Découpe un relevé bancaire annuel (export CSV type Crédit Agricole, contenant
les opérations de plusieurs mois) en un fichier CSV distinct par mois, dans
le même format que les exports mensuels habituels.

Chaque fichier généré peut ensuite être traité normalement par
budget_processor.py, un mois à la fois.

Usage:
    python3 split_releve_annuel.py RELEVE_ANNUEL.csv [--output-dir dossier_sortie]

Par défaut, les fichiers sont créés dans le même dossier que le relevé
d'origine, nommés "Releve_2026-01.csv", "Releve_2026-02.csv", etc.
"""

import argparse
import calendar
import csv
import re
import sys
from collections import defaultdict
from pathlib import Path

# Réutilise exactement la même logique de lecture que budget_processor.py
# pour rester cohérent avec le format des relevés (encodage, en-tête, dates).
sys.path.insert(0, str(Path(__file__).resolve().parent))
from budget_processor import normalize, parse_date  # noqa: E402


def find_header_row(lines: list[str]) -> int:
    for i, line in enumerate(lines):
        first_cell = line.split(";")[0].strip().strip('"')
        if normalize(first_cell) == "date":
            return i
    raise ValueError(
        "Impossible de trouver la ligne d'en-tête du tableau (colonne 'Date'). "
        "Vérifie le format du fichier."
    )


def read_raw(csv_path: Path) -> str:
    for enc in ("utf-8", "utf-8-sig", "cp1252"):
        try:
            return csv_path.read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
    raise ValueError(f"Impossible de lire l'encodage du fichier {csv_path}")


def split_statement(csv_path: Path, output_dir: Path) -> list[Path]:
    raw_text = read_raw(csv_path)
    lines = raw_text.splitlines()
    header_idx = find_header_row(lines)

    # Tout ce qui précède la ligne d'en-tête (identité, solde, période...) :
    # on le recopie tel quel dans chaque fichier mensuel généré.
    preamble_lines = lines[:header_idx]

    # Le tableau lui-même est reparsé avec csv.reader (et pas ligne à ligne)
    # car certains libellés contiennent des retours à la ligne à l'intérieur
    # de guillemets ("Paiement par carte\nX3088 ...").
    table_text = "\n".join(lines[header_idx:])
    reader = csv.reader(table_text.splitlines(), delimiter=";")
    rows = list(reader)
    header_row = rows[0]
    header_norm = [normalize(h) for h in header_row]

    def col_index(*candidates):
        for cand in candidates:
            for i, h in enumerate(header_norm):
                if cand in h:
                    return i
        return None

    idx_date = col_index("date")
    if idx_date is None:
        raise ValueError("Colonne 'Date' introuvable dans le relevé.")

    # Regroupe les lignes de données par mois (AAAA-MM), dans l'ordre
    # d'apparition d'origine.
    rows_by_month: dict[str, list[list[str]]] = defaultdict(list)
    skipped = 0
    for row in rows[1:]:
        if not row or not "".join(row).strip():
            continue
        date_val = row[idx_date] if idx_date < len(row) else ""
        date = parse_date(date_val)
        if date is None:
            skipped += 1
            continue
        month_key = date.strftime("%Y-%m")
        rows_by_month[month_key].append(row)

    if skipped:
        print(f"⚠️  {skipped} ligne(s) sans date exploitable ont été ignorées.")

    if not rows_by_month:
        raise ValueError("Aucune transaction datée trouvée dans le fichier.")

    output_dir.mkdir(parents=True, exist_ok=True)
    written_files = []
    periode_pattern = re.compile(
        r"(entre le )\d{2}/\d{2}/\d{4}( et le )\d{2}/\d{2}/\d{4}"
    )

    for month_key in sorted(rows_by_month.keys()):
        year, month = (int(p) for p in month_key.split("-"))
        last_day = calendar.monthrange(year, month)[1]
        month_preamble = []
        for line in preamble_lines:
            # La ligne "Liste des opérations du compte entre le J1/M1/AAAA et
            # le J2/M2/AAAA" du relevé d'origine ne concerne que le mois
            # source : on la met à jour avec les vraies bornes de CE mois,
            # sinon budget_processor.py déduirait le mauvais mois pour tous
            # les fichiers sauf le premier (il fait confiance à cette ligne
            # en priorité).
            if periode_pattern.search(line):
                line = periode_pattern.sub(
                    rf"\g<1>01/{month:02d}/{year}\g<2>{last_day:02d}/{month:02d}/{year}",
                    line,
                )
            month_preamble.append(line)

        out_path = output_dir / f"Releve_{month_key}.csv"
        with open(out_path, "w", encoding="utf-8", newline="") as f:
            for line in month_preamble:
                f.write(line + "\n")
            writer = csv.writer(f, delimiter=";", quoting=csv.QUOTE_MINIMAL)
            writer.writerow(header_row)
            for row in rows_by_month[month_key]:
                writer.writerow(row)
        written_files.append(out_path)
        print(f"✅ {out_path.name} — {len(rows_by_month[month_key])} transaction(s)")

    return written_files


def main():
    parser = argparse.ArgumentParser(description="Découpe un relevé annuel en fichiers CSV mensuels.")
    parser.add_argument("releve", type=Path, help="Chemin du relevé CSV annuel à découper")
    parser.add_argument("--output-dir", type=Path, default=None,
                         help="Dossier de sortie (par défaut : même dossier que le relevé)")
    args = parser.parse_args()

    if not args.releve.exists():
        print(f"❌ Fichier introuvable : {args.releve}", file=sys.stderr)
        sys.exit(1)

    output_dir = args.output_dir or args.releve.parent
    written = split_statement(args.releve, output_dir)

    print(f"\n🎉 {len(written)} fichier(s) mensuel(s) créé(s) dans : {output_dir}")
    print("   Tu peux maintenant traiter chacun normalement avec budget_processor.py")
    print("   (ou via l'icône Lanceur_Budget), un mois à la fois.")


if __name__ == "__main__":
    main()
