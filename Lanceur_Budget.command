#!/bin/bash
# Lanceur_Budget.command
# ------------------------
# Double-clic sur ce fichier depuis le Bureau (ou n'importe où) :
# ouvre une fenetre pour choisir le releve CSV du mois, puis lance
# le traitement (categorisation dans le navigateur, mise a jour de
# l'Excel).

# Se place dans le dossier ou se trouve ce script (dossier du projet)
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$DIR"

# Fenetre de selection de fichier native macOS (AppleScript)
CSV_PATH=$(osascript -e 'set csvFile to choose file with prompt "Choisis le releve bancaire CSV du mois :" of type {"csv", "public.comma-separated-values-text"}' -e 'POSIX path of csvFile' 2>/dev/null)

if [ -z "$CSV_PATH" ]; then
    echo "Aucun fichier selectionne. Fermeture."
    read -p "Appuie sur Entree pour fermer cette fenetre..."
    exit 0
fi

echo "Fichier selectionne : $CSV_PATH"
echo ""

python3 budget_processor.py "$CSV_PATH"

echo ""
read -p "Termine. Appuie sur Entree pour fermer cette fenetre..."
