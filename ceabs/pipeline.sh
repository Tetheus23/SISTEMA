#!/bin/sh
# Refaz a auditoria CEABS a partir dos e-mails já extraídos (emails/emails.jsonl e anexos em emails/proc).
# Uso: ceabs/pipeline.sh entregas/ceabs
set -e
D=$(dirname "$0")
python3 -I "$D/medicoes.py" "$1" > /dev/null
python3 -I "$D/estruturar.py" "$1"
python3 -I "$D/auditar.py" "$1"
python3 -I "$D/faturas.py" "$1"
