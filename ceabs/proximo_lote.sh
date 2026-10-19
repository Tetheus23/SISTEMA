#!/bin/sh
# Coleta os .bin baixados e lista os próximos N ids pendentes do inventário (exclui os acima de 10 MB).
cd /home/user/SISTEMA/entregas/ceabs || exit 1
python3 -I /home/user/SISTEMA/ceabs/coletar_blobs.py /root/.claude/projects/-home-user-SISTEMA/2e1127f2-f29e-5cba-8378-3e2e880063f4/tool-results emails/raw drive/inventario_final.tsv
F=$(mktemp); ls emails/raw | sed 's/\.msg$//' > $F; cut -d' ' -f1 drive/grandes.txt >> $F
grep -vxF -f $F drive/ids_ordem.txt | head -${1:-30} | tr '\n' ' '; rm -f $F
