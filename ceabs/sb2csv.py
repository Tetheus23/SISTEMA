"""Converte o resultado salvo de uma consulta do Supabase MCP (JSON entre marcadores untrusted-data) em CSV.

Uso: python3 -I sb2csv.py <resultado.txt> <saida.csv>
"""
import csv
import json
import re
import sys

txt = open(sys.argv[1], encoding='utf-8').read()
try:  # o arquivo salvo pode ser o JSON da ferramenta ({"result": "..."}) ou o texto puro
    txt = json.loads(txt)['result']
except (ValueError, KeyError, TypeError):
    pass
m = re.search(r'<untrusted-data-[^>]+>\s*(\[.*\])\s*</untrusted-data', txt, re.S)
linhas = json.loads(m[1])
with open(sys.argv[2], 'w', newline='', encoding='utf-8') as f:
    w = csv.DictWriter(f, fieldnames=list(linhas[0].keys()))
    w.writeheader()
    w.writerows(linhas)
print(f'{len(linhas)} linhas -> {sys.argv[2]}')
