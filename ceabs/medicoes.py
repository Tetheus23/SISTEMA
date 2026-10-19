"""Normaliza todas as planilhas de medição anexadas aos e-mails (norm_medicao.py) em estrutura/medicoes/<sha1>.json.

Uso: python3 -I medicoes.py <pasta_ceabs>
Considera qualquer .xlsx/.xlsm anexado que tenha o cabeçalho de medição da CEABS (REFERÊNCIA + itens).
"""
import csv
import glob
import json
import os
import subprocess
import sys

BASE = sys.argv[1]
EST = os.path.join(BASE, 'estrutura')
NORM = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'norm_medicao.py')
os.makedirs(os.path.join(EST, 'medicoes'), exist_ok=True)
for f in glob.glob(os.path.join(EST, 'medicoes', '*.json')):
    os.remove(f)
for r in csv.DictReader(open(os.path.join(EST, 'anexos_unicos.tsv'), encoding='utf-8'), delimiter='\t'):
    if r['ext'] not in ('.xlsx', '.xlsm'):
        continue
    out = subprocess.run([sys.executable, '-I', NORM, os.path.join(BASE, r['arquivo'])], capture_output=True, text=True)
    if out.returncode:
        continue
    d = json.loads(out.stdout)
    if not d['ref'] or not d['itens']:
        continue
    d.update(arquivo=r['nome'], email_data=r['data'], sha1=r['sha1'], de=r['de'])
    json.dump(d, open(os.path.join(EST, 'medicoes', r['sha1'][:10] + '.json'), 'w'), ensure_ascii=False)
    tot = round(sum(i['valor'] for i in d['itens']), 2)
    print(f"{d['ref']:15} {str(d['cliente'])[:24]:24} {tot:9.2f} {len(d['itens']):4}  {r['nome'][:45]} ({r['data'][:10]})")
