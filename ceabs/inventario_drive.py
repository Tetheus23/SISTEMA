"""Consolida as listagens do Drive (resultados de search_files salvos em tool-results) num inventário único.

Uso: python3 -I inventario_drive.py <pasta_tool_results> <folder_id> <saida.tsv>
"""
import glob, json, os, sys

pasta, folder, saida = sys.argv[1:4]
itens = {}
for p in glob.glob(os.path.join(pasta, 'mcp-Google_Drive-search_files-*.txt')):
    try:
        d = json.load(open(p))
    except Exception:
        continue
    if isinstance(d, list):  # formato [{text: json}]
        d = json.loads(d[0]['text'])
    for f in d.get('files', []):
        if f.get('parentId') == folder:
            itens[f['id']] = f
with open(saida, 'w') as out:
    for f in sorted(itens.values(), key=lambda f: f['createdTime']):
        out.write('\t'.join([f['id'], f['createdTime'], f.get('fileSize', ''), f['title']]) + '\n')
print(len(itens), 'arquivos;', 'de', min(f['createdTime'] for f in itens.values()), 'a', max(f['createdTime'] for f in itens.values()))
