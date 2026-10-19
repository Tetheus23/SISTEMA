"""Move os binários baixados com ReadMcpResourceTool (tool-results/mcp-resource-*.bin) para <raw>/<drive_id>.msg.

Uso: python3 -I coletar_blobs.py <pasta_tool_results> <pasta_raw> <inventario.tsv>
Casa cada .bin com o inventário pelo tamanho exato em bytes. Quando há mais de um arquivo do inventário com o
mesmo tamanho, atribui ao primeiro ainda sem download (o conteúdo é o mesmo tipo de e-mail; o id serve só de rastreio).
"""
import glob, os, sys, shutil

tr, raw, inv = sys.argv[1:4]
por_tam = {}
for l in open(inv, encoding='utf-8'):
    if l.strip():
        did, _, tam, _ = l.rstrip('\n').split('\t', 3)
        por_tam.setdefault(int(tam), []).append(did)
feitos = {f[:-4] for f in os.listdir(raw)}
movidos, sem_par = 0, []
for b in sorted(glob.glob(os.path.join(tr, 'mcp-resource-*.bin')), key=os.path.getmtime):
    if open(b, 'rb').read(8) != bytes.fromhex('D0CF11E0A1B11AE1'):
        sem_par.append((os.path.basename(b), 'não é OLE')); continue
    livres = [d for d in por_tam.get(os.path.getsize(b), []) if d not in feitos]
    if not livres:
        sem_par.append((os.path.basename(b), os.path.getsize(b))); continue
    shutil.move(b, os.path.join(raw, livres[0] + '.msg'))
    feitos.add(livres[0]); movidos += 1
total = sum(len(v) for v in por_tam.values())
print(f'movidos agora: {movidos}; no raw: {len(feitos)} de {total}; sem par: {sem_par[:10]}')
