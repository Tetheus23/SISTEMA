"""Extrai em lote os e-mails .msg baixados do Drive, usando o motor da skill leitor-de-msg.

Uso: python3 -I extrair_emails.py <pasta_raw> <pasta_proc> <inventario.tsv> <saida.jsonl> [--paralelo N]

Para cada <pasta_raw>/<drive_id>.msg:
  - roda analisar_msg.py da skill (corpo, anexos com texto, e-mails aninhados, RELATORIO.md) em <pasta_proc>/<drive_id>/;
  - lê a data de envio/entrega das propriedades MAPI quando o cabeçalho de transporte não traz a data;
  - separa a mensagem do topo (antes do primeiro "De:/Enviada em:" do histórico);
  - lista os anexos (recursivo) com hash, para deduplicar planilhas, NFs e boletos repetidos nas respostas.
Grava uma linha JSON por e-mail em <saida.jsonl>.
"""
import concurrent.futures as cf
import datetime as dt
import email.utils
import hashlib
import json
import os
import re
import struct
import subprocess
import sys

SKILL = '/root/.claude/skills/synced/5a8032e2-0024-488e-860b-eda044e14ab3_8ee8bbb8-8f60-4362-aec8-062049b47f1f/leitor-de-msg/scripts'
sys.path.insert(0, SKILL)
import msg_extract as mx  # noqa: E402

# início do histórico citado (Outlook em português e inglês)
CITACAO = re.compile(r'^\s*(De|From):\s.+\n(\s*(Enviad[ao] em|Sent|Data|Date):)', re.M | re.I)
CITACAO2 = re.compile(r'^-{3,}\s*(Mensagem original|Original Message)', re.M | re.I)


def filetime(v):
    return dt.datetime(1601, 1, 1) + dt.timedelta(microseconds=v // 10)


def datas_mapi(path):
    """Lê PR_CLIENT_SUBMIT_TIME (0x0039), PR_MESSAGE_DELIVERY_TIME (0x0E06) e PR_CREATION_TIME (0x3007)."""
    cfb = mx.CFB(open(path, 'rb').read())
    idx = next(i for i in cfb.children(0) if cfb.entries[i]['name'] == '__properties_version1.0')
    raw = cfb.read(idx)
    out = {}
    for off in range(32, len(raw) - 15, 16):
        tag, = struct.unpack_from('<I', raw, off)
        ptype, pid = tag & 0xFFFF, tag >> 16
        if ptype == 0x0040 and pid in (0x0039, 0x0E06, 0x3007):
            v, = struct.unpack_from('<Q', raw, off + 8)
            if v:
                out[{0x0039: 'envio', 0x0E06: 'entrega', 0x3007: 'criacao'}[pid]] = filetime(v).isoformat(timespec='seconds')
    return out


def anexos_rec(info, nivel=0):
    for a in info.get('anexos', []):
        if 'email_anexado' in a:
            yield {'nome': a['nome'], 'tipo': 'email', 'nivel': nivel, 'assunto': a['email_anexado'].get('assunto'),
                   'data': a['email_anexado'].get('data'), 'pasta': a['pasta']}
            yield from anexos_rec(a['email_anexado'], nivel + 1)
        else:
            p = a['arquivo']
            h = hashlib.sha1(open(p, 'rb').read()).hexdigest() if os.path.exists(p) else None
            txt = p + '.txt' if os.path.exists(p + '.txt') else None
            yield {'nome': a['nome'], 'tipo': 'arquivo', 'nivel': nivel, 'arquivo': p, 'tamanho': a.get('tamanho'),
                   'sha1': h, 'txt': txt, 'imagem_corpo': a.get('embutido_no_corpo', False)}


def processar(raw_path, proc_dir, titulo):
    did = os.path.basename(raw_path)[:-4]
    out = os.path.join(proc_dir, did)
    if not os.path.exists(os.path.join(out, 'RELATORIO.md')):
        subprocess.run([sys.executable, '-I', os.path.join(SKILL, 'analisar_msg.py'), raw_path, out],
                       capture_output=True, timeout=600)
    rec = {'drive_id': did, 'titulo_drive': titulo}
    try:
        info = json.load(open(os.path.join(out, 'resumo.json'), encoding='utf-8'))
    except Exception as e:  # .msg que não abriu
        rec['erro'] = f'não abriu: {e}'
        return rec
    corpo = open(os.path.join(out, 'corpo.txt'), encoding='utf-8').read()
    m = [x.start() for x in (CITACAO.search(corpo), CITACAO2.search(corpo)) if x]
    topo = corpo[:min(m)] if m else corpo
    try:
        mapi = datas_mapi(raw_path)
    except Exception:
        mapi = {}
    data = None
    if info.get('data'):
        try:
            data = email.utils.parsedate_to_datetime(info['data']).astimezone(dt.timezone(dt.timedelta(hours=-3))).replace(tzinfo=None).isoformat(timespec='seconds')
        except Exception:
            data = None
    if not data:  # horário MAPI é UTC; converte para Brasília
        base = mapi.get('envio') or mapi.get('entrega') or mapi.get('criacao')
        if base:
            data = (dt.datetime.fromisoformat(base) - dt.timedelta(hours=3)).isoformat(timespec='seconds')
    rec.update(assunto=info.get('assunto'), de_nome=info.get('de_nome'), de_email=info.get('de_email'),
               para=info.get('para'), cc=info.get('cc'), data=data, datas_mapi=mapi,
               corpo_topo=topo.strip()[:6000], corpo_len=len(corpo),
               anexos=list(anexos_rec(info)),
               pendencias='PENDÊNCIAS DE LEITURA (tratar' in open(os.path.join(out, 'RELATORIO.md'), encoding='utf-8').read())
    return rec


def main():
    raw_dir, proc_dir, inv, saida = sys.argv[1:5]
    par = int(sys.argv[sys.argv.index('--paralelo') + 1]) if '--paralelo' in sys.argv else 4
    titulos = {l.split('\t')[0]: l.rstrip('\n').split('\t')[3] for l in open(inv, encoding='utf-8') if l.strip()}
    arqs = sorted(os.path.join(raw_dir, f) for f in os.listdir(raw_dir) if f.endswith('.msg'))
    os.makedirs(proc_dir, exist_ok=True)
    with cf.ThreadPoolExecutor(par) as ex, open(saida, 'w', encoding='utf-8') as fo:
        for rec in ex.map(lambda p: processar(p, proc_dir, titulos.get(os.path.basename(p)[:-4])), arqs):
            fo.write(json.dumps(rec, ensure_ascii=False) + '\n')
    faltam = set(titulos) - {os.path.basename(p)[:-4] for p in arqs}
    print(f'{len(arqs)} e-mails processados; {len(faltam)} do inventário sem arquivo .msg')


if __name__ == '__main__':
    main()
