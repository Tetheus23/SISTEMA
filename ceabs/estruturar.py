"""Estrutura os dados da auditoria CEABS a partir dos e-mails já extraídos e das medições normalizadas.

Uso: python3 -I estruturar.py <pasta_ceabs>   (ex.: entregas/ceabs)

Lê:
  emails/emails.jsonl            (extrair_emails.py)
  estrutura/medicoes/*.json      (norm_medicao.py sobre cada PlanilhaMedicao anexada)
  dados/2026-07.json             (medição de jul/26 recebida direto no chat)
Grava em estrutura/:
  itens_cobrados.csv   uma linha por item de medição (mês de referência x código CEABS)
  veiculos.csv         cadastro CEABS por chassi (placa mais recente, instalação, cancelamento, retirada)
  pedidos.csv          e-mails da Grillo que pedem cancelamento/retirada, uma linha por placa citada
  eventos_ceabs.csv    e-mails da CEABS que citam placas (retirada, cobrança de equipamento, visita etc.)
"""
import csv
import glob
import json
import os
import re
import sys

BASE = sys.argv[1]
EST = os.path.join(BASE, 'estrutura')
MESES = {m: i for i, m in enumerate(['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto',
                                       'setembro', 'outubro', 'novembro', 'dezembro'], 1)}
PLACA = re.compile(r'\b([A-Z]{3})[- ]?(\d[A-Z0-9]\d{2})\b')
CHASSI = re.compile(r'\b([A-HJ-NPR-Z0-9]{16,17})\b')
PEDIDO = re.compile(r'cancel|retir|rescis|desativ|devolu|remo[cç]', re.I)


def mes_ref(ref):
    nome, ano = ref.lower().split('/')
    return f'{int(ano):04d}-{MESES[nome.strip()]:02d}'


def data_br(s):
    """'03/07/2026 10:09:21' ou '03/07/2026' -> '2026-07-03'."""
    m = re.match(r'(\d{2})/(\d{2})/(\d{4})', str(s or ''))
    return f'{m[3]}-{m[2]}-{m[1]}' if m else ''


def escreve(nome, linhas, campos):
    with open(os.path.join(EST, nome), 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=campos, extrasaction='ignore')
        w.writeheader()
        w.writerows(linhas)


# ---------- medições ----------
medicoes = {}
for p in glob.glob(os.path.join(EST, 'medicoes', '*.json')) + [os.path.join(BASE, 'dados', '2026-07.json')]:
    d = json.load(open(p, encoding='utf-8'))
    if not d.get('ref') or 'GRILLO' not in (d.get('cliente') or 'LOCADORA GRILLO').upper():
        continue
    if 'grillo' in (d.get('de') or ''):  # cópia editada internamente e reencaminhada: vale a original da CEABS
        continue
    mes = mes_ref(d['ref'])
    # a mesma referência pode vir em mais de um e-mail: fica a enviada por último pela CEABS
    if mes not in medicoes or d.get('email_data', '9') > medicoes[mes].get('email_data', ''):
        d['origem'] = d.get('arquivo') or os.path.basename(p)
        medicoes[mes] = d


def anexo(sha_prefixo):
    for r in csv.DictReader(open(os.path.join(EST, 'anexos_unicos.tsv'), encoding='utf-8'), delimiter='\t'):
        if r['sha1'].startswith(sha_prefixo):
            r['arquivo'] = os.path.join(BASE, r['arquivo'])  # caminhos da tabela são relativos à pasta da entrega
            return r
    raise SystemExit(f'anexo {sha_prefixo} não encontrado')


# jun/26 não veio como PlanilhaMedicao: os itens estão na aba DETALHAMENTO do "CONTROLE CEABS.xlsx"
# que o Monitoramento enviou em 01/09/2026 15:38 (cópia da medição, com vencimento e data de pagamento)
import openpyxl  # noqa: E402
CONTROLE = '44ef0b'
ws = openpyxl.load_workbook(anexo(CONTROLE)['arquivo'], data_only=True, read_only=True)['DETALHAMENTO']
cab, *linhas_ctrl = list(ws.iter_rows(values_only=True))
col = {c: i for i, c in enumerate(cab) if c}
for mes in ('2026-06',):
    if mes in medicoes:
        continue
    its = []
    for x in linhas_ctrl:
        if x[0] and str(x[0])[:7] == mes:
            visita = x[col['Monitoramento']] == 45  # a cópia interna mistura a visita improdutiva (R$ 45) com a mensalidade
            its.append(dict(aba='02-Visita Improdutiva' if visita else '01-Principal', codigo=x[col['Código']],
                            chassi=x[col['Chassi']], placa=x[col['Placa']],
                            produto='Visita Improdutiva' if visita else 'Monitoramento GSM',
                            mensal=x[col['Monitoramento']], prorata=x[col['Pró-Rata']] or 0,
                            valor=round(x[col['Valor a Faturar']] or 0, 2), instalacao=x[col['Data de Instalação']],
                            cancelamento=None, retirada=None))
    medicoes[mes] = dict(ref='Junho/2026', itens=its, origem='CONTROLE CEABS.xlsx (Monitoramento, 01/09/2026) - aba DETALHAMENTO')

itens, veic = [], {}
for mes, d in sorted(medicoes.items()):
    for i in d['itens']:
        tipo = i['aba'].split('-', 1)[-1].strip()
        chassi = (i['chassi'] or '').strip().upper()
        itens.append(dict(mes_uso=mes, ref_ceabs=d['ref'], tipo=tipo, produto=i['produto'], codigo=i['codigo'],
                          chassi=chassi, placa=(i['placa'] or '').replace('-', '').strip().upper(),
                          mensal=i['mensal'], prorata=i['prorata'], valor=i['valor'],
                          instalacao=data_br(i['instalacao']), cancelamento=data_br(i['cancelamento']),
                          retirada=data_br(i['retirada']), arquivo=d['origem']))
        v = veic.setdefault(chassi, dict(chassi=chassi, placa='', codigos=set(), instalacao='', cancelamento='',
                                         retirada='', primeiro_mes=mes, ultimo_mes=mes, meses=0, total=0.0))
        if i['placa']:
            v['placa'] = i['placa'].replace('-', '').upper()
        v['codigos'].add(i['codigo'])
        for k in ('instalacao', 'cancelamento', 'retirada'):
            if data_br(i[k]):
                v[k] = data_br(i[k])
        v['ultimo_mes'] = max(v['ultimo_mes'], mes)
        v['total'] += i['valor']
        if tipo == 'Principal':
            v['meses'] += 1
for v in veic.values():
    v['codigos'] = ' '.join(sorted(v['codigos']))
    v['total'] = round(v['total'], 2)
escreve('itens_cobrados.csv', itens, list(itens[0].keys()))
escreve('veiculos.csv', sorted(veic.values(), key=lambda v: v['placa']), list(next(iter(veic.values())).keys()))
placa_chassi = {v['placa']: c for c, v in veic.items() if v['placa']}
chassis = set(veic)


def chassi_de(txt):
    """Chassis citados no texto; aceita o truncado em 16 caracteres que aparece nos e-mails."""
    out = set()
    for c in CHASSI.findall(txt.upper()):
        if c in chassis:
            out.add(c)
        elif len(c) == 16:
            out.update(x for x in chassis if x.endswith(c))
    return out


# ---------- e-mails ----------
NFSE = re.compile(r'Número da NFS-e\s+(\d+).*?Data do RPS\s+(\d{2})/(\d{2})/(\d{2}).*?Valor Total\s+([\d.]+)', re.S)
FATURA = re.compile(r'Data de Emissão:\s*(\d{1,2}) de (\w+) de (\d{4}).*?Nº\.\s+Vencimento\s+Valor\s+(\d+)\s+(\d{2})\.(\d{2})\.(\d{4})\s+([\d.,]+)'
                    r'.*?DESCRIÇÃO\s+Valor\s+(.+?)\s{2,}([\d.,]+)', re.S)
pedidos, eventos, docs = [], [], {}
for l in open(os.path.join(BASE, 'emails', 'emails.jsonl'), encoding='utf-8'):
    r = json.loads(l)
    m = NFSE.search(r['corpo_topo'])
    if m and (r.get('assunto') or '').startswith('RPS'):
        docs[('NFS-e', m[1])] = dict(tipo_doc='NFS-e', numero=m[1], emissao=f'20{m[4]}-{m[3]}-{m[2]}', vencimento='',
                                     valor=float(m[5]), descricao='', fonte=f'e-mail Mastersaf {r["data"][:10]}')
    for a in r['anexos']:  # "Fatura de Serviços" (nota de débito, sem retenção) em PDF
        if a['tipo'] == 'arquivo' and a.get('txt') and a['nome'].lower().endswith('.pdf'):
            t = open(os.path.join(BASE, a['txt']), encoding='utf-8', errors='replace').read()
            f = FATURA.search(t)
            if f and 'Fatura de Serviços' in t:
                val = lambda s: float(s.replace('.', '').replace(',', '.'))
                docs[('Fatura', f[4])] = dict(tipo_doc='Fatura de Serviços', numero=f[4],
                                              emissao=f'{f[3]}-{MESES[f[2].lower()]:02d}-{int(f[1]):02d}',
                                              vencimento=f'{f[7]}-{f[6]}-{f[5]}', valor=val(f[8]),
                                              descricao=f[9].strip()[:60], fonte=f'anexo {a["nome"]} ({r["data"][:10]})')
    de = (r.get('de_email') or '').lower()
    texto = r['corpo_topo']
    assunto = r.get('assunto') or ''
    placas = {a + b for a, b in PLACA.findall(texto.upper())}
    citados = {placa_chassi.get(p, '') or p for p in placas} | chassi_de(texto)
    if not citados:
        continue
    para = ((r.get('para') or '') + ' ' + (r.get('cc') or '')).lower()
    base = dict(data=r['data'][:10], hora=r['data'][11:16], remetente=de, assunto=assunto[:120], drive_id=r['drive_id'])
    for c in sorted(citados):
        reg = dict(base, chassi=c if c in chassis else '', placa=veic[c]['placa'] if c in chassis else c)
        if 'grillo' in de and (PEDIDO.search(assunto) or PEDIDO.search(texto[:2000])):
            reg['para_ceabs'] = 'ceabs' in para
            pedidos.append(reg)
        elif 'ceabs' in de:
            eventos.append(reg)
campos = ['data', 'hora', 'placa', 'chassi', 'remetente', 'para_ceabs', 'assunto', 'drive_id']
escreve('pedidos.csv', sorted(pedidos, key=lambda x: (x['data'], x['placa'])), campos)
escreve('eventos_ceabs.csv', sorted(eventos, key=lambda x: (x['data'], x['placa'])), campos)
docs = sorted(docs.values(), key=lambda d: (d['emissao'], d['numero']))
escreve('documentos_fiscais.csv', docs, list(docs[0].keys()))
print(f'medições: {len(medicoes)} ({min(medicoes)} a {max(medicoes)}); itens: {len(itens)}; veículos: {len(veic)}; '
      f'pedidos: {len(pedidos)}; eventos CEABS: {len(eventos)}')
print('meses:', ' '.join(sorted(medicoes)))
