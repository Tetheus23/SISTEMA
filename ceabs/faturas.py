"""Monta a tabela de faturas CEABS (uma linha por documento) com mês de uso, vencimento, pagamento e conferência fiscal.

Uso: python3 -I faturas.py <pasta_ceabs>   (depois de estruturar.py)

Fontes:
  estrutura/documentos_fiscais.csv  NFS-e (avisos Mastersaf) e Faturas de Serviços em PDF, 2023 a 2025
  CONTROLE PAGAMENTOS CEABS.xlsx    controle do Matheus (vencimentos jan/25 a jan/26, valor e data pagos)
  CONTROLE CEABS.xlsx (aba DIN)     controle do Monitoramento de 2026 (NF, vencimento, valor pago e data)
  estrutura/nfs2026.json e dados/notas.json   NFS-e de 2026 lidas dos PDFs (valores e retenções)
  e-mails de 30/09, 11/10, 13/10 e 17/10/2026  situação das faturas em aberto
Grava estrutura/faturas.csv.
"""
import csv
import datetime as dt
import os
import sys
import warnings
from collections import defaultdict

import openpyxl

warnings.filterwarnings('ignore')
BASE = sys.argv[1]
EST = os.path.join(BASE, 'estrutura')
HOJE = dt.date(2026, 10, 19)


def mes_menos(iso, n=1):
    a, m = int(iso[:4]), int(iso[5:7]) - n
    while m < 1:
        a, m = a - 1, m + 12
    return f'{a:04d}-{m:02d}'


def venc_padrao(mes_uso):
    """Padrão histórico: vencimento no dia 3 do segundo mês após o mês de uso."""
    a, m = int(mes_uso[:4]), int(mes_uso[5:7]) + 2
    if m > 12:
        a, m = a + 1, m - 12
    return dt.date(a, m, 3)


def anexo(nome, data):
    for r in csv.DictReader(open(os.path.join(EST, 'anexos_unicos.tsv'), encoding='utf-8'), delimiter='\t'):
        if r['nome'] == nome and r['data'].startswith(data):
            return os.path.join(BASE, r['arquivo'])


medicao = defaultdict(float)
for r in csv.DictReader(open(os.path.join(EST, 'itens_cobrados.csv'), encoding='utf-8')):
    medicao[r['mes_uso']] += float(r['valor'])

# ---------- pagamentos 2025 (controle do Matheus) ----------
pag, ctrl = {}, []
grupo = None
ws = openpyxl.load_workbook(anexo('CONTROLE PAGAMENTOS CEABS.xlsx', '2026-02-11'), data_only=True).active
for nota, venc, valor, pago, data_pg, obs, *_ in ws.iter_rows(min_row=2, values_only=True):
    if not venc:
        continue
    if pago:
        grupo = dict(total=round(pago, 2), data=data_pg.date() if data_pg else None)
    reg = dict(nota=str(nota or ''), venc=venc.date(), liquido=valor, grupo=grupo, obs=obs or '')
    ctrl.append(reg)
    if nota:
        pag[str(nota)] = reg

linhas = []


def add(mes_uso, tipo, doc, emissao, venc, bruto, liquido, pago_em, situacao, obs='', fiscal='', fonte=''):
    vp = venc_padrao(mes_uso)
    dias = (venc - vp).days if venc else None
    alerta = ''
    if dias is not None and dias < -3:
        alerta = f'Vencimento antecipado ({-dias} dias antes do padrão {vp:%d/%m/%Y})'
    elif dias is not None and dias > 10:
        alerta = f'Vencimento posterior ao padrão ({vp:%d/%m/%Y}): CEABS emitiu com atraso'
    atraso = (pago_em - venc).days if pago_em and venc else ((HOJE - venc).days if venc and not pago_em else None)
    linhas.append(dict(mes_uso=mes_uso, mes_emissao=emissao.strftime('%Y-%m') if emissao else '', tipo_cobranca=tipo,
                       documento=doc, emissao=emissao, vencimento=venc, vencimento_padrao=vp, alerta_vencimento=alerta,
                       valor_bruto=round(bruto, 2), valor_liquido=round(liquido, 2) if liquido is not None else '',
                       pago_em=pago_em or '', situacao=situacao,
                       dias_atraso=atraso if atraso and atraso > 0 else 0, conferencia_fiscal=fiscal, observacao=obs, fonte=fonte))


usados = {'29909'}  # no controle a visita de nov/24 aparece pelo nº do RPS (29909) = NFS-e 31870
TIPO = {'MONITORAMENTO GSM': 'Mensalidade', 'MULTA': 'Taxa de rescisão'}
for r in csv.DictReader(open(os.path.join(EST, 'documentos_fiscais.csv'), encoding='utf-8')):
    em = dt.date.fromisoformat(r['emissao'])
    if em < dt.date(2024, 9, 1) or em >= dt.date(2026, 1, 1):
        continue
    usados.add(r['numero'])
    mes_uso = mes_menos(r['emissao'][:7])
    valor = float(r['valor'])
    p = pag.get(r['numero'])
    tipo = TIPO.get(r['descricao'], 'Mensalidade')
    if r['tipo_doc'] == 'NFS-e' and valor in (45.0, 135.0):
        tipo = 'Visita improdutiva'
    if r['numero'] == '32928':  # emitida em 11/07/2025 fora do ciclo; não consta em nenhuma planilha de medição
        tipo = 'Visita improdutiva (avulsa, fora da medição)'
    venc = dt.date.fromisoformat(r['vencimento']) if r['vencimento'] else (p['venc'] if p else None)
    if not venc:  # NFS-e sem vencimento no aviso: mesmo vencimento da fatura emitida no mesmo dia
        venc = next((dt.date.fromisoformat(x['vencimento']) for x in csv.DictReader(open(os.path.join(EST, 'documentos_fiscais.csv')))
                     if x['emissao'] == r['emissao'] and x['vencimento']), em + dt.timedelta(days=30))
    if p:
        sit = 'Pago' if p['grupo'] and p['grupo']['data'] else 'Sem registro de pagamento'
        pago_em, liq, obs = (p['grupo'] or {}).get('data'), p['liquido'], p['obs']
    else:
        sit, pago_em, liq, obs = 'Sem registro de pagamento', None, None, ''
    fiscal = 'Fatura de Serviços (nota de débito, sem retenção na fonte)' if r['tipo_doc'] != 'NFS-e' else ''
    add(mes_uso, tipo, f'{r["tipo_doc"]} {r["numero"]}', em, venc, valor, liq, pago_em, sit, obs, fiscal, r['fonte'])

# documentos do controle de pagamentos que não vieram nos e-mails
ORIGINAL = {'1400038707': 2272.60}  # fatura com desconto de R$ 217,36 (imposto pago a maior em 04/11/2025)
for c in ctrl:
    if c['nota'] and c['nota'] in usados:
        continue
    mes_uso = mes_menos(c['venc'].strftime('%Y-%m'), 2)
    if mes_uso < '2024-08' or mes_uso > '2025-11':
        continue
    tipo = 'Taxa de rescisão' if 'retirada' in c['obs'] else ('Visita improdutiva' if 'visita' in c['obs'] else 'Mensalidade')
    doc = f'Fatura {c["nota"]}' if c['nota'] else 'Fatura de Serviços (nº não informado no controle)'
    add(mes_uso, tipo, doc, None, c['venc'], ORIGINAL.get(c['nota'], c['liquido']), c['liquido'],
        (c['grupo'] or {}).get('data'), 'Pago' if c['grupo'] and c['grupo']['data'] else 'Sem registro de pagamento',
        c['obs'] or 'Documento não localizado nos e-mails; dados do controle de pagamentos', '', 'CONTROLE PAGAMENTOS CEABS.xlsx')

# ---------- 2026 (controle do Monitoramento + PDFs das NFS-e) ----------
D = dt.date
SEM_RET = 'NFS-e sem retenção destacada (IRRF/PIS/COFINS/CSLL); a Grillo reteve 5,65% no pagamento: conferir recolhimento do DARF'
IRRF15 = 'IRRF de 15% (alíquota de multa/indenização) em NFS-e de serviço, sem PIS/COFINS/CSLL: mesma divergência da NF 8979'
END = 'Endereço do tomador desatualizado (Estrada do Pedregoso / Camboatá / Carlos Seidl)'
docs26 = [
    # mes_uso, tipo, doc, emissão, vencimento, bruto, líquido, pago_em, situação, obs, fiscal
    ('2025-12', 'Mensalidade', 'NFS-e 895 (monitoramento)', D(2026, 1, 22), D(2026, 2, 26), 2272.60, None, D(2026, 2, 27), 'Pago', 'Pago junto com a 896: R$ 5.865,96', SEM_RET),
    ('2025-12', 'Mensalidade', 'NFS-e 896 (gestão de frotas)', D(2026, 1, 22), D(2026, 2, 26), 3944.65, None, D(2026, 2, 27), 'Pago', 'Pago junto com a 895: R$ 5.865,96', SEM_RET),
    ('2025-12', 'Visita improdutiva', 'NFS-e 894', D(2026, 1, 22), D(2026, 2, 26), 45.00, 45.00, None, 'Sem registro de pagamento', '', ''),
    ('2025-12', 'Taxa de rescisão', 'NFS-e 1219', D(2026, 1, 22), D(2026, 2, 26), 180.00, 180.00, None, 'Sem registro de pagamento', '3 equipamentos x R$ 60', 'Taxa de rescisão sem retenção (na 2529 e na 8979 a CEABS reteve IRRF 15%)'),
    ('2026-01', 'Mensalidade', 'NFS-e 2386 (monitoramento)', D(2026, 2, 20), D(2026, 3, 26), 2272.60, None, D(2026, 6, 3), 'Pago', 'Pago junto com a 2261: R$ 5.787,78', SEM_RET),
    ('2026-01', 'Mensalidade', 'NFS-e 2261 (gestão de frotas)', D(2026, 2, 20), D(2026, 3, 26), 3861.78, None, D(2026, 6, 3), 'Pago', 'Pago junto com a 2386: R$ 5.787,78', SEM_RET),
    ('2026-01', 'Visita improdutiva', 'NFS-e 2324', D(2026, 2, 20), D(2026, 3, 26), 90.00, 90.00, None, 'Contestado (em aberto)', '2 visitas; contestada pela Grillo. Monitoramento pediu o detalhamento em 17/10/2026', ''),
    ('2026-01', 'Taxa de rescisão', 'NFS-e 2529', D(2026, 2, 20), D(2026, 3, 26), 240.00, 204.00, D(2026, 6, 3), 'Pago', '4 retiradas x R$ 60', IRRF15),
    ('2026-02', 'Mensalidade', 'NFS-e 3652', D(2026, 3, 18), D(2026, 4, 26), 5944.20, 5608.35, D(2026, 4, 2), 'Pago', '', 'OK: IRRF 1% e PIS/COFINS/CSLL 4,65%'),
    ('2026-03', 'Mensalidade', 'NFS-e 4981', D(2026, 4, 22), D(2026, 5, 5), 5944.20, 5608.35, D(2026, 6, 3), 'Pago', '', 'OK: IRRF 1% e PIS/COFINS/CSLL 4,65%'),
    ('2026-04', 'Mensalidade', 'NFS-e 5887 (monitoramento)', D(2026, 5, 5), D(2026, 6, 5), 2272.60, 2144.19, D(2026, 6, 3), 'Pago', '', 'OK: IRRF 1% e PIS/COFINS/CSLL 4,65%'),
    ('2026-04', 'Mensalidade', 'NFS-e 5889 (gestão de frotas)', D(2026, 5, 5), D(2026, 6, 5), 3671.60, 3464.14, D(2026, 6, 3), 'Pago', '', 'OK: IRRF 1% e PIS/COFINS/CSLL 4,65%'),
    ('2026-05', 'Mensalidade', 'NFS-e 6575 (monitoramento)', D(2026, 6, 3), D(2026, 7, 3), 2272.60, 2144.19, D(2026, 7, 3), 'Pago', '', f'OK nas retenções. {END}'),
    ('2026-05', 'Mensalidade', 'NFS-e 6576 (gestão de frotas)', D(2026, 6, 3), D(2026, 7, 3), 3671.60, 3464.14, D(2026, 7, 3), 'Pago', '', f'OK nas retenções. {END}'),
    ('2026-06', 'Mensalidade', 'NFS-e 8197 (monitoramento)', D(2026, 7, 6), D(2026, 8, 3), 2272.60, 2144.19, D(2026, 8, 3), 'Pago', '', 'OK: IRRF 1% e PIS/COFINS/CSLL 4,65%'),
    ('2026-06', 'Mensalidade', 'NFS-e 8199 (gestão de frotas)', D(2026, 7, 6), D(2026, 8, 3), 3699.93, 3490.88, D(2026, 8, 3), 'Pago', '', 'OK: IRRF 1% e PIS/COFINS/CSLL 4,65%'),
    ('2026-06', 'Visita improdutiva', 'NFS-e 8196', D(2026, 7, 6), D(2026, 8, 3), 45.00, 45.00, D(2026, 8, 3), 'Pago', '', ''),
    ('2026-07', 'Mensalidade', 'NFS-e 8966 (monitoramento)', D(2026, 8, 4), D(2026, 9, 3), 2224.46, 2098.79, D(2026, 9, 3), 'Pago', 'Pago em 03/09/2026 (informado pelo Matheus em 19/10)', f'OK nas retenções. {END}'),
    ('2026-07', 'Mensalidade', 'NFS-e 8967 (gestão de frotas)', D(2026, 8, 4), D(2026, 9, 3), 2279.68, 2150.87, D(2026, 9, 3), 'Pago', 'idem', f'OK nas retenções. {END}'),
    ('2026-07', 'Visita improdutiva', 'NFS-e 8978', D(2026, 8, 4), D(2026, 9, 3), 270.00, 270.00, None, 'Contestado (em aberto)', '6 visitas; e-mail de 30/09: boletos anteriores a 25/09 em análise por contestação', 'Sem PIS/COFINS/CSLL (4,65% = R$ 12,56, acima do mínimo de R$ 10)'),
    ('2026-07', 'Taxa de rescisão', 'NFS-e 8979', D(2026, 8, 4), D(2026, 9, 3), 2180.00, 1853.00, None, 'Contestado (em aberto)', '33 equipamentos; RMF6A73 cobrado R$ 260. E-mail de revisão já redigido', 'IRRF 15% sem PIS/COFINS/CSLL em NFS-e de serviço: pedir reemissão'),
    ('2026-08', 'Mensalidade', 'NFS-e 10277 (monitoramento)', D(2026, 9, 2), D(2026, 9, 25), 2219.30, 2093.91, None, 'Em aberto (vencida)', 'Notas só enviadas em 02/10 a pedido; 11/10 aguardava autorização de pagamento; 16/10 CEABS cobrou atraso', 'OK nas retenções'),
    ('2026-08', 'Mensalidade', 'NFS-e 10276 (gestão de frotas)', D(2026, 9, 2), D(2026, 9, 25), 2125.90, 2005.78, None, 'Em aberto (vencida)', 'idem', 'OK nas retenções'),
    ('2025-09', 'Equipamento não devolvido', 'NFS-e 124199', D(2025, 9, 12), D(2025, 10, 13), 800.00, 800.00, None, 'Contestado (em aberto)', '2 equipamentos x R$ 400 (cobrança de equipamentos 2025); Monitoramento pediu boleto em 17/10/2026', ''),
]
for x in docs26:
    add(*x, fonte='CONTROLE CEABS (aba DIN), PDFs das NFS-e e e-mails')

linhas.sort(key=lambda x: (x['mes_uso'], x['documento']))
for x in linhas:
    for k in ('emissao', 'vencimento', 'vencimento_padrao', 'pago_em'):
        x[k] = x[k].strftime('%d/%m/%Y') if hasattr(x[k], 'strftime') else x[k]
with open(os.path.join(EST, 'faturas.csv'), 'w', newline='', encoding='utf-8') as f:
    w = csv.DictWriter(f, fieldnames=list(linhas[0].keys()))
    w.writeheader()
    w.writerows(linhas)
tot = defaultdict(float)
for x in linhas:
    tot[x['mes_uso']] += x['valor_bruto']
print('mês      faturado  medição')
for m in sorted(tot):
    print(m, f'{tot[m]:9.2f} {medicao.get(m, 0):9.2f}', '' if abs(tot[m] - medicao.get(m, 0)) < 1 else '<- diferença')
