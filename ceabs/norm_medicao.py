import openpyxl, sys, json


def num(v):
    return round(v, 2) if isinstance(v, (int, float)) else 0.0


# uso: norm_medicao.py <xlsx> -> imprime JSON {ref, total, itens:[...]} com todas as abas
wb=openpyxl.load_workbook(sys.argv[1],data_only=True)
ref=None; itens=[]; resumo={}; cliente=None
for ws in wb:
    rows=[r for r in ws.iter_rows(values_only=True) if any(c is not None for c in r)]
    hdr=next((i for i,r in enumerate(rows) if r[0]=='Cliente'),None)
    if hdr is None: continue
    for r in rows[:hdr]:
        if r[0]=='REFERÊNCIA:': ref=r[1]
        if r[0]=='CLIENTE:': cliente=r[1]
        if r[3] and r[3] not in ('RESUMO',) and isinstance(r[4],(int,float)): resumo[r[3]]=round(r[4],2)
    col = {str(c).strip(): i for i, c in enumerate(rows[hdr]) if c}  # colunas pelo cabeçalho: o layout muda entre meses
    get = lambda r, nome: r[col[nome]] if nome in col and col[nome] < len(r) else None
    for r in rows[hdr+1:]:
        if not get(r, 'Código'): continue
        valor = num(get(r, 'Valor a Faturar'))
        mensal = get(r, 'Monitoramento') if 'Monitoramento' in col else valor
        canc, ret = get(r, 'Data de Cancelamento'), get(r, 'Data de Retirada')
        itens.append(dict(aba=ws.title,codigo=get(r,'Código'),chassi=get(r,'Chassi'),placa=get(r,'Placa'),produto=get(r,'Produto'),mensal=mensal,prorata=num(get(r,'Pró-Rata')),valor=valor,instalacao=get(r,'Data de Instalação'),cancelamento=str(canc) if canc else None,retirada=str(ret) if ret else None))
    if ws.title=='Planilha1': break
# ignora abas duplicadas sem cabecalho de resumo (ex.: Planilha1 copia da principal)
seen=set(); ded=[]
for i in itens:
    if i['aba']=='Planilha1': continue
    ded.append(i)
print(json.dumps(dict(ref=ref,cliente=cliente,resumo=resumo,itens=ded),ensure_ascii=False))
