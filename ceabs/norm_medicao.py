import openpyxl, sys, json
# uso: norm_medicao.py <xlsx> -> imprime JSON {ref, total, itens:[...]} com todas as abas
wb=openpyxl.load_workbook(sys.argv[1],data_only=True)
ref=None; itens=[]; resumo={}
for ws in wb:
    rows=[r for r in ws.iter_rows(values_only=True) if any(c is not None for c in r)]
    hdr=next((i for i,r in enumerate(rows) if r[0]=='Cliente'),None)
    if hdr is None: continue
    for r in rows[:hdr]:
        if r[0]=='REFERÊNCIA:': ref=r[1]
        if r[3] and r[3] not in ('RESUMO',) and isinstance(r[4],(int,float)): resumo[r[3]]=round(r[4],2)
    for r in rows[hdr+1:]:
        if not r[3]: continue
        itens.append(dict(aba=ws.title,codigo=r[3],chassi=r[4],placa=r[5],produto=r[6],mensal=r[7],prorata=round(r[8] or 0,2),valor=round(r[9] or 0,2),instalacao=r[10],cancelamento=str(r[11]) if r[11] else None,retirada=str(r[12]) if r[12] else None))
    if ws.title=='Planilha1': break
# ignora abas duplicadas sem cabecalho de resumo (ex.: Planilha1 copia da principal)
seen=set(); ded=[]
for i in itens:
    if i['aba']=='Planilha1': continue
    ded.append(i)
print(json.dumps(dict(ref=ref,resumo=resumo,itens=ded),ensure_ascii=False))
