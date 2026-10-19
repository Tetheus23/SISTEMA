import sys, re, subprocess, json
def txt(p): return subprocess.run(['pdftotext','-layout',p,'-'],capture_output=True,text=True).stdout
def num(s): return float(s.replace('R$','').replace('.','').replace(',','.').strip())
out=[]
for p in sys.argv[1:]:
    t=txt(p); f=re.sub(r'[ \t]+',' ',t)
    d={'arquivo':p.split('/')[-1]}
    if 'NFS-E' in t.upper() and 'DANFSe' in t or 'CHAVE DE ACESSO' in t.upper():
        d['tipo']='NF'
        m=re.search(r'NÚMERO DA NFS-?[eE].*?\n\s*(\d+)\s+(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{2}/\d{4})',t,re.S)
        if m: d['nf'],d['competencia'],d['emissao']=m.group(1),m.group(2),m.group(3)
        m=re.search(r'IBS e CBS\.,\s*([^\n]+)',f); d['descricao']=m.group(1).strip() if m else None
        m=re.search(r'VALOR TOTAL DA NFS-?E.*?\n.*?(R?\$? ?[\d.]+,\d{2})',t,re.S|re.I); d['valor']=num(m.group(1)) if m else None
        m=re.search(r'VALOR L[ÍI]QUIDO DA NFS-?e.*?\n(.*)',t,re.I)
        vals=re.findall(r'[\d.]+,\d{2}',m.group(1)) if m else []
        d['liquido']=num(vals[1]) if len(vals)>=2 else (num(vals[-1]) if vals else None)
        m=re.search(r'IRRF.*?\n(.*)',t); v=re.findall(r'[\d.]+,\d{2}',m.group(1)) if m else []
        d['irrf']=num(v[0]) if v else 0
        d['csrf']=num(v[1]) if len(v)>1 else 0
    else:
        d['tipo']='BOLETO'
        m=re.search(r'(\d{2}/\d{2}/\d{4})\s+(\d{10,})\s+DM',f); 
        if m: d['data_doc'],d['num_doc']=m.group(1),m.group(2)
        venc=re.findall(r'Vencimento\s*\n?.*?(\d{2}/\d{2}/\d{4})',t)
        allv=re.findall(r'\d{2}/\d{2}/\d{4}',t)
        d['vencimento']=next((x for x in allv if x!=d.get('data_doc')),None)
        m=re.search(r'Valor do documento\s*\n.*?R\$\s*([\d.]+,\d{2})',t)
        d['valor']=num(m.group(1)) if m else None
        m=re.search(r'Nosso n[úu]mero.*?\n.*?\n?.*?(\d{1}/\d{8,}-?\d?|0{3,}\d+-\d)',t)
    out.append(d)
print(json.dumps(out,ensure_ascii=False,indent=0))
