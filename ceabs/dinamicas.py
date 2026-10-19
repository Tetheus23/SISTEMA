"""Insere tabelas dinâmicas reais (DataPilot do LibreOffice) na planilha montada por montar_painel.py e salva em .xlsx.

Uso: python3 dinamicas.py <entrada.xlsx> <saida.xlsx>
Precisa do LibreOffice (soffice) e do módulo uno. Sobe um soffice headless próprio numa porta local e encerra no fim.
Cada dinâmica fica numa aba própria, logo depois do Painel, e lê a aba Base inteira.
"""
import os
import shutil
import subprocess
import sys
import tempfile
import time

import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.sheet.DataPilotFieldOrientation import COLUMN, DATA, PAGE, ROW
from com.sun.star.sheet.GeneralFunction import SUM
from com.sun.star.table import CellAddress

ENTRADA, SAIDA = map(os.path.abspath, sys.argv[1:3])
PORTA = 2300 + os.getpid() % 600
DINAMICAS = [
    # aba, linhas, colunas, filtros (página)
    ('Dinâmica 1 - Mês', ['Mês de uso', 'Mês de vencimento'], ['Status'], ['Tipo de cobrança']),
    ('Dinâmica 2 - Cliente', ['Local / cliente no mês'], ['Status'], ['Mês de uso', 'Tipo de cobrança']),
    ('Dinâmica 3 - Veículo', ['Placa', 'Modelo'], ['Mês de uso'], ['Status', 'Tipo de cobrança']),
]


def prop(nome, valor):
    p = PropertyValue()
    p.Name, p.Value = nome, valor
    return p


perfil = tempfile.mkdtemp(prefix='lo_ceabs_')
soffice = subprocess.Popen(['soffice', '--headless', '--invisible', '--norestore', '--nologo',
                            f'-env:UserInstallation=file://{perfil}',
                            f'--accept=socket,host=127.0.0.1,port={PORTA};urp;'],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
try:
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
    for _ in range(60):
        try:
            ctx = resolver.resolve(f'uno:socket,host=127.0.0.1,port={PORTA};urp;StarOffice.ComponentContext')
            break
        except Exception:
            time.sleep(1)
    else:
        raise SystemExit('LibreOffice não respondeu')
    desktop = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
    copia = os.path.join(perfil, 'entrada.xlsx')  # cópia isolada: evita trava (.~lock) de execuções anteriores
    shutil.copy(ENTRADA, copia)
    doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(copia), '_blank', 0, (prop('Hidden', True),))
    if doc is None:
        raise SystemExit(f'LibreOffice não abriu {ENTRADA}')
    folhas = doc.Sheets
    base = folhas.getByName('Base')
    cursor = base.createCursor()
    cursor.gotoEndOfUsedArea(False)
    for pos, (nome, linhas, colunas, paginas) in enumerate(DINAMICAS, 1):
        folhas.insertNewByName(nome, pos)
        # o índice da aba Base muda a cada aba inserida: o intervalo de origem é lido de novo
        fonte = base.getCellRangeByPosition(0, 0, cursor.RangeAddress.EndColumn, cursor.RangeAddress.EndRow).RangeAddress
        aba = folhas.getByName(nome)
        tabelas = aba.getDataPilotTables()
        desc = tabelas.createDataPilotDescriptor()
        desc.setSourceRange(fonte)
        campos = desc.getDataPilotFields()
        for n in paginas:
            campos.getByName(n).Orientation = PAGE
        for n in linhas:
            campos.getByName(n).Orientation = ROW
        for n in colunas:
            campos.getByName(n).Orientation = COLUMN
        valor = campos.getByName('Valor cobrado')
        valor.Orientation = DATA
        valor.Function = SUM
        desc.ShowFilterButton = False
        desc.GrandTotalName = 'Total'
        desc.RowGrand = True
        desc.ColumnGrand = True
        endereco = CellAddress()
        endereco.Sheet, endereco.Column, endereco.Row = pos, 0, 2
        tabelas.insertNewByName(nome.split(' - ')[1].lower(), endereco, desc)
        aba.getCellByPosition(0, 0).String = f'{nome}: soma do valor cobrado (fonte: aba Base). Use os filtros acima da tabela.'
        aba.getCellByPosition(0, 0).CharWeight = 150
        # moeda nos valores e largura das colunas
        cur = aba.createCursor()
        cur.gotoEndOfUsedArea(False)
        area = aba.getCellRangeByPosition(0, 0, cur.RangeAddress.EndColumn, cur.RangeAddress.EndRow)
        fmt = doc.NumberFormats.queryKey('[$R$-416] #.##0,00', doc.CharLocale, False)
        if fmt == -1:
            fmt = doc.NumberFormats.addNew('[$R$-416] #.##0,00', doc.CharLocale)
        area.NumberFormat = fmt
        for c in range(cur.RangeAddress.EndColumn + 1):
            aba.Columns.getByIndex(c).OptimalWidth = True
        aba.Columns.getByIndex(0).Width = 4500 if pos != 2 else 7500
    doc.Sheets.getByName('Painel')  # mantém o Painel como primeira aba
    doc.getCurrentController().setActiveSheet(folhas.getByName('Painel')) if doc.getCurrentController() else None
    doc.calculateAll()
    doc.storeToURL(uno.systemPathToFileUrl(SAIDA), (prop('FilterName', 'Calc MS Excel 2007 XML'),))
    doc.close(True)
finally:
    soffice.terminate()
    try:
        soffice.wait(20)
    except subprocess.TimeoutExpired:
        soffice.kill()
print('ok:', SAIDA)
