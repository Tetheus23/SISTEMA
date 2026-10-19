# Conferência do faturamento CEABS (rastreadores)

Scripts para auditar a cobrança da CEABS: medições, notas, pedidos de cancelamento e uso real de cada veículo.
Os dados e as planilhas geradas ficam em `entregas/ceabs/`, que não é versionado.

## 1. E-mails do Drive
1. `inventario_drive.py`: consolida as listagens da pasta do Drive em um inventário (id, data, tamanho, título).
2. `coletar_blobs.py` e `proximo_lote.sh`: casam os .msg baixados (`ReadMcpResourceTool`) com o inventário e listam o próximo lote.
3. `extrair_emails.py <raw> <proc> <inventario.tsv> <emails.jsonl>`: lê cada .msg com a skill leitor-de-msg (corpo, anexos, e-mails aninhados) e grava um JSON por e-mail.

## 2. Auditoria (`pipeline.sh entregas/ceabs`)
1. `medicoes.py`: normaliza todas as planilhas de medição anexadas (`norm_medicao.py`, colunas pelo cabeçalho).
2. `estruturar.py`: monta itens cobrados, cadastro CEABS por chassi, pedidos de cancelamento da Grillo, eventos da CEABS e documentos fiscais.
3. `auditar.py`: classifica cada item (veículo x mês) em OK, CONTESTAR, SEM USO ou VERIFICAR, usando a frota (`estrutura/frota/*.csv`, exportada do Supabase com `sb2csv.py`), diesel, Ticketlog e o relatório de posição da CEABS. Também gera a análise dos equipamentos não devolvidos e a lista de retiradas.
4. `faturas.py`: uma linha por NFS-e/fatura, com mês de uso, vencimento (padrão: dia 3 do 2º mês após o uso), pagamento e conferência fiscal.

## 3. Planilha
1. `montar_painel.py <pasta> <saida.xlsx>`: Painel, veículos do mês, Contestar CEABS, Equipamentos, Retiradas, Faturas e Base.
2. `dinamicas.py <entrada.xlsx> <saida.xlsx>`: insere as tabelas dinâmicas reais (DataPilot do LibreOffice) e salva em .xlsx.
3. `historico_cobrancas.py <pasta> <saida.xlsx>`: planilha limpa com o histórico item a item (colunas da CEABS + mês apurado, vencimento, notas, valor sem imposto rateado do boleto, pagamento, cliente e status) e um painel com os casos por status e os últimos 5 meses. Com um 3º argumento (a planilha devolvida com ajustes), STATUS, clientes, pagamento e observações vêm dela.
4. `processo_pagamento.py <historico.xlsx> <forpag.xlsb> <AAAA-MM> <saida> <pdfs>`: processo de pagamento no padrão da Ticket Log (FORPAG preenchido, notas e boletos, relatório resumido e consolidado em um PDF) e o e-mail para a diretoria.

`parse_docs.py` lê NFS-e e boletos em PDF; `montar_relatorio.py` é o relatório antigo (substituído pelo painel).
