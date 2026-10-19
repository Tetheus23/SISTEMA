# Conferência do faturamento CEABS (rastreadores)

Scripts para conferir a medição mensal da CEABS contra as NFs/boletos e a nossa frota.

1. `parse_docs.py <pdfs...>`: lê as NFS-e e os boletos (número, valores, retenções, vencimento). Os valores vão conferidos para `notas.json`.
2. `norm_medicao.py <PlanilhaMedicao.xlsx> > AAAA-MM.json`: normaliza a planilha de medição de uma referência (todas as abas: Principal, Taxa de Rescisão, Visita Improdutiva).
3. `montar_relatorio.py <pasta_dados> <saida.xlsx>`: gera a planilha com o resumo, a conferência por veículo, a parte fiscal, o histórico mensal, o histórico por veículo e o detalhe de rescisões e visitas.

A localização dos veículos vem do Supabase (`assistente.frota_veiculos` e `assistente.frota_passagens`), exportada para `frota_agosto.csv`.
Os dados e as planilhas geradas ficam em `entregas/ceabs/`, que não é versionado.
