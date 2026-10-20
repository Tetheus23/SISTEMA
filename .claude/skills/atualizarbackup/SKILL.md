---
name: atualizarbackup
description: "Use SEMPRE que o Matheus escrever #atualizarbackup (ou pedir para atualizar o backup geral / o PDF de contexto para outra IA). Refaz o inventário de tudo (orientações, rotinas agendadas, agenda, banco, Drive, skills, sistemas, sessões e últimas demandas), atualiza a pasta BACKUP do Google Drive só com a última versão de cada arquivo e gera de novo o PDF de contexto."
---

# #atualizarbackup

Objetivo: manter a pasta `00 - ASSISTENTE CLAUDE/BACKUP` do Google Drive com uma cópia atual de tudo o que foi construído com o Claude e um PDF que permita a outra IA (ou a uma sessão nova) continuar de onde parou.

Regra de ouro: **só a última versão de cada arquivo**. Ao trocar um arquivo, mandar o anterior para a lixeira do Drive (`trash_file`, reversível por 30 dias). Nunca apagar nada fora da pasta BACKUP, nunca alterar o banco (só leitura) e nunca versionar dados pessoais no GitHub (o repositório SISTEMA é público).

## Onde fica cada coisa (localizar pelo nome, não por id fixo)
- Drive: pasta `00 - ASSISTENTE CLAUDE` > `BACKUP`, com as subpastas:
  `01 - RESUMO PARA OUTRA IA (PDF)`, `02 - ORIENTACOES E ROTINAS`, `03 - SKILLS`, `04 - AGENDA`, `05 - BANCO DE DADOS (ultima versao)`, `06 - SISTEMAS E CODIGO`. Na raiz: `00 - LEIA-ME`.
  Se alguma não existir, criar.
- Banco: Supabase, projeto "Assistente Matheus - LGR" (`list_projects`), schemas `assistente` (principal), `public` (tabelas `ctr_*` de contratos públicos e `diesel_*` do painel de diesel) e `dropship`.
- Backups noturnos do banco: Drive `00 - ASSISTENTE CLAUDE/08 - BACKUPS DO BANCO (somente consulta)`, arquivos `AAAA-MM-DD - BACKUP <tabela>`.
- Rotinas agendadas: `list_triggers` (claude-code-remote).
- Sessões anteriores: `list_sessions` (título e `post_turn_summary`) e, para o pedido exato, `list_events` com `kinds: ["user"]`.
- Skills: as sincronizadas na conta (no Claude Code ficam em `~/.claude/skills/synced/*/manifest.json`) e as do repositório (`.claude/skills/`).
- Código: repositório GitHub `Tetheus23/SISTEMA` (todas as branches) e projetos da Vercel (`list_projects`).
- Agenda: Google Agenda principal.
- Se a sessão tiver memória (`/preferences.md`, `/areas/*.md`) ou acesso ao PC (Desktop Commander / remote-devices: `H:\Meu Drive\00 - ASSISTENTE CLAUDE`, importador em `%USERPROFILE%\.assistente-claude`), incluir também. Se não tiver, dizer em uma linha o que ficou de fora.

## Passos
1. **Avisar** em uma linha que vai levar alguns minutos.
2. **Inventário** (só leitura, em paralelo quando der):
   - rotinas agendadas: nome, horário, prompt completo, último status (falhou? por quê?);
   - agenda: hoje até +45 dias (agrupar os recorrentes em uma linha cada);
   - banco: tabelas com contagem de linhas; `max(alterado_em)` por tabela em `assistente.historico`; tarefas abertas (não 'Concluída'/'Cancelada'), prazos vencidos e dos próximos 15 dias; últimas 30 linhas do `diario_bordo`; e-mails com status de pendência; documentos atuais do `repositorio_documentos` com validade; `sistema_demandas` pendentes; contratos com fim/renovação/reajuste próximos;
   - Drive: árvore de pastas de `00 - ASSISTENTE CLAUDE` e arquivos modificados desde o último backup;
   - sessões novas desde o último backup: título, resumo, pedidos do Matheus;
   - skills, repositório (branches e últimos commits) e Vercel.
3. **Banco na pasta 05**: para cada tabela pequena do schema `assistente` (exceto `config`, que tem token, e `historico`):
   - se o backup noturno mais recente da tabela é posterior à última alteração no `historico`, copiar com `copy_file` (título `<tabela> (estado em AAAA-MM-DD)`);
   - se não, exportar o CSV na hora (padrão de SQL do backup noturno) e criar com `contentMimeType: text/csv`;
   - mandar para a lixeira a versão anterior da mesma tabela na pasta 05;
   - tabelas grandes (dezenas de milhares de linhas: `importacoes`, `manut_*`, `frota_*`, `apresentacoes`, `documento_chunks`, `diesel_abastecimentos`...) não cabem no conector: listar nome e contagem no LEIA-ME e dizer que o original está no Supabase.
4. **Orientações e rotinas (pasta 02)**: regravar `ORIENTACOES E ROTINAS` com as regras de atendimento (seção 2 do PDF) e o prompt completo de cada rotina agendada.
5. **Skills (pasta 03)**: cada `SKILL.md` e cada script das skills próprias (as que não são exemplos da Anthropic) como arquivo de texto (`<skill> - SKILL.md`, `<skill> - scripts - <arquivo>`), mais o `atualizarbackup.zip` para subir no claude.ai. Só substituir o que mudou (comparar `updatedAt` do manifest com a data do arquivo no Drive).
6. **Agenda (pasta 04)**: CSV `agenda (AAAA-MM-DD a AAAA-MM-DD)` com data, horário, título, descrição curta.
7. **Sistemas (pasta 06)**: se houver commit novo, o código do repositório como texto, em partes de até ~50 KB (`SISTEMA-<frente> - codigo parte N de M.txt`), cada arquivo começando com `===== ARQUIVO: <caminho> =====`. Montar as partes com `git show <branch>:<arquivo>`.
8. **PDF (pasta 01)**: `CONTEXTO MATHEUS - AAAA-MM-DD.pdf`, em português, nesta ordem:
   0. Como usar este documento (instruções para a outra IA);
   1. Quem é o Matheus, empresa, hierarquia e equipe;
   2. Orientações de atendimento (tom, formato, regras de banco, agenda, nomes de arquivo, Drive);
   3. Ecossistema (banco, Drive, agenda, GitHub, Vercel, Shopify, PC) e onde fica cada coisa;
   4. Rotinas automáticas e status;
   5. Skills: o que cada uma faz e quando usar;
   6. Frentes de trabalho com o contexto base de cada uma;
   7. Agenda atual;
   8. Últimas demandas (mais recente primeiro);
   9. Radar: o mais urgente, com datas;
   10. Inconsistências e pontos de atenção (dados que não batem, rotinas falhando);
   11. Limitações deste backup.
   Mandar o PDF anterior da pasta 01 para a lixeira. Também enviar o PDF ao Matheus no chat.
9. **LEIA-ME** na raiz da BACKUP: índice das pastas, data da atualização, o que foi trocado e o que ficou de fora.
10. **Compartilhamento**: a outra IA usa a mesma conta Google, então não precisa compartilhar. Só compartilhar se o Matheus der um e-mail; nesse caso, `share_file` na pasta BACKUP como `reader`.
11. **Resposta final curta**: link da pasta, link do PDF, lista do que mudou desde o último backup, o que não deu para incluir e os 3 itens mais urgentes do radar.

## Como subir arquivos pelo conector do Drive (aprendido no 1º backup)
- `create_file` só recebe o conteúdo dentro da chamada. Base64 de binário (PDF comprimido, zip) é caro e fácil de corromper: evitar.
- PDF: gerar um PDF só com texto ASCII (fontes padrão Helvetica, acentos como escapes octais WinAnsi, sem compressão, ~40 KB para 7 páginas) e subir com `textContent`, `contentMimeType: application/pdf` e `disableConversionToGoogleType: true`. O gerador está nesta pasta: `scripts/pdf_enxuto.py <pasta_do_conteudo.py> <saida.pdf>` (o conteúdo fica em `conteudo.py`, lista `SECOES`).
- Código e skills: texto puro com `disableConversionToGoogleType: true` (fica .md/.py/.txt, não vira Google Doc).
- Tabelas e agenda: `text/csv` (vira Google Sheets). Textos para ler/colar: `text/markdown` (vira Google Docs).
- Conferir cada envio pelo `fileSize` devolvido, que tem que bater com o tamanho local em bytes. Para binário, baixar com `download_file_content` (o resultado grande vai para arquivo) e comparar o MD5.

## Regras
- Só leitura no banco e na agenda. A única escrita é na pasta BACKUP do Drive.
- Não incluir tokens, senhas ou o conteúdo de `assistente.config`.
- Datas em DD/MM/AAAA no texto; nomes de arquivo `AAAA-MM-DD - ...` sem `\ / : * ? " < > |`.
- Tom do Matheus: português, direto, curto, carioca.
- Se o Drive bloquear um envio (verificação de segurança), tentar de novo uma vez e, se falhar, dizer qual arquivo ficou de fora.
