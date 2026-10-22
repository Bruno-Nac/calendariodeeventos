# Dash Blip - Disparos WhatsApp

Painel semanal de volume e custo dos disparos (Board Academy). Arquivo unico `index.html`, funciona no celular e no computador.

## Atualizacao semanal
1. Substitua `dados/relatorio_blip_consolidado.xlsx` pelo relatorio novo (mesmo nome).
2. Rode no PowerShell, dentro desta pasta:

   powershell -NoProfile -ExecutionPolicy Bypass -File .\atualizar.ps1

3. Faca commit e push de `index.html` e do Excel em `dados/`.

## Atualizacao do mes (arquivo custo_blip_*.xlsx)
Cada atualizacao do mes entra como um arquivo novo em `dados/` com nome `custo_blip_<mes>_<ano>.xlsx` (abas Premissas e Diario).
O script junta todos: os dias de `Diario` mais recentes substituem os anteriores, e os contatos de disparo e a nota vem do arquivo com a data mais recente.
Mantenha os arquivos dos meses anteriores em `dados/` para nao perder o historico. Os totais mensais e o comparativo anual sao recalculados a partir do historico diario.

## Abas Campanhas e Fatura
O script le as abas `Resumo`, `Fatura` e `Resultados por ação` da planilha principal
(`dados/relatorio_blip_consolidado.xlsx`). Se elas nao existirem, usa `dados/relatorio_blip_setembro_ceo.xlsx`.
Para atualizar toda semana: copie essas abas para a planilha principal (Excel: clicar com o direito na aba, Mover ou Copiar).
Para cada novo mes, copie as tres abas de novo e renomeie (ex.: `Resultados out26`, `Resumo out26`, `Fatura out26`). O mes e lido do titulo da aba
(celula A1/A2, formato "outubro/2026"). A aba Campanhas do painel mostra um seletor com todos os meses encontrados.
Tarifas confirmadas: o mes citado no titulo da aba `Fatura` e marcado como confirmado na aba Fatura do painel.

## Arquivos
- `index.html`: painel gerado (nao editar a mao).
- `dashboard_template.html`: modelo visual. Alteracoes de layout vao aqui.
- `atualizar.ps1`: le o Excel e gera o `index.html`.
- `logo-b.png`: logo (B dourado).
- `dados/`: Excel de origem.

## Acesso
Hoje o painel abre sem senha, entao use **repositorio privado** (os dados de custo ficam no HTML).
Opcao futura: `atualizar.ps1 -Senha "..."` gera o painel com tela de senha e dados criptografados.
Acesso via SSO Microsoft 365 (Azure Static Web Apps + Entra ID) fica para uma proxima etapa.

