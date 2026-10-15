# Régua de Turmas

Dashboard do cronograma de turmas e programas da Board Academy. Site estático (`index.html`) + funções serverless que processam as planilhas de origem e mantêm o painel atualizado — sem Python, sem SharePoint automatizado, sem depender de ninguém rodar comando nenhum.

## Como funciona

Fluxo manual (atalho — o fluxo automático está na próxima seção):

1. Alguém baixa as três planilhas mais recentes do SharePoint:
   - **Cronograma Oficial Board - Turmas.xlsx** — cronograma de aulas por turma, de todos os produtos ativos.
   - **Ocupação de Turmas.xlsx** — número de alunos/vagas por turma (fonte principal de ocupação).
   - **Planilha_Master_Controle_Alunos_2026.xlsx** — controle de matrículas (usada como reforço para turmas que a Ocupação ainda não cobre).
2. Abre `/admin` neste site, digita a senha e arrasta os três arquivos.
3. A função `/api/process` lê as planilhas, roda a mesma lógica de extração que os scripts Python faziam (portada e estendida para JavaScript, ver `lib/`), cruza cronograma + ocupação + matrículas, e salva o resultado no Vercel Blob.
4. O painel (`/`) busca os dados de `/api/dashboard-data` a cada carregamento — sempre reflete o último upload, sem precisar de redeploy.

## Atualização automática (sem ninguém subir arquivo)

O painel se atualiza sozinho: um cron do Vercel busca as planilhas no Microsoft 365 e reprocessa. `/admin` continua funcionando como atalho manual, e turmas cadastradas à mão em `/admin` (origem `manual`) são preservadas a cada atualização.

### Configuração

`api/sync.js` baixa as 3 planilhas direto do SharePoint/OneDrive pela Microsoft Graph e reprocessa. Roda todo dia às 10h UTC (7h em Brasília) pelo cron em `vercel.json`, e também pode ser chamada a qualquer momento (`GET /api/sync` com o header `x-admin-password`).

Configuração única — exige um administrador do Microsoft 365 (Entra ID / Azure AD):

1. **Registrar um app** em *Microsoft Entra ID → App registrations → New registration* (qualquer nome, ex.: "Regua de Turmas"). Anote o *Directory (tenant) ID* e o *Application (client) ID*.
2. Em *Certificates & secrets*, criar um **client secret** e copiar o valor.
3. Em *API permissions → Microsoft Graph → Application permissions*, adicionar **`Files.Read.All`** e clicar em **Grant admin consent**. (Se o admin preferir restringir a só esses 3 arquivos/sites, `Sites.Selected` também funciona, concedendo acesso por site.)
4. No SharePoint/OneDrive, copiar o link de compartilhamento de cada uma das 3 planilhas (*Compartilhar → Copiar link*).
5. No Vercel (*Settings → Environment Variables*), cadastrar:

| Variável | Valor |
|---|---|
| `MS_TENANT_ID` | Directory (tenant) ID do passo 1 |
| `MS_CLIENT_ID` | Application (client) ID do passo 1 |
| `MS_CLIENT_SECRET` | secret do passo 2 |
| `SRC_CRONOGRAMA_URL` | link do *Cronograma Oficial Board - Turmas.xlsx* |
| `SRC_OCUPACAO_URL` | link da *Ocupação de Turmas.xlsx* |
| `SRC_MATRICULAS_URL` | link da *Planilha_Master_Controle_Alunos_2026.xlsx* |
| `CRON_SECRET` | qualquer string longa e aleatória (o Vercel a envia no cron) |

6. Redeploy. Para testar sem esperar o horário: `GET /api/sync` com `x-admin-password: <UPLOAD_PASSWORD>` — a resposta traz o mesmo resumo do upload manual.

> **Frequência:** o plano Hobby do Vercel só permite cron **diário**. Em plano Pro, troque `schedule` em `vercel.json` por algo como `"0 * * * *"` (de hora em hora). Em qualquer plano, um agendador externo pode chamar `/api/sync` com a frequência que quiser.

## Produtos cobertos

PFIC, PFCC, PFCC LATAM, ABP, Lean Governance, Board 360, CES, Board Sprint, MBAI, LinkedIn Program, Advisor Program, IA na Prática, B2B e CNP (Conselho na Prática — tratado como turma de sessão única, já que não é uma turma com disciplina/facilitador como as demais). Produtos com aba oculta/descontinuada no Cronograma (GRC, Empremax, Master Board, C-Level Program, Succession Program, AI Leadership Program, Global Program, PFCC FORTAL, ESG Starter, AI onBoard...) não entram — reavalie `lib/parseTurmas.js`'s `TARGETS` se algum desses voltar a ser mantido.

## Como o cruzamento de ocupação funciona

`lib/parseOcupacao.js` lê a aba `PROGRAMAS` da planilha de Ocupação (uma linha por turma, todos os produtos) e normaliza o rótulo livre da coluna PROGRAMA ("PFCC MX 2ª/4ª", "ABP Online 3ª/5ª"...) para o nome canônico usado no Cronograma. `lib/buildDashboard.js` casa cada turma extraída do Cronograma com uma linha da Ocupação por número de turma primeiro ("PFCC116" ↔ "T116") e, se isso falhar, pela data de início mais próxima dentro do mesmo programa (tolerância de 10 dias) — como as duas planilhas são mantidas por pessoas diferentes, nomes e numeração às vezes divergem. Se uma linha da Ocupação não casar com nenhuma turma, ela aparece em `ocupacao_nao_vinculada` na resposta do upload — vale conferir essa lista depois de cada atualização.

## Setup necessário no Vercel (uma vez só)

1. **Vercel Blob**: no projeto, vá em `Storage` → `Create Database` → `Blob` e conecte ao projeto. Isso cria automaticamente a variável de ambiente `BLOB_READ_WRITE_TOKEN`.
2. **Senha de upload**: em `Settings` → `Environment Variables`, adicione `UPLOAD_PASSWORD` com uma senha forte. É ela que protege quem pode escrever dados novos em `/admin`.
3. Depois de configurar as duas, faça um redeploy para as variáveis entrarem em vigor.

Sem isso, `/admin` continua acessível mas todo upload falha (a função recusa se `UPLOAD_PASSWORD` não estiver configurada, e não tem onde salvar sem o Blob conectado).

## Estrutura

- `index.html` — o painel (busca dados via fetch, não tem mais dado embutido).
- `admin/index.html` — página de upload self-service (3 arquivos).
- `api/process.js` — upload manual: recebe os 3 arquivos, processa, salva no Blob.
- `api/sync.js` + `lib/graphSource.js` — atualização automática: baixa os 3 arquivos do Microsoft 365 (Graph API) e reprocessa; chamada pelo cron de `vercel.json`.
- `lib/pipeline.js` — o núcleo compartilhado pelo upload manual e o sync (parse → cruzamento → publicação, preservando turmas manuais).
- `api/dashboard-data.js` — serve os dados atuais pro painel (com fallback para o `data/dashboard-data.seed.json` embutido, caso nunca tenha havido upload — esse seed ainda reflete a geração anterior, de só 8 produtos e 2 arquivos; ele é substituído no primeiro upload feito com os 3 arquivos novos).
- `lib/` — a lógica de parsing. `parseTurmas.js`/`parseMatriculas.js` começaram como porte fiel de `pipeline/*.py` (validado linha a linha contra a saída Python: 238/238 turmas, 2433/2433 campos de aula, 78/78 matrículas, zero divergências) e depois ganharam cobertura extra só em JS — os 5 produtos novos, CNP e `parseOcupacao.js` **não têm equivalente no Python** e não foram validados contra uma saída "ground truth" separada, só conferidos manualmente linha a linha contra as planilhas reais durante o desenvolvimento. Rode `test/e2e.js` contra as planilhas reais depois de qualquer mudança aqui.
- `pipeline/` — os scripts Python originais, mantidos como referência/fallback manual (não são publicados pelo Vercel — ver `.vercelignore`). **Ficaram desatualizados**: não cobrem os produtos novos, CNP nem Ocupação de Turmas — servem só de referência histórica agora, não como fallback completo.

## Fallback manual (se algum dia precisar)

Os scripts em `pipeline/` continuam funcionando para os 8 produtos originais: baixe as planilhas para `Downloads`, rode `python extract_turmas.py`, `python extract_matriculas.py`, `python build_dashboard_data.py` (requer `pip install python_calamine`). Isso gera um `regua_de_turmas.html` autocontido — útil para um snapshot avulso dos produtos antigos, mas não cobre os produtos adicionados depois (MBAI, LinkedIn Program, Advisor Program, IA na Prática, B2B, CNP) nem afeta o site publicado no Vercel (que só muda via `/admin` ou editando `index.html`/`data/dashboard-data.seed.json` diretamente).

