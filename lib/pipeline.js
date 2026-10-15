'use strict';

// Single entry point that turns the three source workbooks (as Buffers) into
// the published dashboard blob. Shared by every way a refresh can start:
// manual upload (api/process.js), scheduled pull from Microsoft 365
// (api/sync.js) and file push from Power Automate (api/ingest.js).

const XLSX = require('xlsx');

const { extractTurmas } = require('./parseTurmas');
const { extractMatriculas } = require('./parseMatriculas');
const { extractOcupacao } = require('./parseOcupacao');
const { buildDashboard, sortTurmas } = require('./buildDashboard');
const { readDashboard, writeDashboard } = require('./turmaStore');

function readWorkbook(buf) {
  return XLSX.read(buf, { type: 'buffer', cellDates: true });
}

async function runPipeline({ cronograma, matriculas, ocupacao }, now = new Date()) {
  const cronogramaWb = readWorkbook(cronograma);
  const { turmas: rawTurmas, errors } = extractTurmas(cronogramaWb);
  const matriculasData = extractMatriculas(readWorkbook(matriculas));
  const ocupacaoData = extractOcupacao(readWorkbook(ocupacao));
  const eventosWorksheet = cronogramaWb.Sheets['Encontros Comunidades'];

  const dashboard = buildDashboard({
    rawTurmas,
    matriculas: matriculasData,
    ocupacao: ocupacaoData,
    eventosWorksheet,
    now,
  });

  // A scheduled refresh must not wipe turmas someone registered or edited by
  // hand in /admin — those carry origem: 'manual' (see normalizeManualTurma)
  // and win over a spreadsheet turma with the same id.
  const previous = await readDashboard();
  const manual = (previous.turmas || []).filter((t) => t.origem === 'manual');
  if (manual.length) {
    const manualIds = new Set(manual.map((t) => t.id));
    dashboard.turmas = dashboard.turmas.filter((t) => !manualIds.has(t.id)).concat(manual);
    sortTurmas(dashboard.turmas);
  }

  await writeDashboard(dashboard);

  return { dashboard, parseErrors: [...errors, ...ocupacaoData.errors] };
}

function summarize(dashboard, parseErrors) {
  const status = {};
  for (const t of dashboard.turmas) status[t.status] = (status[t.status] || 0) + 1;
  return {
    ok: true,
    gerado_em: dashboard.gerado_em,
    total_turmas: dashboard.turmas.length,
    status,
    eventos: dashboard.eventos_comunidades.length,
    parse_errors: parseErrors,
    ocupacao_nao_vinculada: dashboard.ocupacao_nao_vinculada,
  };
}

module.exports = { runPipeline, summarize };
