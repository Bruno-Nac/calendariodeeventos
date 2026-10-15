'use strict';

// Port of pipeline/build_dashboard_data.py's merge/status/events logic.

const { sheetToGrid, norm } = require('./xlsxGrid');
const { isoDate, isoFromJsDate } = require('./dateUtils');

const PROGRAM_ORDER = [
  'PFCC', 'CES', 'Lean Governance', 'Board 360', 'ABP', 'PFCC LATAM', 'PFIC', 'Board Sprint',
  'MBAI', 'LinkedIn Program', 'Advisor Program', 'IA na Prática', 'B2B', 'CNP',
];
const PROGRAM_LABELS = {
  PFCC: 'Programa de Formação e Certificação de Conselheiros - PFCC',
  CES: 'Conselho Empresarial da Saúde - CES',
  'Lean Governance': 'Lean Governance',
  'Board 360': 'Board 360',
  ABP: 'Advanced Board Program - ABP',
  'PFCC LATAM': 'Programa de Formación y Certificación de Consejeros Latam',
  PFIC: 'Programa de Finanças para Conselheiros - PFIC',
  'Board Sprint': 'Board Sprint',
  MBAI: 'MBAI',
  'LinkedIn Program': 'LinkedIn Program',
  'Advisor Program': 'Advisor Program',
  'IA na Prática': 'IA na Prática',
  B2B: 'B2B',
  CNP: 'Conselho na Prática - CNP',
};

const EXCLUDE_PATTERNS = /cancelad|rascunho/i;

// Each program tries its patterns in order; the first one that matches the
// turma title wins. Kept as arrays (not one regex per program) because two
// title conventions coexist within some sheets — e.g. ABP's oldest, already-
// concluded turmas are titled "ABP Híbrido - T1" while every turma since is
// "ABP16", "ABP17"... Only used to look turmas up in the Master-file
// (parseMatriculas.js) fallback source; the primary source (parseOcupacao.js)
// matches by its own "T40"-style code instead, see resolveOcupacaoMatch().
const NUM_PATTERNS = {
  PFCC: [/PFCC\s*(\d+)/i],
  'PFCC LATAM': [/PFCC\s*(\d+)/i],
  CES: [/CES\s*(\d+)/i],
  ABP: [/ABP\s*(\d+)/i, /\bT\s*(\d+)\b/i],
  'Board Sprint': [/Turma\s*(\d+)/i],
  'Lean Governance': [/Turma\s*(\d+)/i],
  PFIC: [/Turma\s*(\d+)/i],
  'Board 360': [/^T\s*(\d+)/i],
  MBAI: [/Turma\s*(\d+)/i],
  'LinkedIn Program': [/Program\s*(\d+)/i, /Turma\s*(\d+)/i],
  'Advisor Program': [/Turma\s*(\d+)/i],
  'IA na Prática': [/PR[AÁ]TICA\s*(\d+)/i, /Turma\s*(\d+)/i],
  B2B: [/Turma\s*(\d+)/i],
  CNP: [/Sess[aã]o\s*(\d+)/i],
};

const MESES = {
  janeiro: 1, fevereiro: 2, março: 3, marco: 3, abril: 4, maio: 5, junho: 6,
  julho: 7, agosto: 8, setembro: 9, outubro: 10, novembro: 11, dezembro: 12,
};

const COMUNIDADES_SERIES = ['Board Select', 'Lean Board na Prática', 'Health Board Meeting', 'AI Business Community'];

function modalidadeFromTitle(title) {
  const t = title.toLowerCase();
  if (t.includes('presencial')) return 'Presencial';
  if (t.includes('online')) return 'Online';
  return '';
}

function cleanTitle(title) {
  return title.replace(/\s+/g, ' ').trim();
}

function turmaNumero(programa, titulo) {
  const pats = NUM_PATTERNS[programa];
  if (!pats) return null;
  for (const pat of pats) {
    const m = pat.exec(titulo);
    if (m) return parseInt(m[1], 10);
  }
  return null;
}

function lookupMatricula(matriculas, programa, titulo) {
  const num = turmaNumero(programa, titulo);
  if (num === null) return null;
  return matriculas[`${programa}|${num}`] || null;
}

function dateDiffDays(isoA, isoB) {
  const a = Date.parse(`${isoA}T00:00:00Z`);
  const b = Date.parse(`${isoB}T00:00:00Z`);
  return Math.round((a - b) / 86400000);
}

// Primary occupancy source (see parseOcupacao.js): first try matching by
// turma number (Ocupação's "T13" vs the number parsed from the Cronograma
// title, same idea as lookupMatricula() above but against the Ocupação
// data instead of the Master file). If that fails — a title format
// turmaNumero() doesn't recognize, or the two spreadsheets are just out of
// sync on numbering — fall back to matching the closest "início" date
// within the same program, since both files carry real dates. `claimed`
// prevents one Ocupação row from being matched to two different turmas.
const DATE_MATCH_TOLERANCE_DAYS = 10;

function resolveOcupacaoMatch(ocupacao, programa, numero, dataInicio, claimed) {
  if (!ocupacao) return null;

  if (numero !== null) {
    const rec = ocupacao.byKey[`${programa}|${numero}`];
    if (rec) {
      claimed.add(`${programa}|${rec.turma_codigo}`);
      return rec;
    }
  }

  if (dataInicio) {
    let best = null, bestDiff = Infinity;
    for (const rec of ocupacao.records) {
      if (rec.programa !== programa || !rec.inicio) continue;
      const claimKey = `${programa}|${rec.turmaCodigo}`;
      if (claimed.has(claimKey)) continue;
      const diff = Math.abs(dateDiffDays(rec.inicio, dataInicio));
      if (diff <= DATE_MATCH_TOLERANCE_DAYS && diff < bestDiff) { bestDiff = diff; best = rec; }
    }
    if (best) {
      claimed.add(`${programa}|${best.turmaCodigo}`);
      return {
        min: best.min,
        max: best.ideal,
        atual: best.atual,
        faltam_para_minimo: best.min !== null ? Math.max(0, best.min - (best.atual || 0)) : null,
        dias_para_comecar: null,
        modalidade_ctrl: null,
        turma_codigo: best.turmaCodigo,
      };
    }
  }

  return null;
}

function progKey(p) {
  const idx = PROGRAM_ORDER.indexOf(p);
  return idx === -1 ? PROGRAM_ORDER.length : idx;
}

function slugify(programa, turma) {
  return `${programa}-${turma}`
    .replace(/[^a-zA-Z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .toLowerCase();
}

// Matrículas entered directly on a manual turma record (min/max/atual, no
// number-matching against a separate spreadsheet) — same capacity-status
// derivation as the xlsx path, but faltam_para_minimo is computed here
// instead of coming from a source column, since there isn't one.
function computeMatriculasOut(mat) {
  if (!mat) return null;
  const min = mat.min === undefined || mat.min === null || mat.min === '' ? null : Number(mat.min);
  const max = mat.max === undefined || mat.max === null || mat.max === '' ? null : Number(mat.max);
  const atual = mat.atual === undefined || mat.atual === null || mat.atual === '' ? null : Number(mat.atual);
  if (min === null && max === null && atual === null) return null;

  const atualNum = atual || 0;
  let capStatus;
  if (min !== null && max !== null && (min > 0 || max > 0)) {
    capStatus = atualNum >= (min || 0) ? 'atingido' : 'abaixo';
  } else {
    capStatus = 'sem_config';
  }
  return {
    min,
    max,
    atual,
    faltam_para_minimo: min !== null ? Math.max(0, min - atualNum) : null,
    status_capacidade: capStatus,
  };
}

// Normalizes one manually-entered turma into the exact same record shape
// buildTurmas() produces from the xlsx path, so the front-end never has to
// know which source a turma came from. Used by api/turmas.js for both create
// (no id yet — caller assigns one after checking for collisions) and update
// (id passed through unchanged so links/state referencing it keep working).
function normalizeManualTurma(input, todayIso) {
  const programa = input.programa;
  const turma = cleanTitle(String(input.turma || ''));
  const modalidade = input.modalidade || modalidadeFromTitle(turma);

  const aulas = (Array.isArray(input.aulas) ? input.aulas : []).map((a) => ({
    data: a.data || null,
    data_raw: a.data_raw || null,
    dia: a.dia || '',
    horario: a.horario || '',
    disciplina: a.disciplina || '',
    facilitador: a.facilitador || '',
  }));
  const aulasSorted = aulas.slice().sort((a, b) => {
    const da = a.data || '9999-99-99', db = b.data || '9999-99-99';
    return da < db ? -1 : da > db ? 1 : 0;
  });

  const dates = aulas.map((a) => a.data).filter(Boolean).sort();
  let dataInicio = null, dataFim = null, status;
  if (dates.length) {
    dataInicio = dates[0];
    dataFim = dates[dates.length - 1];
    if (dataInicio >= todayIso) status = 'futura';
    else if (dataFim < todayIso) status = 'concluida';
    else status = 'andamento';
  } else {
    status = 'sem_data';
  }

  return {
    id: input.id || slugify(programa, turma),
    origem: 'manual',
    programa,
    turma,
    modalidade,
    status,
    data_inicio: dataInicio,
    data_fim: dataFim,
    total_aulas: aulas.length,
    matriculas: computeMatriculasOut(input.matriculas),
    aulas: aulasSorted,
  };
}

function sortTurmas(turmas) {
  turmas.sort((a, b) => {
    const ka = progKey(a.programa), kb = progKey(b.programa);
    if (ka !== kb) return ka - kb;
    const da = a.data_inicio || '9999', db = b.data_inicio || '9999';
    return da < db ? -1 : da > db ? 1 : 0;
  });
  return turmas;
}

function buildTurmas(rawTurmas, matriculas, todayIso, ocupacao) {
  const turmasOut = [];
  const skipped = [];
  let matchedCount = 0;
  const claimedOcupacao = new Set();

  for (const t of rawTurmas) {
    if (EXCLUDE_PATTERNS.test(t.turma)) {
      skipped.push(`${t.programa} | ${t.turma}`);
      continue;
    }
    const dates = t.aulas.map((a) => a.data).filter(Boolean).sort();
    let dataInicio = null, dataFim = null, status;
    if (dates.length) {
      dataInicio = dates[0];
      dataFim = dates[dates.length - 1];
      if (dataInicio >= todayIso) status = 'futura';
      else if (dataFim < todayIso) status = 'concluida';
      else status = 'andamento';
    } else {
      status = 'sem_data';
    }

    const aulasSorted = t.aulas.slice().sort((a, b) => {
      const da = a.data || '9999-99-99', db = b.data || '9999-99-99';
      return da < db ? -1 : da > db ? 1 : 0;
    });

    const num = turmaNumero(t.programa, t.turma);
    const mat = resolveOcupacaoMatch(ocupacao, t.programa, num, dataInicio, claimedOcupacao)
      || (num !== null ? matriculas[`${t.programa}|${num}`] || null : null);
    let matriculasOut = null;
    if (mat && (mat.min !== null || mat.max !== null || mat.atual !== null)) {
      matchedCount += 1;
      const atual = mat.atual || 0;
      const minimo = mat.min, maximo = mat.max;
      let capStatus;
      if (minimo !== null && maximo !== null && (minimo > 0 || maximo > 0)) {
        capStatus = atual >= (minimo || 0) ? 'atingido' : 'abaixo';
      } else {
        capStatus = 'sem_config';
      }
      matriculasOut = {
        min: minimo,
        max: maximo,
        atual: mat.atual,
        faltam_para_minimo: mat.faltam_para_minimo,
        status_capacidade: capStatus,
      };
    }

    turmasOut.push({
      id: slugify(t.programa, t.turma),
      origem: 'planilha',
      programa: t.programa,
      turma: cleanTitle(t.turma),
      modalidade: modalidadeFromTitle(t.turma),
      status,
      data_inicio: dataInicio,
      data_fim: dataFim,
      total_aulas: t.aulas.length,
      matriculas: matriculasOut,
      aulas: aulasSorted.map((a) => ({
        data: a.data,
        data_raw: a.data_raw,
        dia: a.dia,
        horario: a.horario,
        disciplina: a.disciplina,
        facilitador: a.facilitador,
      })),
    });
  }

  turmasOut.sort((a, b) => {
    const ka = progKey(a.programa), kb = progKey(b.programa);
    if (ka !== kb) return ka - kb;
    const da = a.data_inicio || '9999', db = b.data_inicio || '9999';
    return da < db ? -1 : da > db ? 1 : 0;
  });

  // Ocupação rows that never got claimed by a turma — either its program
  // label didn't map to anything (see parseOcupacao.js's PROGRAM_RULES) or
  // no Cronograma turma matched it by number or by nearby start date.
  // Surfaced instead of silently dropped so a mismatch (naming drift,
  // extra/renamed turma) shows up in the admin upload result to review,
  // rather than quietly showing wrong or missing capacity numbers.
  const unmatchedOcupacao = ocupacao
    ? ocupacao.records
      .filter((rec) => rec.programa === null || !claimedOcupacao.has(`${rec.programa}|${rec.turmaCodigo}`))
      .map((rec) => ({ programa: rec.programa, programa_raw: rec.programaRaw, turma_codigo: rec.turmaCodigo, inicio: rec.inicio }))
    : [];

  return { turmasOut, skipped, matchedCount, unmatchedOcupacao };
}

function parseEvtDate(v) {
  if (v instanceof Date) return isoFromJsDate(v);
  if (typeof v === 'string') {
    const m = /(\d{1,2})\s*de\s*([a-zçãéêô]+)\s*de\s*(\d{4})/i.exec(v);
    if (m) {
      const day = parseInt(m[1], 10);
      const mesName = m[2].toLowerCase();
      const year = parseInt(m[3], 10);
      const mes = MESES[mesName];
      if (mes) return isoDate(year, mes, day);
    }
  }
  return null;
}

function extractEventos(worksheet, now) {
  if (!worksheet) return [];
  const grid = sheetToGrid(worksheet);

  const sections = [];
  for (let i = 0; i < grid.length; i++) {
    const row = grid[i];
    const nonempty = row.filter((c) => norm(c) !== '');
    if (nonempty.length === 1 && typeof nonempty[0] === 'string' && norm(row[0]) === nonempty[0]) {
      sections.push([i, norm(row[0])]);
    }
  }

  const cutoff = isoDate(now.getFullYear(), now.getMonth() + 1, 1);
  const eventos = [];

  for (let idx = 0; idx < sections.length; idx++) {
    const [rowI, serie] = sections[idx];
    const headerI = rowI + 1;
    const headers = (grid[headerI] || []).map(norm);
    const endI = idx + 1 < sections.length ? sections[idx + 1][0] : grid.length;

    for (let r = headerI + 1; r < endI; r++) {
      const row = grid[r] || [];
      const vals = {};
      for (let k = 0; k < headers.length; k++) vals[headers[k]] = norm(row[k]);
      const dataVal = vals.Data;
      if (dataVal === '' || dataVal === undefined || dataVal === null) continue;
      const iso = parseEvtDate(dataVal);
      if (!iso || iso < cutoff) continue;
      eventos.push({
        serie,
        data: iso,
        status: vals.Status || '',
        convidado: vals.Convidado || '',
        tema: vals.Tema || '',
        programas: vals['Programas que participam'] || '',
      });
    }
  }

  eventos.sort((a, b) => {
    if (a.data !== b.data) return a.data < b.data ? -1 : 1;
    const ia = COMUNIDADES_SERIES.includes(a.serie) ? COMUNIDADES_SERIES.indexOf(a.serie) : 9;
    const ib = COMUNIDADES_SERIES.includes(b.serie) ? COMUNIDADES_SERIES.indexOf(b.serie) : 9;
    return ia - ib;
  });

  return eventos;
}

function buildDashboard({ rawTurmas, matriculas, ocupacao, eventosWorksheet, now = new Date() }) {
  const todayIso = isoDate(now.getFullYear(), now.getMonth() + 1, now.getDate());
  const { turmasOut, unmatchedOcupacao } = buildTurmas(rawTurmas, matriculas, todayIso, ocupacao);
  const eventos = extractEventos(eventosWorksheet, now);
  return {
    gerado_em: now.toISOString(),
    fonte_arquivo: 'Cronograma Oficial Board - Turmas.xlsx',
    program_order: PROGRAM_ORDER,
    program_labels: PROGRAM_LABELS,
    comunidades_series: COMUNIDADES_SERIES,
    turmas: turmasOut,
    eventos_comunidades: eventos,
    ocupacao_nao_vinculada: unmatchedOcupacao,
  };
}

module.exports = {
  buildDashboard,
  buildTurmas,
  extractEventos,
  normalizeManualTurma,
  sortTurmas,
  slugify,
  PROGRAM_ORDER,
  PROGRAM_LABELS,
};
