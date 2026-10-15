'use strict';

// Parses "Ocupação de Turmas vNN.xlsx" — the weekly-maintained occupancy
// dashboard (PROGRAMAS sheet: one row per turma, every product, with
// min/ideal/current-enrollment/percentage/trend already computed). This is
// the primary source for occupancy numbers (see buildDashboard.js's merge
// logic); Planilha_Master_Controle_Alunos_2026 (parseMatriculas.js) is kept
// as a secondary/fallback source for turmas this file doesn't cover.
//
// The sheet's own "PROGRAMA" column mixes the product with schedule/
// modality info ("PFCC MX  2ª/4ª", "ABP Online  3ª/5ª", "CNP Pres. T.
// Presencial"...) instead of the plain canonical names parseTurmas.js uses
// ("PFCC LATAM", "ABP", "CNP"...) — mapPrograma() below normalizes that.

const { sheetToGrid, norm } = require('./xlsxGrid');
const { isoDate, isoFromJsDate } = require('./dateUtils');

const SHEET_NAME = 'PROGRAMAS';

// Longest/most-specific pattern first, since e.g. "PFCC MX" must not fall
// through to the plain "PFCC" rule.
const PROGRAM_RULES = [
  [/^CNP\b/i, 'CNP'],
  [/PFCC\s*MX/i, 'PFCC LATAM'],
  [/PFCC\s*LATAM/i, 'PFCC LATAM'],
  [/PFCC/i, 'PFCC'],
  [/^ABP/i, 'ABP'],
  [/Lean\s*Governance/i, 'Lean Governance'],
  [/^PFIC/i, 'PFIC'],
  [/Board\s*Sprint/i, 'Board Sprint'],
  [/Board\s*360/i, 'Board 360'],
  [/^CES/i, 'CES'],
  [/IA\s*na\s*Pr[aá]tica/i, 'IA na Prática'],
  [/LinkedIn/i, 'LinkedIn Program'],
  [/Advisor\s*Program/i, 'Advisor Program'],
  [/MBAI/i, 'MBAI'],
  [/B2B/i, 'B2B'],
];

function mapPrograma(rawLabel) {
  const s = (rawLabel || '').trim();
  for (const [pat, canonical] of PROGRAM_RULES) {
    if (pat.test(s)) return canonical;
  }
  return null;
}

function cellStr(v) {
  return typeof v === 'string' ? v.trim() : v;
}

// "T40", "T01", "T7" -> 40 / 1 / 7. Returns null if the code doesn't look
// like "T<digits>".
function turmaCodeNumero(code) {
  if (typeof code !== 'string') return null;
  const m = /^T\s*0*(\d+)$/i.exec(code.trim());
  return m ? parseInt(m[1], 10) : null;
}

// Ocupação's "início" column is date-formatted for most rows but a few slip
// through as raw serial numbers or as "12/dez" without a year — same drift
// pattern already seen in the Cronograma workbook.
function parseInicio(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date) return isoFromJsDate(v);
  if (typeof v === 'number') {
    // Excel serial date (days since 1899-12-30), matches XLSX's own epoch.
    const ms = Math.round((v - 25569) * 86400 * 1000);
    return isoFromJsDate(new Date(ms));
  }
  if (typeof v === 'string') {
    const s = v.trim();
    let m = /^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?$/.exec(s);
    if (m) {
      const d = parseInt(m[1], 10), mo = parseInt(m[2], 10);
      let y = m[3] ? parseInt(m[3], 10) : new Date().getFullYear();
      if (y < 100) y += 2000;
      return isoDate(y, mo, d);
    }
  }
  return null;
}

function findHeaderRow(grid) {
  for (let r = 0; r < grid.length; r++) {
    const row = grid[r].map(cellStr);
    if (row.some((c) => c === 'PROGRAMA') && row.some((c) => c === 'TURMA')) return r;
  }
  return -1;
}

// Returns { records, byKey, errors }.
//   records: every parsed row, including ones we couldn't canonicalize
//            (programa: null) or number-match (turmaNumero: null) — kept so
//            buildDashboard.js can still try a date-proximity match and so
//            unmatched rows can be surfaced instead of silently dropped.
//   byKey:   records with a resolved `${programa}|${numero}` key, in the
//            exact shape parseMatriculas.js produces, for direct reuse by
//            the existing matricula-lookup path in buildDashboard.js.
function extractOcupacao(workbook) {
  const ws = workbook.Sheets[SHEET_NAME];
  if (!ws) throw new Error(`sheet not found: ${SHEET_NAME}`);
  const grid = sheetToGrid(ws);

  const headerRow = findHeaderRow(grid);
  if (headerRow === -1) throw new Error('header row (PROGRAMA...TENDÊNCIA) not found in PROGRAMAS');

  const headers = grid[headerRow].map((h) => cellStr(norm(h)));
  const col = (name) => headers.findIndex((h) => h === name);
  const idx = {
    programa: col('PROGRAMA'),
    turma: col('TURMA'),
    inicio: col('INÍCIO'),
    min: headers.findIndex((h) => /^M[IÍ]N\.?$/i.test(h)),
    ideal: col('IDEAL'),
    atual: headers.findIndex((h) => /^ALUNOS/i.test(h)),
    percent: headers.findIndex((h) => /^%/i.test(h)),
    tendencia: col('TENDÊNCIA'),
  };
  if (idx.programa === -1 || idx.turma === -1) {
    throw new Error('required columns (PROGRAMA/TURMA) not found in PROGRAMAS header');
  }

  const records = [];
  const errors = [];

  for (let r = headerRow + 1; r < grid.length; r++) {
    const row = grid[r];
    const programaRaw = cellStr(row[idx.programa]);
    const turmaCodigo = cellStr(row[idx.turma]);
    if (!programaRaw && !turmaCodigo) continue; // blank row
    if (!programaRaw || !turmaCodigo) continue; // stray fragment (legend, footer count, ...)

    const programa = mapPrograma(programaRaw);
    if (!programa) {
      errors.push(`programa não reconhecido: "${programaRaw}" (turma ${turmaCodigo})`);
    }

    const numero = turmaCodeNumero(turmaCodigo);
    const percentRaw = idx.percent >= 0 ? row[idx.percent] : null;
    const percentAtual = typeof percentRaw === 'number' ? Math.round(percentRaw * 100) : null;

    records.push({
      programa,
      programaRaw,
      turmaCodigo,
      turmaNumero: numero,
      inicio: idx.inicio >= 0 ? parseInicio(row[idx.inicio]) : null,
      min: typeof row[idx.min] === 'number' ? row[idx.min] : null,
      ideal: typeof row[idx.ideal] === 'number' ? row[idx.ideal] : null,
      atual: typeof row[idx.atual] === 'number' ? row[idx.atual] : null,
      percent_atual: percentAtual,
      tendencia: idx.tendencia >= 0 ? cellStr(row[idx.tendencia]) : '',
    });
  }

  const byKey = {};
  for (const rec of records) {
    if (rec.programa === null || rec.turmaNumero === null) continue;
    byKey[`${rec.programa}|${rec.turmaNumero}`] = {
      min: rec.min,
      max: rec.ideal,
      atual: rec.atual,
      faltam_para_minimo: rec.min !== null ? Math.max(0, rec.min - (rec.atual || 0)) : null,
      dias_para_comecar: null,
      modalidade_ctrl: null,
      turma_codigo: rec.turmaCodigo,
    };
  }

  return { records, byKey, errors };
}

module.exports = { extractOcupacao, mapPrograma, turmaCodeNumero, parseInicio, PROGRAM_RULES };
