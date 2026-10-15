'use strict';

// Port of pipeline/extract_matriculas.py.
//
// The Python original hardcodes the header row/column position
// (data[4][15:25]) because calamine's grid happened to align that way for
// this sheet. SheetJS trims leading all-blank columns from its output
// (this sheet's `!ref` starts at C1, not A1), which shifts every fixed index
// — so instead of hardcoding a position, we search the grid for the exact
// header sequence and read relative to wherever it actually is. This is
// strictly more robust than the original (survives column-layout drift too).

const { sheetToGrid } = require('./xlsxGrid');

// Each entry is the canonical key we use downstream, plus every header text
// seen in the wild for that column. "MaxAlunos" got silently renamed to
// "Ideal" in the source spreadsheet at some point (same column, same
// meaning) — matching by canonical key with aliases survives that kind of
// drift instead of hard-failing on it like a single fixed name would.
const HEADER_FIELDS = [
  { key: 'FormacaoCodigo', aliases: ['FormacaoCodigo'] },
  { key: 'Modalidade', aliases: ['Modalidade'] },
  { key: 'Turma', aliases: ['Turma'] },
  { key: 'DataInicio', aliases: ['DataInicio'] },
  { key: 'DataFim', aliases: ['DataFim'] },
  { key: 'MinAlunos', aliases: ['MinAlunos'] },
  { key: 'MaxAlunos', aliases: ['MaxAlunos', 'Ideal'] },
  { key: 'InscritosAtivos', aliases: ['InscritosAtivos'] },
  { key: 'DiasParaComecar', aliases: ['DiasParaComecar'] },
  { key: 'FaltamParaMinimo', aliases: ['FaltamParaMinimo'] },
];

// Extended from the original 9 codes to match every FormacaoCodigo found in
// CONFIG-CM's product list once the dashboard started covering all of
// Board Academy's products instead of just the original 8 turma sheets.
const CODE_TO_PROGRAMA = {
  ABP: 'ABP',
  B_SPRINT: 'Board Sprint',
  BOARD_360: 'Board 360',
  PFIC: 'PFIC',
  PFCC: 'PFCC',
  PFCC_MX: 'PFCC LATAM',
  PFCC_CH: 'PFCC LATAM',
  CES: 'CES',
  LEAN_GOVERNANCE: 'Lean Governance',
  MBAI: 'MBAI',
  LINKEDIN_PROGRAM: 'LinkedIn Program',
  ADVISOR_PROGRAM: 'Advisor Program',
  IA_NA_PRÁTICA: 'IA na Prática',
  CNP_PRESENCIAL: 'CNP',
  CNP_ONLINE: 'CNP',
};

function trailingNumber(s) {
  if (typeof s !== 'string') return null;
  const m = /(\d+)\s*$/.exec(s.trim());
  return m ? parseInt(m[1], 10) : null;
}

// Python: isinstance(v, (int, float)) and v != '' — the `v != ''` half is
// always true for a real number in Python 3, so this reduces to a type check.
function numOrNone(v) {
  return typeof v === 'number' && !Number.isNaN(v) ? v : null;
}

function cellStr(v) {
  return typeof v === 'string' ? v.trim() : v;
}

function findHeaderBlock(grid) {
  for (let r = 0; r < grid.length; r++) {
    const row = grid[r];
    for (let c = 0; c < row.length; c++) {
      if (cellStr(row[c]) !== 'FormacaoCodigo') continue;
      const slice = row.slice(c, c + HEADER_FIELDS.length).map(cellStr);
      if (HEADER_FIELDS.every((f, i) => f.aliases.includes(slice[i]))) return { row: r, col: c };
    }
  }
  return null;
}

function extractMatriculas(workbook) {
  const ws = workbook.Sheets.CONFIGURACOES;
  if (!ws) throw new Error('sheet not found: CONFIGURACOES');
  const grid = sheetToGrid(ws);

  const headerPos = findHeaderBlock(grid);
  if (!headerPos) throw new Error('header row (FormacaoCodigo...FaltamParaMinimo) not found in CONFIGURACOES');

  const colStart = headerPos.col;
  const rows = [];
  for (let i = headerPos.row + 1; i < grid.length; i++) {
    const r = grid[i];
    const seg = r.slice(colStart, colStart + HEADER_FIELDS.length);
    if (!seg.some((x) => x !== null && x !== '')) continue;
    if (seg[0] === null || seg[0] === '') continue;
    const obj = {};
    for (let k = 0; k < HEADER_FIELDS.length; k++) obj[HEADER_FIELDS[k].key] = seg[k];
    rows.push(obj);
  }

  const matriculas = {};
  for (const r of rows) {
    const programa = CODE_TO_PROGRAMA[r.FormacaoCodigo];
    if (!programa) continue;
    const num = trailingNumber(r.Turma);
    if (num === null) continue;
    matriculas[`${programa}|${num}`] = {
      min: numOrNone(r.MinAlunos),
      max: numOrNone(r.MaxAlunos),
      atual: numOrNone(r.InscritosAtivos),
      faltam_para_minimo: numOrNone(r.FaltamParaMinimo),
      dias_para_comecar: numOrNone(r.DiasParaComecar),
      modalidade_ctrl: r.Modalidade,
      turma_codigo: r.Turma,
    };
  }
  return matriculas;
}

module.exports = { extractMatriculas, trailingNumber };
