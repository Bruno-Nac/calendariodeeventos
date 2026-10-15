'use strict';

// Port of pipeline/extract_turmas.py — same block-detection algorithm, operating
// on the padded grid from xlsxGrid.sheetToGrid() instead of a calamine grid.

const { sheetToGrid, norm } = require('./xlsxGrid');
const { isoDate, isoFromJsDate } = require('./dateUtils');

const MESES = {
  janeiro: 1, fevereiro: 2, março: 3, marco: 3, abril: 4, maio: 5, junho: 6,
  julho: 7, agosto: 8, setembro: 9, outubro: 10, novembro: 11, dezembro: 12,
};

// Returns [isoDateOrNull, rawWeekdayOrUnparsedTextOrNull] — mirrors parse_date()
// in extract_turmas.py, including its fall-through when a month name doesn't
// match (falls to the DD/MM/YYYY check) rather than short-circuiting.
function parseDate(v) {
  if (v === null || v === undefined || v === '') return [null, null];
  if (v instanceof Date) return [isoFromJsDate(v), null];
  if (typeof v !== 'string') return [null, null];

  const s = v.trim();

  let m = /(\d{1,2})\s*de\s*([a-zçãéêô]+)\s*de\s*(\d{4})/i.exec(s);
  if (m) {
    const day = parseInt(m[1], 10);
    const mesName = m[2].toLowerCase();
    const year = parseInt(m[3], 10);
    const mes = MESES[mesName];
    if (mes) {
      const iso = isoDate(year, mes, day);
      if (iso) {
        const weekday = s.indexOf(',') !== -1 ? s.split(',')[0].trim() : null;
        return [iso, weekday];
      }
      return [null, s];
    }
  }

  m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/.exec(s);
  if (m) {
    let d = parseInt(m[1], 10);
    const mo = parseInt(m[2], 10);
    let y = parseInt(m[3], 10);
    if (y < 100) y += 2000;
    const iso = isoDate(y, mo, d);
    if (iso) return [iso, null];
    return [null, s];
  }

  const upper = s.toUpperCase();
  if (upper === 'XXXX' || upper === 'TBD' || upper === '') return [null, null];
  return [null, s];
}

function findBlocks(rows) {
  const blocks = [];
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    let c = 0;
    while (c < row.length) {
      const val = norm(row[c]);
      if (typeof val === 'string' && val.toLowerCase() === 'data') {
        const headers = [];
        let cc = c;
        while (cc < row.length) {
          const v = norm(row[cc]);
          if (v === '') {
            const nxt = cc + 1 < row.length ? norm(row[cc + 1]) : '';
            if (nxt === '' || (typeof nxt === 'string' && nxt.toLowerCase() === 'data')) break;
            headers.push('');
            cc += 1;
            continue;
          }
          if (typeof v === 'string' && v.toLowerCase() === 'data' && cc !== c) break;
          headers.push(v);
          cc += 1;
        }
        blocks.push({ row: r, col: c, width: cc - c, headers });
        c = cc;
      } else {
        c += 1;
      }
    }
  }
  return blocks;
}

function colRangeOverlaps(a, b) {
  const aStart = a.col, aEnd = a.col + Math.max(a.width, 6);
  const bStart = b.col, bEnd = b.col + Math.max(b.width, 6);
  return !(aEnd <= bStart || bEnd <= aStart);
}

function extractSheetGrid(rows, programLabel) {
  const nrows = rows.length;
  const blocks = findBlocks(rows);
  const turmas = [];

  for (let i = 0; i < blocks.length; i++) {
    const blk = blocks[i];
    const r = blk.row, c = blk.col, headers = blk.headers;

    // Some sheets spread one block's title across several adjacent cells
    // (e.g. "Board Sprint" | "Turma 01" | "Presencial" as three separate
    // cells instead of one combined string). Collect every contiguous
    // non-blank cell starting from the first one found, so the title isn't
    // silently truncated to just the first fragment — but stop at the first
    // gap, so unrelated text further along the row (a "Legenda"/"Financeiro"
    // label sharing the row) never gets pulled in.
    let title = null;
    for (const back of [1, 2, 3]) {
      const rrUp = r - back;
      if (rrUp < 0) break;
      const rowUp = rows[rrUp];
      const start = Math.max(0, c - 2);
      const end = Math.min(rowUp.length, c + headers.length + 4);
      const parts = [];
      for (let cc = start; cc < end; cc++) {
        const cand = norm(rowUp[cc]);
        const isUrl = typeof cand === 'string' && /^https?:\/\//i.test(cand);
        if (cand !== '' && !isUrl) {
          parts.push(cand);
        } else if (parts.length) {
          break;
        }
      }
      if (parts.length) {
        title = parts.join(' - ');
        break;
      }
    }

    let endRow = nrows;
    let fallbackRow = null;
    let brokeInner = false;
    for (let j = i + 1; j < blocks.length; j++) {
      const other = blocks[j];
      if (other.row > r) {
        if (fallbackRow === null) fallbackRow = other.row;
        if (colRangeOverlaps(blk, other)) {
          endRow = other.row - 1;
          brokeInner = true;
          break;
        }
      }
    }
    if (!brokeInner && fallbackRow !== null) endRow = fallbackRow - 1;

    const colStart = c, colEnd = c + headers.length;
    const aulas = [];
    let blankStreak = 0;
    let lastDateIso = null;
    let lastDateRaw = null;
    let rr = r + 1;

    while (rr < Math.min(endRow, nrows)) {
      const row = rows[rr];
      const segment = [];
      for (let cc = colStart; cc < colEnd; cc++) segment.push(cc < row.length ? norm(row[cc]) : '');

      if (segment.every((x) => x === '')) {
        blankStreak += 1;
        if (blankStreak >= 2) break;
        rr += 1;
        continue;
      }
      blankStreak = 0;

      const rowdict = {};
      for (let k = 0; k < headers.length; k++) rowdict[headers[k]] = segment[k];

      const rawDate = Object.prototype.hasOwnProperty.call(rowdict, 'Data') ? rowdict.Data : '';
      let iso = null, rawTxt = null;
      if (rawDate !== '') [iso, rawTxt] = parseDate(rawDate);

      if (iso) {
        lastDateIso = iso;
        lastDateRaw = typeof rawDate === 'string' ? rawDate : null;
      } else if (rawTxt) {
        lastDateRaw = rawTxt;
        lastDateIso = null;
      }

      const disciplina = rowdict.Disciplina || rowdict.Disciplinas || '';
      const facilitador = rowdict.Facilitador || '';
      const horario = rowdict['Horário'] || rowdict.Horario || '';

      if (disciplina || facilitador || horario) {
        aulas.push({
          data: rawDate !== '' ? iso : lastDateIso,
          data_raw: typeof rawDate === 'string' && rawDate ? rawDate : lastDateRaw,
          dia: rowdict.Dia || '',
          horario,
          disciplina,
          facilitador,
          confirmacao: rowdict['Confirmação'] || '',
          mediador: rowdict.Mediador || '',
        });
      }
      rr += 1;
    }

    if (aulas.length > 0 && title) {
      turmas.push({ programa: programLabel, turma: title, aulas });
    }
  }

  return turmas;
}

const TARGETS = [
  ['PFIC ', 'PFIC'],
  ['PFCC', 'PFCC'],
  ['PFCC LATAM', 'PFCC LATAM'],
  ['ABP', 'ABP'],
  ['Lean Governance', 'Lean Governance'],
  ['Board 360', 'Board 360'],
  ['CES', 'CES'],
  ['Sprint', 'Board Sprint'],
  // Added when the Cronograma workbook grew from 8 to 11 actively-maintained
  // product tabs — same block layout ("Data" header cell + Disciplina/
  // Facilitador rows), so the existing block-detection algorithm above
  // handles them without any changes.
  ['MBAI', 'MBAI'],
  ['B2B', 'B2B'],
  ['Linkedln Program', 'LinkedIn Program'],
  ['Advisor Program', 'Advisor Program'],
  ['IA na Prática', 'IA na Prática'],
];

// CNP ("Conselho na Prática") isn't a class series with a facilitator
// teaching a discipline per session — it's a single recurring event
// (Presencial or Online) that students from several PFCC/CES turmas attend
// together. It doesn't fit the generic block algorithm above (which groups
// consecutive dated rows into one turma's "aulas"): here each row already
// *is* one full session, so it becomes its own one-aula turma instead.
const CNP_SHEET = 'CNP ';

function cnpSubmodalidade(title) {
  const t = title.toLowerCase();
  if (t.includes('online')) return 'Online';
  if (t.includes('experience')) return 'Experience';
  if (t.includes('presencial')) return 'Presencial';
  return '';
}

function extractCnpTurmas(workbook) {
  const ws = workbook.Sheets[CNP_SHEET];
  if (!ws) return { turmas: [], errors: [`CNP: sheet not found: ${CNP_SHEET}`] };
  const rows = sheetToGrid(ws);
  const blocks = findBlocks(rows);
  const turmas = [];

  for (let bi = 0; bi < blocks.length; bi++) {
    const blk = blocks[bi];
    const { row: r, col: c, headers } = blk;

    // Title sits 1-2 rows above the header, same search window as the
    // generic block parser, but CNP titles are consistently one row up.
    let title = '';
    for (const back of [1, 2]) {
      const rrUp = r - back;
      if (rrUp < 0) break;
      const rowUp = rows[rrUp];
      const parts = [];
      for (let cc = Math.max(0, c - 2); cc < Math.min(rowUp.length, c + headers.length + 2); cc++) {
        const cand = norm(rowUp[cc]);
        if (cand !== '') parts.push(cand);
        else if (parts.length) break;
      }
      if (parts.length) { title = parts.join(' - '); break; }
    }
    const submodalidade = cnpSubmodalidade(title);

    // Column roles by header text rather than fixed position, since the
    // three CNP sub-blocks don't share exactly the same column layout
    // (some have a spacer column, singular vs plural "Facilitador(es)").
    const idxData = headers.findIndex((h) => /^data$/i.test(h));
    const idxHorario = headers.findIndex((h) => /^hor[aá]rio$/i.test(h));
    const idxTurmaNum = headers.findIndex((h) => /^turmas$/i.test(h));
    const facilitadorIdx = headers.reduce((acc, h, i) => (/^facilitador/i.test(h) ? [...acc, i] : acc), []);
    const notaIdx = headers.findIndex((h, i) => /indicad|observa/i.test(h) && i > (idxTurmaNum >= 0 ? idxTurmaNum : 0));

    let endRow = rows.length;
    for (const other of blocks.slice(bi + 1)) {
      if (other.row > r) { endRow = other.row; break; }
    }

    let blankStreak = 0;
    for (let rr = r + 1; rr < endRow; rr++) {
      const row = rows[rr] || [];
      const cell = (i) => (i >= 0 && i < row.length ? norm(row[i]) : '');
      const segment = headers.map((_, i) => cell(c + i));
      if (segment.every((x) => x === '')) {
        blankStreak += 1;
        if (blankStreak >= 2) break;
        continue;
      }
      blankStreak = 0;

      const rawDate = cell(c + idxData);
      if (rawDate === '') continue;
      const [iso, rawTxt] = parseDate(rawDate);

      const facilitador = facilitadorIdx.map((i) => cell(c + i)).filter(Boolean).join(' / ');
      const nota = notaIdx >= 0 ? cell(c + notaIdx) : '';
      const turmaNum = idxTurmaNum >= 0 ? cell(c + idxTurmaNum) : '';
      const sessionLabel = turmaNum !== '' ? `Sessão ${turmaNum}` : `${rr}`;
      const turmaTitle = [`CNP ${submodalidade}`.trim(), sessionLabel, nota && /cancelad/i.test(nota) ? 'CANCELADA' : '']
        .filter(Boolean).join(' - ');

      turmas.push({
        programa: 'CNP',
        turma: turmaTitle,
        aulas: [{
          data: iso,
          data_raw: typeof rawDate === 'string' && rawDate ? rawDate : rawTxt,
          dia: '',
          horario: idxHorario >= 0 ? cell(c + idxHorario) : '',
          disciplina: nota && !/cancelad/i.test(nota) ? nota : '',
          facilitador,
          confirmacao: '',
          mediador: '',
        }],
      });
    }
  }

  return { turmas, errors: [] };
}

function extractTurmas(workbook) {
  const allTurmas = [];
  const errors = [];
  for (const [sheetName, label] of TARGETS) {
    try {
      const ws = workbook.Sheets[sheetName];
      if (!ws) throw new Error(`sheet not found: ${sheetName}`);
      const grid = sheetToGrid(ws);
      const turmas = extractSheetGrid(grid, label);
      allTurmas.push(...turmas);
    } catch (e) {
      errors.push(`${label}: ${e.message}`);
    }
  }

  const cnp = extractCnpTurmas(workbook);
  allTurmas.push(...cnp.turmas);
  errors.push(...cnp.errors);

  return { turmas: allTurmas, errors };
}

module.exports = { extractTurmas, extractSheetGrid, extractCnpTurmas, parseDate, findBlocks, TARGETS };
