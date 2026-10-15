'use strict';

const XLSX = require('xlsx');

// Array-of-arrays for one sheet, uniformly padded to the sheet's used range —
// the same shape python_calamine's `.to_python()` gives the pipeline scripts.
// Critically, SheetJS does NOT propagate a merged range's value to its slave
// cells (calamine doesn't either — only the anchor cell holds a value in the
// raw XML); the block parser relies on that rawness to detect "same
// date/weekday as the row above" and carry it forward itself.
function sheetToGrid(ws) {
  if (!ws) return [];
  return XLSX.utils.sheet_to_json(ws, {
    header: 1,
    raw: true,
    defval: null,
    blankrows: true,
  });
}

function norm(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'string') return v.trim();
  return v;
}

module.exports = { sheetToGrid, norm };
