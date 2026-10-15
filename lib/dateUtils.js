'use strict';

function pad2(n) {
  return String(n).padStart(2, '0');
}

// Validates a real calendar date and returns "YYYY-MM-DD", or null if invalid
// (e.g. day=31 in a 30-day month) — mirrors Python's datetime.date(...) raising ValueError.
function isoDate(y, m, d) {
  if (!Number.isInteger(y) || !Number.isInteger(m) || !Number.isInteger(d)) return null;
  const dt = new Date(Date.UTC(y, m - 1, d));
  if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return null;
  return `${y}-${pad2(m)}-${pad2(d)}`;
}

// exceljs returns date-formatted cells as JS Date objects built from the UTC
// serial value, so read the calendar date back out via UTC getters (never the
// local getters) to avoid timezone drift.
function isoFromJsDate(v) {
  return isoDate(v.getUTCFullYear(), v.getUTCMonth() + 1, v.getUTCDate());
}

module.exports = { isoDate, isoFromJsDate };
