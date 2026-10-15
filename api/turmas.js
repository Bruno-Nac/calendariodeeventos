'use strict';

// Manual CRUD for individual turmas — create/update/delete a single record
// directly, without going through an xlsx upload. Reuses the exact same
// normalization (status/id/data_inicio/data_fim/matriculas derivation) as
// the xlsx path (lib/buildDashboard.js) so a manually-entered turma renders
// identically to one that came from a spreadsheet.

const { readDashboard, writeDashboard } = require('../lib/turmaStore');
const { normalizeManualTurma, sortTurmas } = require('../lib/buildDashboard');
const { isoDate } = require('../lib/dateUtils');

function checkAuth(req) {
  const pw = req.headers['x-admin-password'];
  return typeof pw === 'string' && pw.length > 0 && pw === process.env.UPLOAD_PASSWORD;
}

function readJsonBody(req) {
  return new Promise((resolve, reject) => {
    if (req.body !== undefined && req.body !== null) {
      if (typeof req.body === 'string') {
        try { resolve(req.body ? JSON.parse(req.body) : {}); } catch (e) { reject(e); }
      } else {
        resolve(req.body);
      }
      return;
    }
    let data = '';
    req.on('data', (c) => { data += c; });
    req.on('end', () => {
      try { resolve(data ? JSON.parse(data) : {}); } catch (e) { reject(e); }
    });
    req.on('error', reject);
  });
}

module.exports = async function handler(req, res) {
  if (!process.env.UPLOAD_PASSWORD) {
    res.status(500).json({ error: 'server_not_configured', message: 'UPLOAD_PASSWORD env var is not set.' });
    return;
  }
  if (!checkAuth(req)) {
    res.status(401).json({ error: 'unauthorized' });
    return;
  }

  const now = new Date();
  const todayIso = isoDate(now.getFullYear(), now.getMonth() + 1, now.getDate());

  try {
    const data = await readDashboard();
    if (!Array.isArray(data.turmas)) data.turmas = [];

    if (req.method === 'POST') {
      const body = await readJsonBody(req);
      if (!body.programa || !body.turma) {
        res.status(400).json({ error: 'missing_fields', message: 'programa e turma são obrigatórios.' });
        return;
      }
      const rec = normalizeManualTurma(body, todayIso);
      let id = rec.id, n = 2;
      while (data.turmas.some((t) => t.id === id)) id = `${rec.id}-${n++}`;
      rec.id = id;

      data.turmas.push(rec);
      sortTurmas(data.turmas);
      data.gerado_em = now.toISOString();
      await writeDashboard(data);
      res.status(200).json({ ok: true, turma: rec });
      return;
    }

    if (req.method === 'PUT') {
      const body = await readJsonBody(req);
      if (!body.id) {
        res.status(400).json({ error: 'missing_id' });
        return;
      }
      const idx = data.turmas.findIndex((t) => t.id === body.id);
      if (idx === -1) {
        res.status(404).json({ error: 'not_found' });
        return;
      }
      const rec = normalizeManualTurma(body, todayIso);
      data.turmas[idx] = rec;
      sortTurmas(data.turmas);
      data.gerado_em = now.toISOString();
      await writeDashboard(data);
      res.status(200).json({ ok: true, turma: rec });
      return;
    }

    if (req.method === 'DELETE') {
      const body = await readJsonBody(req);
      if (!body.id) {
        res.status(400).json({ error: 'missing_id' });
        return;
      }
      const before = data.turmas.length;
      data.turmas = data.turmas.filter((t) => t.id !== body.id);
      if (data.turmas.length === before) {
        res.status(404).json({ error: 'not_found' });
        return;
      }
      data.gerado_em = now.toISOString();
      await writeDashboard(data);
      res.status(200).json({ ok: true });
      return;
    }

    res.status(405).json({ error: 'method_not_allowed' });
  } catch (e) {
    res.status(500).json({ error: 'failed', message: e.message });
  }
};
