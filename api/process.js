'use strict';

// Manual refresh: receives the three source spreadsheets from /admin and
// runs them through the shared pipeline (lib/pipeline.js). The automatic
// paths — api/sync.js (scheduled pull) and api/ingest.js (push from Power
// Automate) — end up in the same place; this endpoint stays as the
// fallback for when someone wants to force a refresh with local files.

const fs = require('fs/promises');
const { formidable } = require('formidable');

const { runPipeline, summarize } = require('../lib/pipeline');

function first(v) {
  return Array.isArray(v) ? v[0] : v;
}

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'method_not_allowed' });
    return;
  }

  if (!process.env.UPLOAD_PASSWORD) {
    res.status(500).json({ error: 'server_not_configured', message: 'UPLOAD_PASSWORD env var is not set.' });
    return;
  }

  let fields, files;
  try {
    const form = formidable({ maxFileSize: 30 * 1024 * 1024, multiples: false });
    [fields, files] = await form.parse(req);
  } catch (e) {
    res.status(400).json({ error: 'bad_upload', message: e.message });
    return;
  }

  const password = first(fields.password);
  if (password !== process.env.UPLOAD_PASSWORD) {
    res.status(401).json({ error: 'unauthorized' });
    return;
  }

  const cronogramaFile = first(files.cronograma);
  const matriculasFile = first(files.matriculas);
  const ocupacaoFile = first(files.ocupacao);
  if (!cronogramaFile || !matriculasFile || !ocupacaoFile) {
    res.status(400).json({ error: 'missing_files', message: 'Envie os três arquivos: Cronograma, Ocupação de Turmas e Planilha Master.' });
    return;
  }

  try {
    const { dashboard, parseErrors } = await runPipeline({
      cronograma: await fs.readFile(cronogramaFile.filepath),
      matriculas: await fs.readFile(matriculasFile.filepath),
      ocupacao: await fs.readFile(ocupacaoFile.filepath),
    });
    res.status(200).json(summarize(dashboard, parseErrors));
  } catch (e) {
    res.status(500).json({ error: 'processing_failed', message: e.message });
  } finally {
    for (const f of [cronogramaFile, matriculasFile, ocupacaoFile]) {
      if (f && f.filepath) fs.unlink(f.filepath).catch(() => {});
    }
  }
};
