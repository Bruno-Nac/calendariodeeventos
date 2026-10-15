'use strict';
// Ad-hoc validation script (not shipped): loads the real xlsx fixtures with
// exceljs, runs the JS port, and diffs the result against the Python
// pipeline's ground-truth JSON for the same files.

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const { extractTurmas } = require('../lib/parseTurmas');
const { extractMatriculas } = require('../lib/parseMatriculas');

function loadWorkbook(filePath) {
  const raw = fs.readFileSync(filePath);
  return XLSX.read(raw, { type: 'buffer', cellDates: true });
}

const CRONOGRAMA = process.argv[2];
const MATRICULAS_XLSX = process.argv[3];
const TURMAS_JSON_GROUND_TRUTH = process.argv[4];
const MATRICULAS_JSON_GROUND_TRUTH = process.argv[5];

async function main() {
  const wb = loadWorkbook(CRONOGRAMA);
  const { turmas, errors } = extractTurmas(wb);

  console.log('--- TURMAS ---');
  console.log('errors:', errors);
  console.log('total turmas (JS):', turmas.length);

  const byPrograma = {};
  let totalAulas = 0;
  for (const t of turmas) {
    byPrograma[t.programa] = (byPrograma[t.programa] || 0) + 1;
    totalAulas += t.aulas.length;
  }
  console.log('por programa (JS):', byPrograma);
  console.log('total aulas (JS):', totalAulas);

  if (TURMAS_JSON_GROUND_TRUTH) {
    const gt = JSON.parse(fs.readFileSync(TURMAS_JSON_GROUND_TRUTH, 'utf8'));
    console.log('total turmas (PY):', gt.turmas.length);
    const gtByPrograma = {};
    let gtAulas = 0;
    for (const t of gt.turmas) {
      gtByPrograma[t.programa] = (gtByPrograma[t.programa] || 0) + 1;
      gtAulas += t.aulas.length;
    }
    console.log('por programa (PY):', gtByPrograma);
    console.log('total aulas (PY):', gtAulas);

    // Deep-compare turma-by-turma (order should match since both scan sheets top-to-bottom).
    let mismatches = 0;
    const n = Math.max(turmas.length, gt.turmas.length);
    for (let i = 0; i < n; i++) {
      const a = turmas[i], b = gt.turmas[i];
      if (!a || !b) {
        console.log(`#${i}: MISSING — JS=${a ? a.turma : 'undefined'} PY=${b ? b.turma : 'undefined'}`);
        mismatches++;
        continue;
      }
      if (a.programa !== b.programa || a.turma !== b.turma || a.aulas.length !== b.aulas.length) {
        console.log(`#${i}: MISMATCH JS=[${a.programa}] "${a.turma}" (${a.aulas.length} aulas) vs PY=[${b.programa}] "${b.turma}" (${b.aulas.length} aulas)`);
        mismatches++;
        continue;
      }
      for (let k = 0; k < a.aulas.length; k++) {
        const aa = a.aulas[k], bb = b.aulas[k];
        const fields = ['data', 'dia', 'horario', 'disciplina', 'facilitador'];
        for (const f of fields) {
          if ((aa[f] || '') !== (bb[f] || '')) {
            console.log(`#${i} "${a.turma}" aula ${k} field ${f}: JS=${JSON.stringify(aa[f])} PY=${JSON.stringify(bb[f])}`);
            mismatches++;
          }
        }
      }
    }
    console.log('total mismatches:', mismatches);
  }

  if (MATRICULAS_XLSX) {
    const wb2 = loadWorkbook(MATRICULAS_XLSX);
    const matriculas = extractMatriculas(wb2);
    console.log('\n--- MATRICULAS ---');
    console.log('total chaves (JS):', Object.keys(matriculas).length);

    if (MATRICULAS_JSON_GROUND_TRUTH) {
      const gtm = JSON.parse(fs.readFileSync(MATRICULAS_JSON_GROUND_TRUTH, 'utf8'));
      console.log('total chaves (PY):', Object.keys(gtm).length);
      let mismatches = 0;
      const allKeys = new Set([...Object.keys(matriculas), ...Object.keys(gtm)]);
      for (const k of allKeys) {
        const a = matriculas[k], b = gtm[k];
        if (!a || !b) {
          console.log(`key ${k}: MISSING — JS=${!!a} PY=${!!b}`);
          mismatches++;
          continue;
        }
        for (const f of ['min', 'max', 'atual', 'faltam_para_minimo']) {
          if ((a[f] ?? null) !== (b[f] ?? null)) {
            console.log(`key ${k} field ${f}: JS=${a[f]} PY=${b[f]}`);
            mismatches++;
          }
        }
      }
      console.log('total mismatches:', mismatches);
    }
  }
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
