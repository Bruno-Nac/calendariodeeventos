const assert = require('assert');
const { normalizeManualTurma, sortTurmas } = require('../lib/buildDashboard');

const TODAY = '2026-09-20';

// New turma, future dates, full matriculas
const t1 = normalizeManualTurma({
  programa: 'PFCC',
  turma: 'PFCC200 Online - Turma de Segunda',
  aulas: [
    { data: '2026-10-05', horario: '19h-21h', disciplina: 'Fundamentos', facilitador: 'Fulano' },
    { data: '2026-10-12', horario: '19h-21h', disciplina: 'Estratégia', facilitador: 'Fulano' },
  ],
  matriculas: { min: 24, max: 40, atual: 10 },
}, TODAY);

assert.strictEqual(t1.id, 'pfcc-pfcc200-online-turma-de-segunda');
assert.strictEqual(t1.status, 'futura');
assert.strictEqual(t1.data_inicio, '2026-10-05');
assert.strictEqual(t1.data_fim, '2026-10-12');
assert.strictEqual(t1.modalidade, 'Online');
assert.strictEqual(t1.total_aulas, 2);
assert.strictEqual(t1.matriculas.status_capacidade, 'abaixo');
assert.strictEqual(t1.matriculas.faltam_para_minimo, 14);
console.log('t1 ok:', JSON.stringify(t1, null, 2));

// No aulas at all -> sem_data
const t2 = normalizeManualTurma({ programa: 'ABP', turma: 'ABP99 | Online', aulas: [] }, TODAY);
assert.strictEqual(t2.status, 'sem_data');
assert.strictEqual(t2.matriculas, null);
console.log('t2 ok');

// Past dates -> concluida; matriculas atingido
const t3 = normalizeManualTurma({
  programa: 'CES',
  turma: 'CES99 - Presencial',
  aulas: [{ data: '2020-01-01', horario: '', disciplina: 'X', facilitador: 'Y' }],
  matriculas: { min: 10, max: 20, atual: 15 },
}, TODAY);
assert.strictEqual(t3.status, 'concluida');
assert.strictEqual(t3.matriculas.status_capacidade, 'atingido');
console.log('t3 ok');

// Sorting: PFCC before CES before Lean Governance (per PROGRAM_ORDER), and within
// same program, earlier data_inicio first
const arr = [t3, t1];
sortTurmas(arr);
assert.strictEqual(arr[0].programa, 'PFCC'); // PFCC(0) before CES(1) in PROGRAM_ORDER
console.log('sort ok:', arr.map((t) => t.programa));

// id collision resolution simulation (mirrors what api/turmas.js does)
const existing = [{ id: 'pfcc-pfcc200-online-turma-de-segunda' }];
let id = t1.id, n = 2;
while (existing.some((t) => t.id === id)) id = `${t1.id}-${n++}`;
assert.strictEqual(id, 'pfcc-pfcc200-online-turma-de-segunda-2');
console.log('collision ok:', id);

console.log('\nALL PASS');
