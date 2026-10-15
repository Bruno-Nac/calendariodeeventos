# -*- coding: utf-8 -*-
import json, re, glob, os
from python_calamine import CalamineWorkbook

DOWNLOADS = r'C:\Users\Bruno Nascimento\Downloads'

def find_latest(pattern):
    matches = glob.glob(os.path.join(DOWNLOADS, pattern))
    if not matches:
        raise FileNotFoundError(f'Nenhum arquivo encontrado em Downloads casando com {pattern!r}. Baixe uma cópia fresca do Controle de Alunos antes de rodar este script.')
    return max(matches, key=os.path.getmtime)

SRC = find_latest('Planilha_Master_Controle_Alunos_2026*.xlsx')
print('Usando arquivo:', SRC)

wb = CalamineWorkbook.from_path(SRC)
ws = wb.get_sheet_by_name('CONFIGURACOES')
data = ws.to_python()

headers = data[4][15:25]
rows = []
for r in data[5:]:
    seg = r[15:25]
    if not any(x not in (None, '') for x in seg):
        continue
    if seg[0] in (None, ''):
        continue
    rows.append(dict(zip(headers, seg)))

# FormacaoCodigo -> internal programa key used by the dashboard
CODE_TO_PROGRAMA = {
    'ABP': 'ABP',
    'B_SPRINT': 'Board Sprint',
    'BOARD_360': 'Board 360',
    'PFIC': 'PFIC',
    'PFCC': 'PFCC',
    'PFCC_MX': 'PFCC LATAM',
    'PFCC_CH': 'PFCC LATAM',
    'CES': 'CES',
    'LEAN_GOVERNANCE': 'Lean Governance',
}

def trailing_number(s):
    if not isinstance(s, str):
        return None
    m = re.search(r'(\d+)\s*$', s.strip())
    return int(m.group(1)) if m else None

matriculas = {}  # (programa, numero) -> dict
unmatched = []
for r in rows:
    programa = CODE_TO_PROGRAMA.get(r['FormacaoCodigo'])
    if not programa:
        continue
    num = trailing_number(r['Turma'])
    if num is None:
        unmatched.append(r)
        continue
    def num_or_none(v):
        return int(v) if isinstance(v, (int, float)) and v != '' else None
    key = (programa, num)
    matriculas[f'{programa}|{num}'] = {
        'min': num_or_none(r['MinAlunos']),
        'max': num_or_none(r['MaxAlunos']),
        'atual': num_or_none(r['InscritosAtivos']),
        'faltam_para_minimo': num_or_none(r['FaltamParaMinimo']),
        'dias_para_comecar': num_or_none(r['DiasParaComecar']),
        'modalidade_ctrl': r['Modalidade'],
        'turma_codigo': r['Turma'],
    }

print('total linhas mapeadas:', len(matriculas))
print('linhas ignoradas (sem numero / fora do escopo):', len(unmatched))
from collections import Counter
print(Counter(CODE_TO_PROGRAMA.get(r['FormacaoCodigo'], 'FORA_DE_ESCOPO') for r in rows))

with open('matriculas_data.json', 'w', encoding='utf-8') as f:
    json.dump(matriculas, f, ensure_ascii=False, indent=None, separators=(',', ':'))
print('salvo matriculas_data.json com', len(matriculas), 'chaves')
