# -*- coding: utf-8 -*-
import json, re, datetime, os, glob
from python_calamine import CalamineWorkbook

BASE = r'C:\Users\Bruno Nascimento\Documents\Tabela Price\Regua de Turmas'
DOWNLOADS = r'C:\Users\Bruno Nascimento\Downloads'
TODAY = datetime.date.today().isoformat()
NOW_ISO = datetime.datetime.now().astimezone().isoformat(timespec='seconds')

def find_latest(pattern):
    matches = glob.glob(os.path.join(DOWNLOADS, pattern))
    return max(matches, key=os.path.getmtime) if matches else None

XLSX = find_latest('Cronograma 2026 - Turmas*.xlsx')

with open(os.path.join(BASE, 'turmas_data.json'), encoding='utf-8') as f:
    raw = json.load(f)

# internal programa key -> full display label requested by the user
PROGRAM_ORDER = ['PFCC', 'CES', 'Lean Governance', 'Board 360', 'ABP', 'PFCC LATAM', 'PFIC', 'Board Sprint']
PROGRAM_LABELS = {
    'PFCC': 'Programa de Formação e Certificação de Conselheiros - PFCC',
    'CES': 'Conselho Empresarial da Saúde - CES',
    'Lean Governance': 'Lean Governance',
    'Board 360': 'Board 360',
    'ABP': 'Advanced Board Program - ABP',
    'PFCC LATAM': 'Programa de Formación y Certificación de Consejeros Latam',
    'PFIC': 'Programa de Finanças para Conselheiros - PFIC',
    'Board Sprint': 'Board Sprint',
}

def modalidade_from_title(title):
    t = title.lower()
    if 'presencial' in t:
        return 'Presencial'
    if 'online' in t:
        return 'Online'
    return ''

def clean_title(title):
    return re.sub(r'\s+', ' ', title).strip()

EXCLUDE_PATTERNS = re.compile(r'cancelad|rascunho', re.IGNORECASE)

with open(os.path.join(BASE, 'matriculas_data.json'), encoding='utf-8') as f:
    matriculas = json.load(f)

# how to pull the turma number out of each program's title format, to join against
# the enrollment-control sheet (which keys turmas as "{codigo}_2026_{numero}")
NUM_PATTERNS = {
    'PFCC': re.compile(r'PFCC(\d+)'),
    'PFCC LATAM': re.compile(r'PFCC(\d+)'),
    'CES': re.compile(r'CES(\d+)'),
    'ABP': re.compile(r'ABP(\d+)'),
    'Board Sprint': re.compile(r'Turma (\d+)'),
    'Lean Governance': re.compile(r'Turma (\d+)'),
    'PFIC': re.compile(r'Turma (\d+)'),
    'Board 360': re.compile(r'^T(\d+)'),
}

def turma_numero(programa, titulo):
    pat = NUM_PATTERNS.get(programa)
    if not pat:
        return None
    m = pat.search(titulo)
    return int(m.group(1)) if m else None

def lookup_matricula(programa, titulo):
    num = turma_numero(programa, titulo)
    if num is None:
        return None
    return matriculas.get(f'{programa}|{num}')

turmas_out = []
skipped = []
matched_count = 0
for t in raw['turmas']:
    if EXCLUDE_PATTERNS.search(t['turma']):
        skipped.append(f"{t['programa']} | {t['turma']}")
        continue
    dates = sorted([a['data'] for a in t['aulas'] if a['data']])
    if dates:
        data_inicio, data_fim = dates[0], dates[-1]
        if data_inicio >= TODAY:
            status = 'futura'
        elif data_fim < TODAY:
            status = 'concluida'
        else:
            status = 'andamento'
    else:
        data_inicio = data_fim = None
        status = 'sem_data'

    aulas_sorted = sorted(t['aulas'], key=lambda a: (a['data'] or '9999-99-99'))
    mat = lookup_matricula(t['programa'], t['turma'])
    matriculas_out = None
    if mat and (mat['min'] is not None or mat['max'] is not None or mat['atual'] is not None):
        matched_count += 1
        atual = mat['atual'] or 0
        minimo = mat['min']
        maximo = mat['max']
        if minimo is not None and maximo is not None and (minimo > 0 or maximo > 0):
            cap_status = 'atingido' if atual >= (minimo or 0) else 'abaixo'
        else:
            cap_status = 'sem_config'
        matriculas_out = {
            'min': minimo,
            'max': maximo,
            'atual': mat['atual'],
            'faltam_para_minimo': mat['faltam_para_minimo'],
            'status_capacidade': cap_status,
        }
    turmas_out.append({
        'id': re.sub(r'[^a-zA-Z0-9]+', '-', f"{t['programa']}-{t['turma']}").strip('-').lower(),
        'programa': t['programa'],
        'turma': clean_title(t['turma']),
        'modalidade': modalidade_from_title(t['turma']),
        'status': status,
        'data_inicio': data_inicio,
        'data_fim': data_fim,
        'total_aulas': len(t['aulas']),
        'matriculas': matriculas_out,
        'aulas': [
            {
                'data': a['data'],
                'data_raw': a['data_raw'],
                'dia': a['dia'],
                'horario': a['horario'],
                'disciplina': a['disciplina'],
                'facilitador': a['facilitador'],
            }
            for a in aulas_sorted
        ],
    })

def prog_key(p):
    try:
        return PROGRAM_ORDER.index(p)
    except ValueError:
        return len(PROGRAM_ORDER)

turmas_out.sort(key=lambda t: (prog_key(t['programa']), t['data_inicio'] or '9999'))

# ---------------- Encontros de Comunidades: parsed live from the sheet ----------------
MESES = {
    'janeiro': 1, 'fevereiro': 2, 'março': 3, 'marco': 3, 'abril': 4, 'maio': 5, 'junho': 6,
    'julho': 7, 'agosto': 8, 'setembro': 9, 'outubro': 10, 'novembro': 11, 'dezembro': 12
}

def parse_evt_date(v):
    if isinstance(v, datetime.date):
        return v.isoformat()
    if isinstance(v, str):
        m = re.search(r'(\d{1,2})\s*de\s*([a-zçãéêô]+)\s*de\s*(\d{4})', v, re.IGNORECASE)
        if m:
            day, mes_name, year = int(m.group(1)), m.group(2).lower(), int(m.group(3))
            mes = MESES.get(mes_name)
            if mes:
                try:
                    return datetime.date(year, mes, day).isoformat()
                except ValueError:
                    return None
    return None

EVENTOS_CUTOFF = (datetime.date.today().replace(day=1)).isoformat()

if XLSX:
    wb = CalamineWorkbook.from_path(XLSX)
    ws = wb.get_sheet_by_name('Encontros Comunidades')
    rows = ws.to_python()

    def norm(v):
        if v is None:
            return ''
        return v.strip() if isinstance(v, str) else v

    # series section titles are rows with exactly one non-empty cell
    sections = []
    for i, row in enumerate(rows):
        nonempty = [c for c in row if norm(c) != '']
        if len(nonempty) == 1 and isinstance(nonempty[0], str) and norm(row[0]) == nonempty[0]:
            sections.append((i, norm(row[0])))

    eventos = []
    for idx, (row_i, serie) in enumerate(sections):
        header_i = row_i + 1
        headers = [norm(c) for c in rows[header_i]]
        end_i = sections[idx + 1][0] if idx + 1 < len(sections) else len(rows)
        for r in range(header_i + 1, end_i):
            row = rows[r]
            vals = dict(zip(headers, [norm(c) for c in row]))
            data_val = vals.get('Data')
            if data_val == '' or data_val is None:
                continue
            iso = parse_evt_date(data_val)
            if not iso or iso < EVENTOS_CUTOFF:
                continue
            eventos.append({
                'serie': serie,
                'data': iso,
                'status': vals.get('Status') or '',
                'convidado': vals.get('Convidado') or '',
                'tema': vals.get('Tema') or '',
                'programas': vals.get('Programas que participam') or '',
            })
else:
    # source workbook not found in Downloads this run; reuse events from the last build
    with open(os.path.join(BASE, 'dashboard_data.json'), encoding='utf-8') as f:
        eventos = json.load(f)['eventos_comunidades']
    print('aviso: xlsx do cronograma nao encontrado em Downloads; reaproveitando eventos_comunidades do build anterior')

# Board Select: requested as a 4th community series; not always present in the source sheet.
COMUNIDADES_SERIES = ['Board Select', 'Lean Board na Prática', 'Health Board Meeting', 'AI Business Community']
eventos.sort(key=lambda e: (e['data'], COMUNIDADES_SERIES.index(e['serie']) if e['serie'] in COMUNIDADES_SERIES else 9))

out = {
    'gerado_em': NOW_ISO,
    'fonte_arquivo': 'Cronograma 2026 - Turmas.xlsx',
    'program_order': PROGRAM_ORDER,
    'program_labels': PROGRAM_LABELS,
    'comunidades_series': COMUNIDADES_SERIES,
    'turmas': turmas_out,
    'eventos_comunidades': eventos,
}

dashboard_data_path = os.path.join(BASE, 'dashboard_data.json')
with open(dashboard_data_path, 'w', encoding='utf-8') as f:
    json.dump(out, f, ensure_ascii=False, indent=None, separators=(',', ':'))

from collections import Counter
status_counts = Counter(t['status'] for t in turmas_out)
prog_counts = Counter(t['programa'] for t in turmas_out)
print('status:', dict(status_counts))
print('por programa:', dict(prog_counts))
print('total turmas:', len(turmas_out))
print('total eventos:', len(eventos))
print('series encontradas na planilha:', sorted(set(e['serie'] for e in eventos)))
print('turmas excluidas (canceladas/rascunho):', len(skipped))
print('turmas com dados de matricula cruzados:', matched_count, '/', len(turmas_out))
print('tamanho dashboard_data.json:', os.path.getsize(dashboard_data_path), 'bytes')

# ---------------- Inject into the HTML template to produce the final artifact file ----------------
data_json = json.dumps(out, ensure_ascii=False, separators=(',', ':'))
data_json_safe = data_json.replace('</script', '<\\/script').replace('<!--', '<\\!--')

template_path = os.path.join(BASE, 'dashboard_template.html')
with open(template_path, encoding='utf-8') as f:
    template = f.read()

# keep the JS-side TODAY constant in sync with the python-computed TODAY used above
template = re.sub(r'var TODAY = "[\d-]+";', f'var TODAY = "{TODAY}";', template, count=1)

assert '__DATA_PLACEHOLDER__' in template, 'placeholder not found in template — did dashboard_template.html change?'
final_html = template.replace('__DATA_PLACEHOLDER__', data_json_safe)

final_path = os.path.join(BASE, 'regua_de_turmas.html')
with open(final_path, 'w', encoding='utf-8') as f:
    f.write(final_html)
print('HTML final gerado em', final_path, '(', os.path.getsize(final_path), 'bytes )')

# ---------------- Clean up only the specific downloads this run consumed ----------------
# (never mass-delete every matching file in Downloads — older matches may be the
# user's own unrelated copies that this pipeline didn't touch. Extra safety: only
# remove a file if it was modified in the last 30 minutes, i.e. plausibly downloaded
# by THIS run, never a stale pre-existing copy that merely happened to be "latest".)
RECENCY_WINDOW_SECONDS = 30 * 60
matriculas_xlsx = find_latest('Planilha_Master_Controle_Alunos_2026*.xlsx')
cleaned = []
skipped_stale = []
now = datetime.datetime.now().timestamp()
for p in (XLSX, matriculas_xlsx):
    if not p:
        continue
    age = now - os.path.getmtime(p)
    if age > RECENCY_WINDOW_SECONDS:
        skipped_stale.append(p)
        continue
    try:
        os.remove(p)
        cleaned.append(p)
    except OSError as e:
        print('nao foi possivel remover', p, ':', e)
print('arquivos temporarios removidos de Downloads:', cleaned)
if skipped_stale:
    print('arquivos NAO removidos por serem antigos demais (nao baixados nesta execucao):', skipped_stale)
