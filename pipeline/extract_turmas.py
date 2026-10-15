# -*- coding: utf-8 -*-
import re, json, datetime, glob, os
from python_calamine import CalamineWorkbook

DOWNLOADS = r'C:\Users\Bruno Nascimento\Downloads'
OUT = r'C:\Users\Bruno Nascimento\Documents\Tabela Price\Regua de Turmas\turmas_data.json'

def find_latest(pattern):
    matches = glob.glob(os.path.join(DOWNLOADS, pattern))
    if not matches:
        raise FileNotFoundError(f'Nenhum arquivo encontrado em Downloads casando com {pattern!r}. Baixe uma cópia fresca do Cronograma antes de rodar este script.')
    return max(matches, key=os.path.getmtime)

SRC = find_latest('Cronograma 2026 - Turmas*.xlsx')
print('Usando arquivo:', SRC)

wb = CalamineWorkbook.from_path(SRC)

MESES = {
    'janeiro': 1, 'fevereiro': 2, 'março': 3, 'marco': 3, 'abril': 4, 'maio': 5, 'junho': 6,
    'julho': 7, 'agosto': 8, 'setembro': 9, 'outubro': 10, 'novembro': 11, 'dezembro': 12
}

def norm(v):
    if v is None:
        return ''
    if isinstance(v, str):
        return v.strip()
    return v

def parse_date(v):
    """Return (iso_date_str_or_None, weekday_str_or_None)."""
    if v is None or v == '':
        return None, None
    if isinstance(v, datetime.date):
        return v.isoformat(), None
    if isinstance(v, datetime.datetime):
        return v.date().isoformat(), None
    if isinstance(v, str):
        s = v.strip()
        # pattern: "terça-feira, 15 de janeiro de 2024" or "15 de janeiro de 2024"
        m = re.search(r'(\d{1,2})\s*de\s*([a-zçãéêô]+)\s*de\s*(\d{4})', s, re.IGNORECASE)
        if m:
            day = int(m.group(1))
            mes_name = m.group(2).lower()
            mes = MESES.get(mes_name)
            year = int(m.group(3))
            if mes:
                try:
                    return datetime.date(year, mes, day).isoformat(), s.split(',')[0].strip() if ',' in s else None
                except ValueError:
                    return None, s
        # pattern: DD/MM/YYYY
        m = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{2,4})$', s)
        if m:
            d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
            if y < 100:
                y += 2000
            try:
                return datetime.date(y, mo, d).isoformat(), None
            except ValueError:
                return None, s
        if s.upper() in ('XXXX', 'TBD', ''):
            return None, None
        return None, s  # unparsed text, keep raw
    return None, None

def find_blocks(rows):
    blocks = []
    for r, row in enumerate(rows):
        c = 0
        while c < len(row):
            val = norm(row[c])
            if isinstance(val, str) and val.lower() in ('data',):
                headers = []
                cc = c
                while cc < len(row):
                    v = norm(row[cc])
                    if v == '':
                        nxt = norm(row[cc + 1]) if cc + 1 < len(row) else ''
                        if nxt == '' or (isinstance(nxt, str) and nxt.lower() == 'data'):
                            break
                        headers.append('')
                        cc += 1
                        continue
                    if isinstance(v, str) and v.lower() == 'data' and cc != c:
                        break
                    headers.append(v)
                    cc += 1
                blocks.append({'row': r, 'col': c, 'width': cc - c, 'headers': headers})
                c = cc
            else:
                c += 1
    return blocks

def col_range_overlaps(a, b):
    a_start, a_end = a['col'], a['col'] + max(a['width'], 6)
    b_start, b_end = b['col'], b['col'] + max(b['width'], 6)
    return not (a_end <= b_start or b_end <= a_start)

def extract_sheet(sheet_name, program_label):
    ws = wb.get_sheet_by_name(sheet_name)
    rows = ws.to_python()
    nrows = len(rows)
    blocks = find_blocks(rows)
    turmas = []

    for i, blk in enumerate(blocks):
        r, c, headers = blk['row'], blk['col'], blk['headers']
        title = ''
        for back in (1, 2, 3):
            rr_up = r - back
            if rr_up < 0:
                break
            row_up = rows[rr_up]
            for cc in range(max(0, c - 2), min(len(row_up), c + len(headers) + 4)):
                cand = norm(row_up[cc])
                if cand and not re.match(r'^https?://', cand, re.IGNORECASE):
                    title = cand
                    break
            if title:
                break
        if not title:
            title = None  # mark as unresolved fragment; dropped later

        # find end row: nearest later block that overlaps in columns; else nearest later block at all (any column);
        # else end of sheet. This avoids one block swallowing unrelated rows far below it.
        end_row = nrows
        fallback_row = None
        for other in blocks[i+1:]:
            if other['row'] > r:
                if fallback_row is None:
                    fallback_row = other['row']
                if col_range_overlaps(blk, other):
                    end_row = other['row'] - 1
                    break
        else:
            if fallback_row is not None:
                end_row = fallback_row - 1

        col_start = c
        col_end = c + len(headers)

        aulas = []
        blank_streak = 0
        last_date_iso = None
        last_date_raw = None
        rr = r + 1
        while rr < min(end_row, nrows):
            row = rows[rr]
            segment = [norm(row[cc]) if cc < len(row) else '' for cc in range(col_start, col_end)]
            if all(x == '' for x in segment):
                blank_streak += 1
                if blank_streak >= 2:
                    break
                rr += 1
                continue
            blank_streak = 0
            rowdict = dict(zip(headers, segment))
            raw_date = rowdict.get('Data', '')
            iso, raw_txt = parse_date(raw_date) if raw_date != '' else (None, None)
            if iso:
                last_date_iso = iso
                last_date_raw = raw_date if isinstance(raw_date, str) else None
            elif raw_txt:
                last_date_raw = raw_txt
                last_date_iso = None
            # skip rows that are pure legend/noise (no discipline & no facilitator)
            disciplina = rowdict.get('Disciplina') or rowdict.get('Disciplinas') or ''
            facilitador = rowdict.get('Facilitador') or ''
            horario = rowdict.get('Horário') or rowdict.get('Horario') or ''
            if disciplina or facilitador or horario:
                aulas.append({
                    'data': iso if (raw_date != '') else last_date_iso,
                    'data_raw': (raw_date if isinstance(raw_date, str) and raw_date else last_date_raw),
                    'dia': rowdict.get('Dia') or '',
                    'horario': horario,
                    'disciplina': disciplina,
                    'facilitador': facilitador,
                    'confirmacao': rowdict.get('Confirmação') or '',
                    'mediador': rowdict.get('Mediador') or '',
                })
            rr += 1

        if aulas and title:
            turmas.append({
                'programa': program_label,
                'turma': title,
                'aulas': aulas,
            })

    return turmas

TARGETS = [
    ('PFIC ', 'PFIC'),
    ('PFCC', 'PFCC'),
    ('PFCC LATAM', 'PFCC LATAM'),
    ('ABP', 'ABP'),
    ('Lean Governance', 'Lean Governance'),
    ('Board 360', 'Board 360'),
    ('CES', 'CES'),
    ('Sprint', 'Board Sprint'),
]

def main():
    all_turmas = []
    errors = []
    for sheet_name, label in TARGETS:
        try:
            t = extract_sheet(sheet_name, label)
            print(f'{label}: {len(t)} turmas, {sum(len(x["aulas"]) for x in t)} aulas')
            all_turmas.extend(t)
        except Exception as e:
            errors.append(f'{label}: {e}')
            print(f'{label}: ERRO {e}')

    with open(OUT, 'w', encoding='utf-8') as f:
        json.dump({'turmas': all_turmas, 'errors': errors}, f, ensure_ascii=False, indent=2, default=str)

    print('---')
    print('Total turmas:', len(all_turmas))
    print('Salvo em', OUT)

if __name__ == '__main__':
    main()
