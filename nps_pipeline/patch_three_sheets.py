#!/usr/bin/env python3
"""Патчит SHEETS[Факторы], SHEETS[Пульс_день], SHEETS[ЖК_детально] в HTML-снапшоте.
Плюс: подгрузка Graphik LCG @font-face, лёгкое обновление бренд-палитры.
Далее вызывается build_v2_patches.py — он подхватит новые SHEETS и пересоберёт XLSX.
"""
import re, json, os, sys

ROOT = '/home/user/workspace'
SRC = f'{ROOT}/nps_pipeline/_snapshots/09_46_v2.html'
FACTORS = json.load(open(f'{ROOT}/nps_pipeline/db/factors.json', encoding='utf-8'))
ZHK = json.load(open(f'{ROOT}/nps_pipeline/db/zhk_detail.json', encoding='utf-8'))
PULSE = json.load(open(f'{ROOT}/nps_pipeline/db/pulse.json', encoding='utf-8'))

html = open(SRC, encoding='utf-8').read()

# --- парсим объект SHEETS ---
m = re.search(r'const\s+SHEETS\s*=\s*', html)
start = m.end()
depth = 0; i = start
while i < len(html):
    c = html[i]
    if c == '{': depth += 1
    elif c == '}':
        depth -= 1
        if depth == 0: end = i + 1; break
    i += 1
SHEETS = json.loads(html[start:end])
print('SHEETS keys before:', list(SHEETS.keys()))

# ============ 1. ФАКТОРЫ ============
# Формат: [ФАКТОРЫ: заголовок], [сноска],
# затем повторяющиеся блоки:
#   [Название блока]
#   [подпись/сноска]
#   ['Сегмент', 'В базе', 'Доля', '', 'NPS', '% крит', '% нейтр', '% пром', 'Индекс', 'CSAT']
#   [сегмент, n, доля%, '', NPS%, ...]
def fmt_pct_sign(v):
    if v is None: return ''
    s = ('+' if v>0 else '') + f'{v:.1f}%'.replace('.',',').replace('-','−')
    return s
def fmt_pct(v):
    if v is None: return ''
    return f'{v:.1f}%'.replace('.',',').replace('-','−')

fact_rows = [
    ['ФАКТОРЫ: как разные группы влияют на общий NPS-2026'],
    ['Доли сегментов — от всех ответивших на вопрос. NPS/распределение П/Н/К считается только внутри сегмента. n<50 — гипотеза.'],
]

def add_block(title, sub, segs, order=None):
    fact_rows.append([title])
    if sub: fact_rows.append([sub])
    fact_rows.append(['Сегмент', 'В базе', 'Доля', '', 'NPS', '% критиков', '% нейтралов', '% промоутеров', '', ''])
    total = sum(s['n'] for s in segs)
    for s in segs:
        share = round(s['n']/total*100,1) if total else 0
        fact_rows.append([
            s['name'], s['n'], f'{share:.1f}%'.replace('.',','), '',
            fmt_pct_sign(s['nps']),
            fmt_pct(s['pct_c']), fmt_pct(s['pct_n']), fmt_pct(s['pct_p']),
            '', ''
        ])

add_block('Пол', None, FACTORS['gender'])
add_block('Возраст', None, FACTORS['age'])
add_block('Регион', None, FACTORS['region'])
add_block('Стаж жителя', None, FACTORS['tenure'])
add_block('Посещение офиса УК за 12 мес.', None, FACTORS['office_visit'])
if FACTORS.get('office_freq'):
    add_block('Частота визитов (среди «Да»)', None, FACTORS['office_freq'])
add_block('Обращения по гарантии', None, FACTORS['garant_ask'])
if FACTORS.get('garant_impact'):
    add_block('Как повлияла работа с гарантией', 'В зачёт — только по ответившим «Да» в вопросе о гарантии', FACTORS['garant_impact'])
# оценка процесса 1–5
gr = FACTORS.get('garant_rate') or {}
if gr.get('total'):
    fact_rows.append(['Оценка процесса решения гарантийных вопросов (шкала 1–5)'])
    fact_rows.append([f'Средний балл — {str(gr["avg"]).replace(".",",")}; CSAT (4–5) — {str(gr["csat45"]).replace(".",",")}%; n = {gr["total"]}'])
    fact_rows.append(['Сегмент','В базе','Доля','','','','','','',''])
    for k in ['1','2','3','4','5']:
        n = gr['dist'].get(k, 0)
        share = gr['dist_pct'].get(k, 0)
        fact_rows.append([f'{k} балл', n, f'{share:.1f}%'.replace('.',','), '', '', '', '', '', '', ''])
add_block('Мероприятия ПИК-Комфорт по группам ЖК', 'Группа А — программа максимум; B — программа минимум; С — без мероприятий.', FACTORS['event'])

SHEETS['Факторы'] = fact_rows

# ============ 2. ПУЛЬС_ДЕНЬ ============
# Добавляем колонку «Коммуникация» после NPS накоп.
pulse_rows = [
    ['ПУЛЬС ДЕНЬ NPS-2026'],
    ['Ежедневный прирост, накопительный NPS и коммуникации'],
    ['Дата','Ответов накоп.','Прирост за день','NPS накоп.','NPS дня','% критиков накоп.','% нейтралов накоп.','% промоутеров накоп.','Критики (n)','Нейтралы (n)','Промоутеры (n)','Коммуникация'],
]
prev_nps = None
for r in PULSE['rows']:
    pulse_rows.append([
        r['date'],
        r['n_cum'],
        r['n_day'],
        fmt_pct_sign(r['nps_cum']),
        fmt_pct_sign(r['nps_day']),
        fmt_pct(r['pct_c_cum']),
        fmt_pct(r['pct_n_cum']),
        fmt_pct(r['pct_p_cum']),
        r['crit_cum'],
        r['neut_cum'],
        r['promo_cum'],
        r['comm'],
    ])
SHEETS['Пульс_день'] = pulse_rows

# ============ 3. ЖК_ДЕТАЛЬНО ============
# Структура для showZhkDetailView:
#   row0 [Титул]
#   row1 [Сноска]
#   row2 [datesRow — пусто на базовых, дата над тройкой Отв/NPS/Инд]
#   row3 ['№','ЖК','Регион','Группа',...,'Отв','NPS','Инд',...]
#   row4+ data
dates = ZHK['dates']
# датасет со всеми деталями: базовые + групповые
base_cols = ['№','ЖК','Регион','Группа','Класс','Ответов','NPS накоп.','% критиков','% нейтралов','% промоутеров','NPS 2025','n 2025','NPS 2024','n 2024']
# распределение 0..10 — 11 колонок
dist_cols = [f'{i}' for i in range(11)]
# после datesRow дневные колонки Отв/NPS/Инд для каждой даты
dates_row = [''] * len(base_cols) + [''] * len(dist_cols)
for d in dates:
    dates_row += [d, d, d]  # Отв, NPS, Инд
header = base_cols + dist_cols + (['Отв','NPS','Инд']*len(dates))

zhk_rows = [
    ['ЖК ДЕТАЛЬНО — весь периметр (n<50 — маркирована ⚠ гипотеза)'],
    ['Сортировка кликом по шапке. Заморожены № / ЖК / Регион / Группа. Справа — распределение оценок 0-10 (%) и динамика по дням.'],
    dates_row,
    header,
]

def fmt_pct_or_dash(v):
    if v is None or v == '' or v == '—': return ''
    try:
        f = float(str(v).replace(',','.').replace('−','-').replace('%',''))
        return fmt_pct_sign(f)
    except: return str(v)

for idx, z in enumerate(ZHK['rows'], 1):
    name = z['name']
    if z['n'] < 50: name += ' ⚠'
    row = [
        idx,
        name,
        z['region'],
        z['event'] or 'Не в матрице',
        z['klass'] or '—',
        z['n'],
        fmt_pct_sign(z['nps']),
        fmt_pct(z['pct_c']),
        fmt_pct(z['pct_n']),
        fmt_pct(z['pct_p']),
        fmt_pct_or_dash(z.get('nps_2025')),
        z.get('n_2025') or '',
        fmt_pct_or_dash(z.get('nps_2024')),
        z.get('n_2024') or '',
    ]
    for i in range(11):
        p = z['dist'].get(str(i), z['dist'].get(i, 0))
        row.append(f'{p:.1f}%'.replace('.',',') if p else '')
    for d in dates:
        dd = z['daily'].get(d, {})
        n = dd.get('n', 0)
        nps = dd.get('nps')
        row.append(n if n else '')
        row.append(fmt_pct_sign(nps) if nps is not None else '')
        row.append('')  # Инд опыта дня — оставляю пустым (нет расчёта в json)
    zhk_rows.append(row)

SHEETS['ЖК_детально'] = zhk_rows

# ---- вставляем SHEETS обратно ----
new_sheets_js = json.dumps(SHEETS, ensure_ascii=False)
html = html[:start] + new_sheets_js + html[end:]

# ---- Стиль: шрифт Graphik LCG (fallback Inter) ----
font_face = """
<style id="graphik-font">
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
@font-face { font-family: 'Graphik LCG'; src: url('https://fonts.cdnfonts.com/s/16009/GraphikLCG-Regular.woff') format('woff'); font-weight: 400; font-display: swap; }
@font-face { font-family: 'Graphik LCG'; src: url('https://fonts.cdnfonts.com/s/16009/GraphikLCG-Medium.woff') format('woff'); font-weight: 500; font-display: swap; }
@font-face { font-family: 'Graphik LCG'; src: url('https://fonts.cdnfonts.com/s/16009/GraphikLCG-Semibold.woff') format('woff'); font-weight: 600; font-display: swap; }
body, .tab, .factor-card, .zhk-table, .sheet-title, h2, h3, table, td, th, button, input { font-family: 'Graphik LCG', 'Inter', system-ui, -apple-system, 'Segoe UI', sans-serif !important; }
.sheet-title, h2, h3 { font-weight: 600 !important; letter-spacing: -0.01em; }
</style>
"""
if 'id="graphik-font"' not in html:
    html = html.replace('</head>', font_face + '</head>')

with open(SRC, 'w', encoding='utf-8') as f:
    f.write(html)
print('HTML patched.')
print('  factors blocks rows:', len(fact_rows))
print('  pulse rows:', len(pulse_rows))
print('  zhk rows:', len(zhk_rows), 'cols in header:', len(header), 'dates:', dates)
