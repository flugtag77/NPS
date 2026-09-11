#!/usr/bin/env python3
"""Патч: добавляет колонку 'Коммуникация' в SHEETS['Пульс_день'] уже собранного HTML.
Работает поверх build_v2_patches — вызывается ПОСЛЕ него."""
import re, json
SRC = '/home/user/workspace/nps_pipeline/_snapshots/09_46_v2.html'
COMMS = {
    '07.09.2026': 'Рассылка в MAX и TG каналах',
    '08.09.2026': 'Пост-МиП с праздников с ссылкой на NPS',
    '09.09.2026': 'PUSH в ЛКК: Петр 1, Лефортово парк, Ясеневая, Люберцы парк, Митино парк',
    '10.09.2026': 'PUSH в ЛКК: Второй Нагатинский, Ильинские луга, Белая дача парк, Жемчужина Зеленограда, Академика Павлова, Открытый парк, Одинцово 1, Никольские луга, Черняховского 19, Зеленый парк',
    '11.09.2026': 'PUSH в ЛК: Красноказарменная 15, Столичные поляны, Середневский лес, Green park, Измайловский лес, Мякинино парк, Новохохловская 15, Барклая 6, Кузьминский лес, Кольская 8',
}
html = open(SRC, encoding='utf-8').read()
m = re.search(r'const\s+SHEETS\s*=\s*', html); s=m.end(); d=0; i=s
while i<len(html):
    c=html[i]
    if c=='{': d+=1
    elif c=='}':
        d-=1
        if d==0: e=i+1; break
    i+=1
S = json.loads(html[s:e])
pulse = S['Пульс_день']
# Сначала чистим все уже вписанные 'Коммуникация'-хвосты (защита от повторного вызова)
COMMS_VALS = set(COMMS.values()) | {''}
cleaned = []
for row in pulse:
    if isinstance(row, list):
        # чистим все хвосты 'Коммуникация' в заголовке
        while row and row[-1] == 'Коммуникация':
            row = row[:-1]
        # в датовых строках: срезаем все хвосты, которые есть в наборе COMMS-текстов
        if isinstance(row[0] if row else None, str) and re.match(r'\d\d\.\d\d\.\d{4}', row[0]):
            while row and isinstance(row[-1], str) and row[-1] in COMMS_VALS:
                row = row[:-1]
    cleaned.append(row)
pulse = cleaned
new = []
header_seen = False
for row in pulse:
    if not isinstance(row, list) or not row:
        new.append(row); continue
    if row[0] == 'Дата' and not header_seen:
        header_seen = True
        new.append(row + ['Коммуникация'])
    elif header_seen and isinstance(row[0], str) and re.match(r'\d\d\.\d\d\.\d{4}', row[0]):
        new.append(row + [COMMS.get(row[0], '')])
    else:
        new.append(row)
S['Пульс_день'] = new
html = html[:s] + json.dumps(S, ensure_ascii=False) + html[e:]
open(SRC, 'w', encoding='utf-8').write(html)
print('OK: колонка Коммуникация добавлена, строк в Пульс_день:', len(new))
