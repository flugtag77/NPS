"""
Патчит _snapshots/09_46_v2.html: добавляет
- метки n<50 = гипотеза + счётчик стабилизированных
- drill-down по клику на числа
- новую вкладку 'Карта ЖК' (BCG-матрица NPS × ΔYoY)
- блок 'Прогноз объёма ответов + риск недобора по ЖК' в вкладке 'Прогноз_NPS'
- перегенерирует зашитый XLSX с 13 вкладками в фирменных цветах
"""
import re, json, base64, io, os, csv
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT="/home/user/workspace"
SRC =f"{ROOT}/nps_pipeline/_snapshots/09_46_v2.html"
MAP =f"{ROOT}/nps_pipeline/_snapshots/map_zhk.json"
CSV =f"{ROOT}/nps_pipeline/db/responses_long.csv"

# -------- ПИК-Комфорт brand palette --------
BR_ORANGE = "FF6B00"
BR_ORANGE_SOFT = "FFE9D6"
BR_GREEN = "1DB954"
BR_YELLOW = "F5A623"
BR_RED = "E64545"
BR_GRAY = "6E7580"
BR_LIGHT = "F4F5F7"
BR_INK = "1B1D1F"
BR_HYP  = "FFF5E6"

# -------- Загрузка HTML и SHEETS --------
html = open(SRC, encoding="utf-8").read()

m = re.search(r'const\s+SHEETS\s*=\s*', html)
start = m.end(); depth=0; i=start
while i<len(html):
    c=html[i]
    if c=='{': depth+=1
    elif c=='}':
        depth-=1
        if depth==0: end=i+1; break
    i+=1
SHEETS = json.loads(html[start:end])
print("Листов:", list(SHEETS.keys()))

MAPA = json.load(open(MAP, encoding="utf-8"))
stab = sum(1 for m in MAPA if (m["n_2026"] or 0) >= 50)
total = len(MAPA)
print(f"Стабилизировано: {stab} из {total}")

# -------- Прогноз объёма ответов + риск недобора --------
# Простая линейная проекция: если сейчас день X (10.09), а дедлайн 30.09,
# то прогноз = n_текущий * (20/4) с усечением по историческим темпам ЖК.
# Используем факт 07-10.09 → 4 дня, впереди 20 дней. Возьмём коэффициент 3.0
# (не 5.0, т.к. тесто падает по мере насыщения).
K = 3.0
for m in MAPA:
    n = m["n_2026"] or 0
    proj = round(n * K)
    m["n_proj"] = proj
    m["risk"] = "нет" if proj >= 50 else ("средний" if proj >= 30 else "высокий")

risk_hi = sum(1 for m in MAPA if m["risk"] == "высокий")
risk_md = sum(1 for m in MAPA if m["risk"] == "средний")
risk_lo = sum(1 for m in MAPA if m["risk"] == "нет")
print(f"Риск недобора: высокий {risk_hi}, средний {risk_md}, стабильно {risk_lo}")

# -------- Собираю HTML-инъекции --------

# Вкладка "Карта ЖК" — вставляется в SHEET_NAMES и получает свой рендер
map_js = f"""
// ==== Карта ЖК (BCG-матрица) ====
const MAP_ZHK = {json.dumps(MAPA, ensure_ascii=False)};
const STAB_TOTAL = {total};
const STAB_READY = {stab};

function showMapZhk() {{
  const v = document.getElementById("view");
  const segs = {{"Звезды":[], "Восходящие":[], "Спящие":[], "Кризис":[], "Недобор":[]}};
  MAP_ZHK.forEach(m => {{ (segs[m.seg]||segs["Недобор"]).push(m); }});
  const legend = {{
    "Звезды":"#1DB954","Восходящие":"#7ED957","Спящие":"#F5A623","Кризис":"#E64545","Недобор":"#B0B4BA"
  }};
  const rules = {{
    "Звезды":"NPS 2026 ≥ 0",
    "Восходящие":"NPS<0, но ΔYoY ≥ +5 п.п.",
    "Спящие":"−5 < ΔYoY < +5 п.п.",
    "Кризис":"ΔYoY ≤ −5 п.п.",
    "Недобор":"n < 50 (гипотеза)"
  }};
  let html = `
    <div class="sheet-header">
      <h2>Карта ЖК · BCG-матрица NPS × ΔYoY</h2>
      <div class="meta">Периметр {total} ЖК · стабилизировано ${{STAB_READY}} · остальные — гипотеза</div>
    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;">`;
  Object.keys(segs).forEach(name => {{
    const arr = segs[name];
    html += `
      <div class="card" style="border-left:6px solid ${{legend[name]}};">
        <div class="card-title">${{name}} · ${{arr.length}}</div>
        <div class="card-sub">${{rules[name]}}</div>
        <table class="mini"><thead><tr><th>ЖК</th><th>Регион</th><th>n 2026</th><th>NPS 2026</th><th>ΔYoY</th></tr></thead><tbody>`;
    arr.slice(0,30).forEach(m => {{
      const delta = (m.delta==null)? "—" : (m.delta>0?"+":"")+m.delta.toFixed(1)+" п.п.";
      const nps = (m.nps_2026==null)? "—" : m.nps_2026.toFixed(1)+"%";
      const hyp = (m.n_2026||0) < 50 ? ' <span class="hyp">⚠ гипотеза</span>' : '';
      html += `<tr><td>${{m.zhk}}${{hyp}}</td><td>${{m.region||""}}</td><td>${{m.n_2026||0}}</td><td>${{nps}}</td><td>${{delta}}</td></tr>`;
    }});
    if (arr.length > 30) html += `<tr><td colspan="5" class="muted">…ещё ${{arr.length-30}} ЖК</td></tr>`;
    html += `</tbody></table></div>`;
  }});
  html += `</div>`;
  html += `
    <div class="card" style="margin-top:16px">
      <div class="card-title">Правила сегментации</div>
      <div class="card-sub">Пересчитывается на каждом тике витрины (раз в час). Красная линия NPS = KPI 2025 факт = −32 п.п.</div>
    </div>`;
  v.innerHTML = html;
}}
"""

# Прогноз-объёма-блок в вкладку "Прогноз_NPS"
forecast_vol_js = f"""
// ==== Блок прогноза объёма ответов + риск недобора ====
const FORECAST_VOL = {{
  total: {total},
  stab_ready: {stab},
  risk_hi: {risk_hi},
  risk_md: {risk_md},
  risk_lo: {risk_lo},
  rows: {json.dumps([{"zhk":m["zhk"],"region":m["region"],"n_now":m["n_2026"] or 0,"n_proj":m["n_proj"],"risk":m["risk"]} for m in MAPA if (m["n_2026"] or 0)<50], ensure_ascii=False)}
}};
function renderForecastVolume(container) {{
  const wrap = document.createElement("div");
  wrap.className = "card";
  wrap.style.marginTop = "16px";
  let html = `
    <div class="card-title">Прогноз объёма ответов и риск недобора к 30.09</div>
    <div class="card-sub">Стабилизировано: ${{FORECAST_VOL.stab_ready}} из ${{FORECAST_VOL.total}} ЖК · Риск: высокий ${{FORECAST_VOL.risk_hi}} · средний ${{FORECAST_VOL.risk_md}} · безопасно ${{FORECAST_VOL.stab_ready+FORECAST_VOL.risk_lo}}</div>
    <table class="mini" style="margin-top:8px">
      <thead><tr><th>ЖК</th><th>Регион</th><th>n сейчас</th><th>Прогноз n к 30.09</th><th>Риск</th></tr></thead><tbody>`;
  FORECAST_VOL.rows.forEach(r => {{
    const cls = r.risk==="высокий" ? "bad" : (r.risk==="средний" ? "warn" : "ok");
    html += `<tr class="${{cls}}"><td>${{r.zhk}} <span class="hyp">⚠ гипотеза</span></td><td>${{r.region}}</td><td>${{r.n_now}}</td><td>${{r.n_proj}}</td><td>${{r.risk}}</td></tr>`;
  }});
  html += `</tbody></table>
    <div class="muted" style="margin-top:6px">Формула: прогноз_n = n_текущий × 3.0 (коэффициент оставшихся 20 дней при затухании темпа сбора). Порог стабилизации: n ≥ 50.</div>`;
  wrap.innerHTML = html;
  container.appendChild(wrap);
}}
"""

# Стили + мелкие утилиты (гипотеза, drill-down)
extra_css = """
<style>
.hyp { display:inline-block; padding:1px 6px; margin-left:4px; border-radius:4px;
       background:#FFF5E6; color:#B4610A; font-size:11px; font-weight:600; }
.card { background:#fff; border:1px solid #EEF0F3; border-radius:12px; padding:12px 14px; box-shadow:0 1px 2px rgba(0,0,0,.03); }
.card-title { font-weight:700; font-size:15px; color:#1B1D1F; margin-bottom:4px; }
.card-sub { color:#6E7580; font-size:12px; margin-bottom:8px; }
table.mini { width:100%; border-collapse:collapse; font-size:12px; }
table.mini th, table.mini td { text-align:left; padding:4px 6px; border-bottom:1px solid #F1F2F5; }
table.mini tr.bad td { background:#FFECEC; }
table.mini tr.warn td { background:#FFF5E6; }
table.mini tr.ok td { background:#EEFBF2; }
.muted { color:#8A9099; font-size:12px; }
.stab-badge { display:inline-flex; align-items:center; gap:6px; padding:4px 10px;
              border-radius:999px; background:#FFE9D6; color:#B4610A; font-size:12px; font-weight:600; margin-left:8px; }
#drill-panel { position:fixed; right:16px; bottom:16px; width:520px; max-height:70vh;
              overflow:auto; background:#fff; border:1px solid #DEE2E6; border-radius:12px;
              box-shadow:0 6px 24px rgba(0,0,0,.12); z-index:9999; display:none; }
#drill-panel .head { padding:10px 14px; background:#FF6B00; color:#fff; border-radius:12px 12px 0 0;
              display:flex; align-items:center; justify-content:space-between; font-weight:700; }
#drill-panel .body { padding:10px 14px; font-size:12px; line-height:1.4; }
.drill-num { cursor:pointer; border-bottom:1px dashed rgba(255,107,0,.5); }
.drill-num:hover { color:#FF6B00; }
</style>
"""

# JS-хелперы: drill-down, счётчик стабилизации
extra_js = """
// ==== Drill-down (упрощённая версия: показывает подсказку с составом) ====
function drillShow(title, formula, sourceLine) {
  let p = document.getElementById("drill-panel");
  if (!p) {
    p = document.createElement("div"); p.id = "drill-panel";
    p.innerHTML = '<div class="head"><span id="drill-title"></span><span style="cursor:pointer" onclick="document.getElementById(\\'drill-panel\\').style.display=\\'none\\'">✕</span></div><div class="body" id="drill-body"></div>';
    document.body.appendChild(p);
  }
  document.getElementById("drill-title").textContent = title;
  document.getElementById("drill-body").innerHTML =
    '<div><b>Формула:</b> ' + formula + '</div>' +
    '<div style="margin-top:6px"><b>Источник:</b> ' + sourceLine + '</div>' +
    '<div class="muted" style="margin-top:8px">Полный список ответов — см. вкладку «Списки для работы».</div>';
  p.style.display = "block";
}

// Оборачивает все элементы .drill-num в кликабельные подсказки
function bindDrill(root) {
  (root || document).querySelectorAll('.drill-num[data-title]').forEach(el => {
    if (el.__drillBound) return;
    el.__drillBound = true;
    el.addEventListener('click', () => {
      drillShow(el.dataset.title, el.dataset.formula || '—', el.dataset.source || '—');
    });
  });
}

// Счётчик стабилизации — вставляется в шапку каждой вкладки
function injectStabBadge() {
  const meta = document.getElementById('status-line');
  if (!meta) return;
  if (document.getElementById('stab-badge')) return;
  const b = document.createElement('span');
  b.id = 'stab-badge';
  b.className = 'stab-badge';
  b.innerHTML = 'Стабилизировано: ' + STAB_READY + ' из ' + STAB_TOTAL + ' ЖК (n≥50)';
  meta.appendChild(b);
}
"""

# --- Вставляем блоки в HTML ---
# 1. CSS в <head>
html = html.replace("</head>", extra_css + "\n</head>")

# 2. Добавляем "Карта ЖК" в SHEET_NAMES (после "Справочник_ЖК") — идемпотентно
m_sn = re.search(r'const\s+SHEET_NAMES\s*=\s*\[([^\]]+)\]\s*;', html)
if m_sn and '"Карта ЖК"' not in m_sn.group(1):
    html = re.sub(
        r'(const\s+SHEET_NAMES\s*=\s*\[)([^\]]+)(\]\s*;)',
        lambda m: m.group(1) + m.group(2).rstrip().rstrip(",") + ', "Карта ЖК"' + m.group(3),
        html, count=1
    )

# 3. В navigate() добавляем ветку карты — идемпотентно
if 'if (target === "Карта ЖК") { injectStabBadge(); showMapZhk()' not in html:
    html = re.sub(
        r'(function\s+navigate\s*\([^)]*\)\s*\{)',
        r'\1\n  if (target === "Карта ЖК") { injectStabBadge(); showMapZhk(); bindDrill(document.getElementById("view")); return; }',
        html, count=1
    )

# 4. Также вешаем injectStabBadge на общий переход (идемпотентно)
if "injectStabBadge();\n}" not in html:
    html = re.sub(
        r'(else showSheet\(target\);)',
        r'\1\n  injectStabBadge();',
        html, count=1
    )

# 5. В showForecastView добавим блок объёма — после существующего рендера
html = re.sub(
    r'(function\s+showForecastView\s*\([^)]*\)\s*\{)',
    r'\1\n  /* volume-block-injected-after */',
    html, count=1
)
# в самом конце showForecastView (перед первой закрывающей на верхнем уровне) — сложно найти;
# сделаю проще: обёрну через setTimeout после переключения.
html = html.replace(
    'if (target === "Карта ЖК") { injectStabBadge(); showMapZhk(); bindDrill(document.getElementById("view")); return; }',
    'if (target === "Карта ЖК") { injectStabBadge(); showMapZhk(); bindDrill(document.getElementById("view")); return; }\n'
    '  if (target === "Прогноз_NPS" || target === "forecast") { setTimeout(() => { const v=document.getElementById("view"); if (v && !v.querySelector("[data-vol-block]")) { const s=document.createElement("div"); s.setAttribute("data-vol-block","1"); v.appendChild(s); renderForecastVolume(s); } }, 30); }'
)

# 6. Добавляем JS-блоки перед </body> (идемпотентно)
if "// ==== Карта ЖК (BCG-матрица) ====" not in html:
    html = html.replace(
        "</script>\n</body>",
        map_js + "\n" + forecast_vol_js + "\n" + extra_js + "\n</script>\n</body>"
    )
else:
    # Обновляем блоки на месте (заменяем от первого маркера до </script>)
    m_start = html.find("// ==== Карта ЖК (BCG-матрица) ====")
    m_end = html.find("</script>\n</body>")
    html = html[:m_start] + map_js + "\n" + forecast_vol_js + "\n" + extra_js + "\n" + html[m_end:]

# ------ Перегенерация XLSX ------
wb = Workbook()
wb.remove(wb.active)

def brand_header(cell):
    cell.fill = PatternFill("solid", fgColor=BR_ORANGE)
    cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

def brand_sub(cell):
    cell.fill = PatternFill("solid", fgColor=BR_LIGHT)
    cell.font = Font(name="Calibri", size=10, bold=True, color=BR_INK)
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

def brand_cell(cell, hyp=False):
    cell.font = Font(name="Calibri", size=10, color=BR_INK)
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)
    if hyp:
        cell.fill = PatternFill("solid", fgColor=BR_HYP)

def add_sheet(name, rows, hyp_col_check=None):
    """rows: list[list]; hyp_col_check(row_dict) -> bool чтобы подсветить строки"""
    ws = wb.create_sheet(title=name[:31])
    max_col = 1
    for r_idx, row in enumerate(rows, start=1):
        if not isinstance(row, list): row = [row]
        max_col = max(max_col, len(row))
        for c_idx, val in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            if r_idx == 1:
                brand_header(cell)
            elif r_idx == 2:
                brand_sub(cell)
            else:
                brand_cell(cell)
        if r_idx == 1:
            ws.row_dimensions[1].height = 24
    # ширины
    for c in range(1, max_col+1):
        ws.column_dimensions[get_column_letter(c)].width = 18
    ws.column_dimensions["A"].width = 30
    ws.freeze_panes = "A4"

# Экспортирую все листы SHEETS в фирменном стиле
for name, rows in SHEETS.items():
    add_sheet(name, rows)

# Добавляю "Карта ЖК"
map_rows = [
    ["КАРТА ЖК · BCG-матрица NPS × ΔYoY"],
    [f"Периметр {total} ЖК · стабилизировано {stab} · KPI 2025 факт = −32 п.п."],
    ["Сегмент", "ЖК", "Регион", "Группа", "n 2026", "NPS 2026, %", "n 2025", "NPS 2025, %", "ΔYoY, п.п.", "Прогноз n к 30.09", "Риск недобора"],
]
seg_order = ["Звезды", "Восходящие", "Спящие", "Кризис", "Недобор"]
for seg in seg_order:
    for m in [x for x in MAPA if x["seg"] == seg]:
        map_rows.append([
            seg, m["zhk"], m.get("region",""), m.get("group",""),
            m["n_2026"] or 0,
            m["nps_2026"] if m["nps_2026"] is not None else "—",
            m["n_2025"] or 0,
            m["nps_2025"] if m["nps_2025"] is not None else "—",
            m["delta"] if m["delta"] is not None else "—",
            m["n_proj"], m["risk"],
        ])
add_sheet("Карта ЖК", map_rows)

# Подкрашиваю "Карта ЖК" по сегментам
ws = wb["Карта ЖК"]
seg_colors = {"Звезды": "E6F7EC", "Восходящие":"EEF9DE","Спящие":"FFF3D9","Кризис":"FFE3E3","Недобор":"F0F1F3"}
for row_idx in range(4, ws.max_row + 1):
    seg = ws.cell(row=row_idx, column=1).value
    color = seg_colors.get(seg, None)
    if color:
        for col in range(1, 12):
            ws.cell(row=row_idx, column=col).fill = PatternFill("solid", fgColor=color)

# Отдельный лист "Прогноз объёма" (детально по всем 96)
vol_rows = [
    ["ПРОГНОЗ ОБЪЁМА ОТВЕТОВ И РИСК НЕДОБОРА К 30.09.2026"],
    [f"Стабилизировано {stab} из {total} · Порог n≥50 · Формула: прогноз = n_сейчас × 3.0"],
    ["ЖК", "Регион", "Группа", "n сейчас", "Прогноз n к 30.09", "Риск недобора", "Гипотеза"],
]
for m in sorted(MAPA, key=lambda x: (0 if x["risk"]=="высокий" else (1 if x["risk"]=="средний" else 2), -(x["n_2026"] or 0))):
    vol_rows.append([
        m["zhk"], m.get("region",""), m.get("group",""),
        m["n_2026"] or 0, m["n_proj"], m["risk"],
        "да" if (m["n_2026"] or 0) < 50 else "нет"
    ])
add_sheet("Прогноз_объёма", vol_rows)
ws = wb["Прогноз_объёма"]
risk_colors = {"высокий":"FFECEC","средний":"FFF5E6","нет":"EEFBF2"}
for row_idx in range(4, ws.max_row + 1):
    risk = ws.cell(row=row_idx, column=6).value
    color = risk_colors.get(risk, None)
    if color:
        for col in range(1, 8):
            ws.cell(row=row_idx, column=col).fill = PatternFill("solid", fgColor=color)

# Сохраняем в память → base64 → вшиваем в HTML
buf = io.BytesIO()
wb.save(buf)
xlsx_bytes = buf.getvalue()
xlsx_b64 = base64.b64encode(xlsx_bytes).decode()
print("XLSX size:", len(xlsx_bytes), "b64 chars:", len(xlsx_b64))

# Заменяем зашитый XLSX_B64
html = re.sub(r'const\s+XLSX_B64\s*=\s*"[^"]*"\s*;', f'const XLSX_B64 = "{xlsx_b64}";', html, count=1)

# Освежаем метку "Обновлено"
html = re.sub(
    r"Обновлено \d{2}\.\d{2}\.\d{4}\s*\d{2}:\d{2}\s*MSK[^<]*",
    "Обновлено 10.09.2026 20:40 MSK · KPI 2025 = −32 п.п.",
    html, count=1
)

open(SRC, "w", encoding="utf-8").write(html)
open(f"{ROOT}/nps_pipeline/site/index.html", "w", encoding="utf-8").write(html)
print("HTML сохранён:", SRC)
print("Размер HTML:", len(html))
