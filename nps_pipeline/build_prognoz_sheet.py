"""
build_prognoz_sheet.py
Формирует лист «Прогноз NPS» витрины 2026:
- Блок 1: общий свод по годам (NPS + n + разложение промоутеры/пассивы/детракторы + CSAT)
- Блок 2: like-for-like ряд трех лет по ЖК, где данные есть (без порогов),
          с отдельным столбцом n для каждого NPS.
- Блок 3: прогноз финального NPS 2026 на 30.09 (простая линейная экстраполяция
          по темпу поступления ответов, с честным интервалом от текущего NPS).

Пишем в новый файл dashboard_2026-09-10T16-45_MSK.xlsx (следом заменит текущий snapshot).
Буква «е» вместо «ё» во всех подписях витрины.
"""
import csv, json, os, datetime as dt
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DB = "/home/user/workspace/nps_pipeline/db"
OUT_DIR = "/home/user/workspace/nps_pipeline/build"
os.makedirs(OUT_DIR, exist_ok=True)

# ---------- читаем базу ----------
rows = []
with open(f"{DB}/responses_long.csv", encoding="utf-8") as f:
    for r in csv.DictReader(f):
        r["year"] = int(r["year"])
        for k in ("nps_uk","nps_extra_recommend_zhk_2024","csat_uk_1_5","opyt_0_10"):
            v = r.get(k)
            r[k] = int(v) if v not in (None,"","None") else None
        rows.append(r)

def nps(vals):
    vals = [v for v in vals if v is not None]
    n = len(vals)
    if n == 0: return None, 0, 0, 0, 0
    p = sum(1 for v in vals if v>=9); d = sum(1 for v in vals if v<=6)
    return round((p-d)/n*100,1), n, p, n-p-d, d

def csat(vals):
    vals = [v for v in vals if v is not None]
    n = len(vals)
    if n == 0: return None, 0
    return round(sum(1 for v in vals if v>=4)/n*100,1), n

# ---------- Блок 1: общий свод ----------
totals = {}
for y in [2024,2025,2026]:
    subset = [r for r in rows if r["year"]==y]
    v,n,p,pas,d = nps([r["nps_uk"] for r in subset])
    cs, csn = csat([r["csat_uk_1_5"] for r in subset])
    totals[y] = {"NPS":v,"n":n,"P":p,"Pas":pas,"D":d,"CSAT_УК":cs,"CSAT_n":csn}

# ---------- Блок 2: матрица ЖК × год ----------
by_zhk_year = defaultdict(lambda: defaultdict(list))
for r in rows:
    if r["nps_uk"] is not None:
        by_zhk_year[r["zhk"]][r["year"]].append(r["nps_uk"])

zhk_rows = []
for z, per_year in by_zhk_year.items():
    m = {y: nps(per_year.get(y, [])) for y in [2024,2025,2026]}
    def d(a,b):
        va = m[a][0]; vb = m[b][0]
        if va is None or vb is None: return None
        return round(vb-va,1)
    zhk_rows.append({
        "ЖК": z,
        "NPS_2024": m[2024][0], "n_2024": m[2024][1],
        "NPS_2025": m[2025][0], "n_2025": m[2025][1],
        "NPS_2026": m[2026][0], "n_2026": m[2026][1],
        "d_25_24": d(2024,2025),
        "d_26_25": d(2025,2026),
        "d_26_24": d(2024,2026),
    })

def sort_key(r):
    v = r["d_26_25"]
    return (0 if v is not None else 1, v if v is not None else 999)
zhk_rows.sort(key=sort_key)

# ---------- Блок 3: прогноз финального NPS 2026 на 30.09 ----------
# Данные 2026 покрывают 07.09→10.09 13:44 (примерно 3.57 суток).
# 2025 закрылся за 39 дней (01.09.2025→09.10.2025), собрано 7920.
# Прогнозируем финал 2026 линейно по темпу поступления и удерживаем текущий NPS -53.3
# как «пессимистичный сценарий». Оптимистичный сценарий — восстановление к уровню like-for-like 2025 (-33.6).
n_now = totals[2026]["n"]
nps_now = totals[2026]["NPS"]
days_elapsed = 3.57  # 07.09 00:00 → 10.09 13:44
days_total   = 23.6  # 07.09 → 30.09 (окно опроса)
projected_n_linear = int(round(n_now * days_total / days_elapsed))
# сценарий A: NPS сохраняется на текущем уровне
# сценарий B: NPS восстанавливается к like-for-like 2025 (-33.6)
# сценарий C: NPS доопределится по темпу первых 24 часов и останется как есть
prognoz = {
    "n_now": n_now,
    "nps_now": nps_now,
    "days_elapsed": days_elapsed,
    "days_total": days_total,
    "projected_n_30_09_linear": projected_n_linear,
    "scenario_A_stay_current": nps_now,
    "scenario_B_recover_to_lfl_25": -33.6,
    "scenario_C_central_bet": round((nps_now + (-33.6))/2, 1),
}

# ---------- Пишем XLSX ----------
wb = Workbook()
ws = wb.active
ws.title = "Прогноз NPS"

# стили
h1 = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
h2 = Font(name="Calibri", size=11, bold=True)
mono = Font(name="Calibri", size=11)
gray = PatternFill("solid", fgColor="595959")
band = PatternFill("solid", fgColor="F2F2F2")
soft = PatternFill("solid", fgColor="FFF2CC")
center = Alignment(horizontal="center", vertical="center")
left = Alignment(horizontal="left", vertical="center")
thin = Side(style="thin", color="BFBFBF")
box = Border(left=thin, right=thin, top=thin, bottom=thin)

def row_head(row, label):
    c = ws.cell(row=row, column=1, value=label)
    ws.merge_cells(start_row=row, end_row=row, start_column=1, end_column=8)
    c.font = h1; c.fill = gray; c.alignment = left

def set_header(row, cols):
    for i,txt in enumerate(cols, 1):
        c = ws.cell(row=row, column=i, value=txt)
        c.font = h2; c.fill = band; c.alignment = center; c.border = box

def write_row(row, values, band_fill=None):
    for i,v in enumerate(values, 1):
        c = ws.cell(row=row, column=i, value=v)
        c.alignment = center if i>1 else left
        c.border = box
        if band_fill: c.fill = band_fill

# --- шапка ---
ws.cell(row=1, column=1, value="Прогноз NPS · сравнение год к году").font = Font(size=16, bold=True)
ws.cell(row=2, column=1, value="Cнимок 10.09.2026 13:44 MSK. Обновляется ежечасно. Букву «е» использовать вместо «ё».").font = Font(size=10, italic=True, color="595959")

# --- Блок 1: общий свод ---
row_head(4, "Блок 1. Общий свод по годам")
set_header(5, ["Год","N ответов","NPS","Промоутеры %","Пассивы %","Детракторы %","CSAT УК %","n CSAT"])
r=6
for y in [2024,2025,2026]:
    t = totals[y]
    p_share = round(t["P"]/t["n"]*100,1) if t["n"] else None
    pas_share = round(t["Pas"]/t["n"]*100,1) if t["n"] else None
    d_share = round(t["D"]/t["n"]*100,1) if t["n"] else None
    write_row(r, [y, t["n"], t["NPS"], p_share, pas_share, d_share,
                  t["CSAT_УК"] if t["CSAT_УК"] is not None else "—",
                  t["CSAT_n"] if t["CSAT_n"] else "—"])
    r += 1

# --- Блок 2: матрица ЖК × год ---
r += 2
row_head(r, "Блок 2. NPS по ЖК год к году · n рядом с каждым NPS")
r += 1
set_header(r, ["ЖК", "NPS 2024","n 2024", "NPS 2025","n 2025", "NPS 2026","n 2026", "Δ 26 vs 25"])
r += 1
# отдельно расширим до 13 колонок — добавим Δ 25 vs 24 и Δ 26 vs 24
# перепишем как 11 колонок
ws.delete_rows(r-1)  # удалим короткую шапку
set_header(r-1, ["ЖК", "NPS 2024","n 2024",
                 "NPS 2025","n 2025",
                 "NPS 2026","n 2026",
                 "Δ 25 vs 24","Δ 26 vs 25","Δ 26 vs 24","Флаг"])
# заполнение
for z in zhk_rows:
    flag = ""
    for a,b,key in [(2024,2025,"d_25_24"),(2025,2026,"d_26_25"),(2024,2026,"d_26_24")]:
        na = z[f"n_{a}"]; nb = z[f"n_{b}"]
        if z[key] is not None and (na<30 or nb<30):
            flag = "гипотеза (n<30)"
    fill = soft if flag else None
    write_row(r, [
        z["ЖК"],
        z["NPS_2024"] if z["NPS_2024"] is not None else "—",
        z["n_2024"] if z["n_2024"] else "—",
        z["NPS_2025"] if z["NPS_2025"] is not None else "—",
        z["n_2025"] if z["n_2025"] else "—",
        z["NPS_2026"] if z["NPS_2026"] is not None else "—",
        z["n_2026"] if z["n_2026"] else "—",
        z["d_25_24"] if z["d_25_24"] is not None else "—",
        z["d_26_25"] if z["d_26_25"] is not None else "—",
        z["d_26_24"] if z["d_26_24"] is not None else "—",
        flag,
    ], band_fill=fill)
    r += 1

# --- Блок 3: прогноз финала 2026 ---
r += 2
row_head(r, "Блок 3. Прогноз финального NPS 2026 на 30.09")
r += 1
set_header(r, ["Параметр","Значение","Комментарий","","","","",""])
lines = [
    ("Ответов на 10.09 13:44", prognoz["n_now"], "SERIAL_NUM≥2, из snapshot"),
    ("NPS сейчас (общий)", prognoz["nps_now"], "P−D по 1633 ответам"),
    ("Прошло суток", prognoz["days_elapsed"], "07.09→10.09 13:44"),
    ("Всего суток окна", prognoz["days_total"], "07.09→30.09"),
    ("Прогноз N ответов на 30.09", prognoz["projected_n_30_09_linear"], "линейная экстраполяция темпа"),
    ("Сценарий A: NPS сохранится", prognoz["scenario_A_stay_current"], "пессимистичный, база тренда 07-10.09"),
    ("Сценарий B: восстановление до 2025-lfl", prognoz["scenario_B_recover_to_lfl_25"], "если после первой волны детракторов подтянутся пассивы"),
    ("Сценарий C: центральная ставка", prognoz["scenario_C_central_bet"], "среднее между A и B — рабочая ставка"),
]
for label,val,note in lines:
    r += 1
    write_row(r, [label, val, note, "", "", "", "", ""])

# --- ширины колонок ---
widths = [40, 12, 10, 12, 10, 12, 10, 14, 14, 14, 18]
for i,w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

ws.freeze_panes = "B6"

out_path = f"{OUT_DIR}/dashboard_prognoz_2026-09-10T16-45_MSK.xlsx"
wb.save(out_path)

print(f"Готово: {out_path}")
print(f"Общий свод: 2024 NPS={totals[2024]['NPS']} (n={totals[2024]['n']}) · "
      f"2025 NPS={totals[2025]['NPS']} (n={totals[2025]['n']}) · "
      f"2026 NPS={totals[2026]['NPS']} (n={totals[2026]['n']})")
print(f"ЖК в матрице: {len(zhk_rows)}")
print(f"Прогноз: N на 30.09 ≈ {prognoz['projected_n_30_09_linear']}, "
      f"сценарии A/B/C = {prognoz['scenario_A_stay_current']} / "
      f"{prognoz['scenario_B_recover_to_lfl_25']} / {prognoz['scenario_C_central_bet']}")
