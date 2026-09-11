"""
build_base.py — единая база 2024/2025/2026 для витрины NPS 2026.
Правила фиксированы пользователем:
- NPS = «Останетесь клиентом ПИК-Комфорт» (сквозной вопрос трех лет)
- Сравнение like-for-like: год к году общее и каждый ЖК только сам с собой
- Малые сегменты n<30 — маркер «гипотеза»
- Буква «е» вместо «ё» в текстовых полях витрины (в исходных данных не трогаем)
"""
from openpyxl import load_workbook
import json, os, csv, re
from collections import defaultdict, Counter

ROOT = "/home/user/workspace"
PROJ = f"{ROOT}/projects/nps-2026-EkRj3lrbTCeDbXz0D5Va8g/files"
OUT  = f"{ROOT}/nps_pipeline/db"
os.makedirs(OUT, exist_ok=True)

def to_int(x, lo=0, hi=10):
    if x is None: return None
    s = str(x).strip()
    if not s: return None
    try:
        v = int(float(s.replace(",", ".")))
        if lo <= v <= hi: return v
    except: pass
    return None

def norm_zhk(name):
    if not name: return None
    s = str(name).strip()
    # снимаем множественные пробелы, кавычки-разного-типа
    s = re.sub(r"\s+", " ", s)
    s = s.replace("«","").replace("»","").replace("\"","").strip()
    return s

# ---------- 2024 ----------
rows = []
wb = load_workbook(f"{PROJ}/data_raw/NPS_2024_full.xlsx", data_only=True)
ws = wb["Sheet1"]
for r in range(2, ws.max_row+1):
    zhk = norm_zhk(ws.cell(r,9).value)
    if not zhk: continue
    rows.append({
        "year":2024, "resp_id":ws.cell(r,1).value, "date":str(ws.cell(r,2).value or ""),
        "zhk":zhk,
        "nps_uk": to_int(ws.cell(r,65).value),
        "nps_extra_recommend_zhk_2024": to_int(ws.cell(r,64).value),
        "csat_uk_1_5": None,  # в 2024 нет сводного вопроса «работа УК в целом 1-5»
        "opyt_0_10": None,    # нет вопроса «опыт жизни 0-10»
    })

# ---------- 2025 ----------
wb = load_workbook(f"{PROJ}/data_raw/NPS_2025_full.xlsx", data_only=True)
ws = wb["Sheet1"]
for r in range(2, ws.max_row+1):
    zhk = norm_zhk(ws.cell(r,9).value)
    if not zhk: continue
    rows.append({
        "year":2025, "resp_id":ws.cell(r,1).value, "date":str(ws.cell(r,2).value or ""),
        "zhk":zhk,
        "nps_uk": to_int(ws.cell(r,66).value),
        "nps_extra_recommend_zhk_2024": None,
        "csat_uk_1_5": to_int(ws.cell(r,37).value, 1, 5),
        "opyt_0_10": None,
    })

# ---------- 2026 (snapshot 1797) ----------
wb = load_workbook(f"{PROJ}/model/NPS-2026_snapshot_1797.xlsx", data_only=True)
ws = wb["Sheet1"]
for r in range(2, ws.max_row+1):
    zhk = norm_zhk(ws.cell(r,9).value)
    if not zhk: continue
    rows.append({
        "year":2026, "resp_id":ws.cell(r,1).value, "date":str(ws.cell(r,2).value or ""),
        "zhk":zhk,
        "nps_uk": to_int(ws.cell(r,42).value),
        "nps_extra_recommend_zhk_2024": None,
        "csat_uk_1_5": to_int(ws.cell(r,15).value, 1, 5),
        "opyt_0_10": to_int(ws.cell(r,46).value),
    })

print(f"Total rows: {len(rows)}")
print("Rows per year:", Counter(r["year"] for r in rows))

# ---------- Сохраняем long-таблицу CSV ----------
with open(f"{OUT}/responses_long.csv","w",newline="",encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
    w.writeheader()
    w.writerows(rows)

# ---------- Метрика NPS ----------
def nps_metrics(vals):
    vals = [v for v in vals if v is not None]
    n = len(vals)
    if n==0: return {"n":0,"promoters":0,"passives":0,"detractors":0,"nps":None}
    p = sum(1 for v in vals if v>=9)
    d = sum(1 for v in vals if v<=6)
    pas = n - p - d
    return {"n":n,"promoters":p,"passives":pas,"detractors":d,
            "nps":round((p-d)/n*100,2),
            "p_share":round(p/n*100,2),"d_share":round(d/n*100,2)}

# ---------- Общий свод по годам ----------
years = {}
for y in [2024,2025,2026]:
    subset = [r for r in rows if r["year"]==y]
    m = nps_metrics([r["nps_uk"] for r in subset])
    m["year"] = y
    m["zhk_count"] = len({r["zhk"] for r in subset})
    years[y] = m
    print(f"\n=== {y} ===")
    print(f"  N={m['n']}  NPS={m['nps']}  P={m['promoters']} ({m['p_share']}%)  "
          f"Pas={m['passives']}  D={m['detractors']} ({m['d_share']}%)  ЖК={m['zhk_count']}")

# ---------- Свод по ЖК × год ----------
by_zhk = defaultdict(lambda: defaultdict(list))
for r in rows:
    if r["nps_uk"] is not None:
        by_zhk[r["zhk"]][r["year"]].append(r["nps_uk"])

zhk_year = []
all_zhk = sorted(by_zhk.keys())
for z in all_zhk:
    row = {"zhk":z}
    for y in [2024,2025,2026]:
        m = nps_metrics(by_zhk[z][y])
        row[f"n_{y}"] = m["n"]
        row[f"nps_{y}"] = m["nps"]
    # like-for-like дельты, только если оба года есть и n>=30
    def d(a,b):
        va,vb = row.get(f"nps_{a}"), row.get(f"nps_{b}")
        na,nb = row.get(f"n_{a}"), row.get(f"n_{b}")
        if va is None or vb is None or na<30 or nb<30: return None
        return round(vb-va,1)
    row["delta_25_24"] = d(2024,2025)
    row["delta_26_25"] = d(2025,2026)
    row["delta_26_24"] = d(2024,2026)
    zhk_year.append(row)

with open(f"{OUT}/zhk_year_matrix.csv","w",newline="",encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=list(zhk_year[0].keys()))
    w.writeheader()
    w.writerows(zhk_year)

# ---------- Пересечение ЖК всех 3 лет (like-for-like ряд) ----------
common_3 = [r for r in zhk_year if r["n_2024"]>=30 and r["n_2025"]>=30 and r["n_2026"]>=30]
print(f"\nЖК присутствуют во всех 3 годах с n>=30: {len(common_3)}")
common_2 = [r for r in zhk_year if r["n_2025"]>=30 and r["n_2026"]>=30]
print(f"ЖК присутствуют в 2025 и 2026 с n>=30: {len(common_2)}")

# Свод по подмножеству ЖК, общих для 2024 и 2026 (n>=30 в каждом)
def subset_nps(zhk_set, year):
    vals = [r["nps_uk"] for r in rows if r["year"]==year and r["zhk"] in zhk_set and r["nps_uk"] is not None]
    return nps_metrics(vals)

set_3 = {r["zhk"] for r in common_3}
set_25_26 = {r["zhk"] for r in common_2}

lfl_3 = {y: subset_nps(set_3, y) for y in [2024,2025,2026]}
lfl_25_26 = {y: subset_nps(set_25_26, y) for y in [2025,2026]}

print("\n--- Like-for-like ряд 2024→2025→2026 (только ЖК, где n>=30 во всех 3 годах) ---")
for y in [2024,2025,2026]:
    m = lfl_3[y]
    print(f"  {y}: NPS={m['nps']}  N={m['n']}  ЖК={len(set_3)}")

print("\n--- Like-for-like ряд 2025→2026 (только ЖК, где n>=30 в обоих) ---")
for y in [2025,2026]:
    m = lfl_25_26[y]
    print(f"  {y}: NPS={m['nps']}  N={m['n']}  ЖК={len(set_25_26)}")

# ---------- JSON summary ----------
summary = {
    "generated_at": "2026-09-10T16:20+03:00",
    "source": {
        "2024": {"file":"data_raw/NPS_2024_full.xlsx","N":years[2024]["n"]},
        "2025": {"file":"data_raw/NPS_2025_full.xlsx","N":years[2025]["n"]},
        "2026": {"file":"model/NPS-2026_дашборд_2026-09-10T13-44_MSK.xlsx","N":years[2026]["n"],
                 "snapshot":"13:44 MSK 10.09.2026 — до подключения ежечасного sync"},
    },
    "totals_by_year": years,
    "like_for_like_3y": {
        "zhk_count": len(set_3),
        "zhk_list": sorted(set_3),
        "series": lfl_3,
    },
    "like_for_like_25_26": {
        "zhk_count": len(set_25_26),
        "zhk_list": sorted(set_25_26),
        "series": lfl_25_26,
    },
}
with open(f"{OUT}/summary.json","w",encoding="utf-8") as f:
    json.dump(summary, f, ensure_ascii=False, indent=2)

print("\nСохранено:")
print(f"  {OUT}/responses_long.csv  ({len(rows)} строк)")
print(f"  {OUT}/zhk_year_matrix.csv  ({len(zhk_year)} ЖК)")
print(f"  {OUT}/summary.json")
