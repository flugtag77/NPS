"""
build_zhk_matrix.py
Матрица ЖК × год: NPS + n на каждый год + дельты.
Никаких порогов отсечения — все ЖК попадают в сравнение как есть,
n показывается рядом с каждым NPS.
Так же считаем CSAT (доля 4-5 из 5) по 2025 и 2026 — «работа УК в целом».
"""
import csv, json, os
from collections import defaultdict

DB = "/home/user/workspace/nps_pipeline/db"

# --- читаем long ---
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
    p = sum(1 for v in vals if v>=9)
    d = sum(1 for v in vals if v<=6)
    pas = n - p - d
    return round((p-d)/n*100,1), n, p, pas, d

def csat(vals):
    vals = [v for v in vals if v is not None]
    n = len(vals)
    if n == 0: return None, 0
    top = sum(1 for v in vals if v>=4)
    return round(top/n*100,1), n

# --- Общий свод по годам ---
years = {}
for y in [2024,2025,2026]:
    subset = [r for r in rows if r["year"]==y]
    val,n,p,pas,d = nps([r["nps_uk"] for r in subset])
    csat_val, csat_n = csat([r["csat_uk_1_5"] for r in subset])
    zhk_count = len({r["zhk"] for r in subset})
    years[y] = {
        "N": n, "NPS": val,
        "promoters": p, "passives": pas, "detractors": d,
        "p_share": round(p/n*100,1) if n else None,
        "pas_share": round(pas/n*100,1) if n else None,
        "d_share": round(d/n*100,1) if n else None,
        "CSAT_uk_pct": csat_val, "CSAT_n": csat_n,
        "zhk_count": zhk_count,
    }

# --- Матрица ЖК × год ---
by_zhk_year = defaultdict(lambda: defaultdict(list))
for r in rows:
    if r["nps_uk"] is not None:
        by_zhk_year[r["zhk"]][r["year"]].append(r["nps_uk"])

zhk_rows = []
for z, per_year in by_zhk_year.items():
    row = {"ЖК": z}
    metrics_by_year = {}
    for y in [2024,2025,2026]:
        v,n,p,pas,d = nps(per_year.get(y, []))
        row[f"NPS_{y}"]  = v
        row[f"n_{y}"]    = n
        metrics_by_year[y] = (v, n)
    # дельты — считаем только там, где оба года есть (n>0 и NPS!=None)
    def delta(a,b):
        va, na = metrics_by_year[a]
        vb, nb = metrics_by_year[b]
        if va is None or vb is None: return None
        return round(vb - va, 1)
    row["Δ 2025 vs 2024"] = delta(2024, 2025)
    row["Δ 2026 vs 2025"] = delta(2025, 2026)
    row["Δ 2026 vs 2024"] = delta(2024, 2026)
    # флаг «гипотеза» — если в любом из годов сравнения n<30
    row["hypothesis_2025_2024"] = "гипотеза" if (metrics_by_year[2024][1] < 30 or metrics_by_year[2025][1] < 30) else ""
    row["hypothesis_2026_2025"] = "гипотеза" if (metrics_by_year[2025][1] < 30 or metrics_by_year[2026][1] < 30) else ""
    row["hypothesis_2026_2024"] = "гипотеза" if (metrics_by_year[2024][1] < 30 or metrics_by_year[2026][1] < 30) else ""
    zhk_rows.append(row)

# сортировка: сначала худшая дельта 2026 vs 2025 (у кого сильнее упал — сверху)
def sort_key(r):
    d = r["Δ 2026 vs 2025"]
    return (0 if d is not None else 1, d if d is not None else 999)
zhk_rows.sort(key=sort_key)

# --- Сохраняем ---
cols = ["ЖК",
        "NPS_2024","n_2024",
        "NPS_2025","n_2025",
        "NPS_2026","n_2026",
        "Δ 2025 vs 2024","hypothesis_2025_2024",
        "Δ 2026 vs 2025","hypothesis_2026_2025",
        "Δ 2026 vs 2024","hypothesis_2026_2024"]
with open(f"{DB}/zhk_year_matrix_full.csv","w",newline="",encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=cols)
    w.writeheader()
    for r in zhk_rows:
        w.writerow({k: (r.get(k) if r.get(k) is not None else "") for k in cols})

# --- JSON для витрины ---
out = {
    "generated_at": "2026-09-10T16:42+03:00",
    "totals_by_year": years,
    "zhk_year_matrix": zhk_rows,
    "notes": [
        "Букву «е» использовать вместо «ё» во всех подписях витрины.",
        "Порог отсечения по n не применяется. n показывается рядом с каждым NPS.",
        "Флаг «гипотеза» проставлен там, где хотя бы в одном из сравниваемых годов n<30.",
        "Источники: 2024=NPS_2024_full.xlsx (18 778), 2025=NPS_2025_full.xlsx (7 920), 2026=snapshot 13:44 MSK 10.09.2026 (1 633).",
    ],
}
with open(f"{DB}/zhk_matrix.json","w",encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=2)

# --- Быстрая диагностика в stdout ---
print("=== Общий свод ===")
for y in [2024,2025,2026]:
    m = years[y]
    print(f"{y}: NPS={m['NPS']}  N={m['N']}  ЖК={m['zhk_count']}  "
          f"CSAT_УК={m['CSAT_uk_pct']}% (n={m['CSAT_n']})")

# --- Топ-10 роста и топ-10 падения ЖК 2026 vs 2025 ---
with_delta_26_25 = [r for r in zhk_rows if r["Δ 2026 vs 2025"] is not None]
print(f"\nЖК со сравнением 2025→2026: {len(with_delta_26_25)}")
print("\n=== Топ-10 падения NPS 2026 vs 2025 ===")
print(f"{'ЖК':<32} {'NPS25':>7} {'n25':>5} {'NPS26':>7} {'n26':>5} {'Δ':>7}  {'flag':<10}")
for r in with_delta_26_25[:10]:
    print(f"{r['ЖК'][:32]:<32} {r['NPS_2025']:>7} {r['n_2025']:>5} "
          f"{r['NPS_2026']:>7} {r['n_2026']:>5} {r['Δ 2026 vs 2025']:>7}  {r['hypothesis_2026_2025']:<10}")

print("\n=== Топ-10 роста NPS 2026 vs 2025 ===")
for r in with_delta_26_25[-10:][::-1]:
    print(f"{r['ЖК'][:32]:<32} {r['NPS_2025']:>7} {r['n_2025']:>5} "
          f"{r['NPS_2026']:>7} {r['n_2026']:>5} {r['Δ 2026 vs 2025']:>7}  {r['hypothesis_2026_2025']:<10}")

# --- Сколько ЖК из 96 периметра есть в 2026 ---
per = set()
with open("/home/user/workspace/projects/nps-2026-EkRj3lrbTCeDbXz0D5Va8g/files/reference/approved_zhk_2026_perimeter96.csv", encoding="utf-8") as f:
    for r in csv.DictReader(f):
        name = r.get("ЖК") or r.get("Название ЖК") or r.get("name") or list(r.values())[0]
        if name: per.add(name.strip())
zhk_2026 = {r["ЖК"] for r in zhk_rows if r["n_2026"] > 0}
matched = per & zhk_2026
in_data_not_in_per = zhk_2026 - per
in_per_not_in_data = per - zhk_2026
print(f"\nПериметр 96, найдено в данных 2026: {len(matched)} из {len(per)}")
print(f"В данных, но не в периметре: {len(in_data_not_in_per)} — {sorted(in_data_not_in_per)[:8]}")
print(f"В периметре, но нет данных: {len(in_per_not_in_data)} — {sorted(in_per_not_in_data)[:8]}")

print("\nСохранено:")
print(f"  {DB}/zhk_year_matrix_full.csv  ({len(zhk_rows)} ЖК, без отсечения)")
print(f"  {DB}/zhk_matrix.json")
