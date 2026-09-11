"""
build_site.py — BI-витрина NPS-2026 ПИК-Комфорт.
Каждый расчёт пересчитывается на каждом тике (54 * * * *).
Сегменты ЖК определяются динамически по n>=50 и NPS/ΔYoY.
Якорь: NPS-2025 = -32 (нельзя упасть ниже).
"""
import csv, json, os, datetime as dt
from collections import defaultdict, Counter

DB   = "/home/user/workspace/nps_pipeline/db"
SITE = "/home/user/workspace/nps_pipeline/site"
os.makedirs(SITE, exist_ok=True)

# ============================================================
# КОНСТАНТЫ ПРОДУКТА
# ============================================================
N_MIN          = 50        # минимум респондентов для стабилизации ЖК
YEAR_KPI       = -32.0     # факт NPS-2025 = красная линия
YEAR_TARGET    = -20.0     # цель 2026 (лучше 2025)
SURVEY_START   = dt.date(2026, 9, 7)
SURVEY_END     = dt.date(2026, 9, 30)

# ============================================================
# БАЗА
# ============================================================
rows = []
with open(f"{DB}/responses_long.csv", encoding="utf-8") as f:
    for r in csv.DictReader(f):
        r["year"] = int(r["year"])
        for k in ("nps_uk", "csat_uk_1_5", "opyt_0_10"):
            v = r.get(k)
            r[k] = int(v) if v not in (None, "", "None") else None
        rows.append(r)

# Периметр 96 — берём из zhk_full.tsv (там ровно 96 согласованных ЖК)
perimeter = []
with open(f"{DB}/zhk_full.tsv", encoding="utf-8") as f:
    for r in csv.DictReader(f, delimiter="\t"):
        perimeter.append({"ЖК": r["zhk"], "Регион": r["reg"], "Мероприятия": r["meropr"]})
PERIMETER_ZHK = {p["ЖК"] for p in perimeter}

# ============================================================
# ХЕЛПЕРЫ
# ============================================================
def nps_stats(vals):
    vals = [v for v in vals if v is not None]
    n = len(vals)
    if n == 0:
        return None, 0, 0, 0, 0
    p = sum(1 for v in vals if v >= 9)
    d = sum(1 for v in vals if v <= 6)
    return round((p - d) / n * 100, 1), n, p, n - p - d, d

def csat_pct(vals):
    vals = [v for v in vals if v is not None]
    n = len(vals)
    if n == 0:
        return None, 0
    return round(sum(1 for v in vals if v >= 4) / n * 100, 1), n

def dist_0_10(vals):
    vals = [v for v in vals if v is not None]
    n = len(vals) or 1
    c = Counter(vals)
    out = []
    for s in range(11):
        cnt = c.get(s, 0)
        out.append({"score": s, "n": cnt, "pct": round(cnt / n * 100, 1)})
    return out, len(vals)

def parse_date(s):
    if not s: return None
    try:
        return dt.datetime.strptime(s.split(" ")[0], "%d.%m.%Y").date()
    except: return None

# ============================================================
# ОБЩИЙ СВОД ПО ГОДАМ
# ============================================================
totals = {}
for y in [2024, 2025, 2026]:
    subset = [r for r in rows if r["year"] == y]
    v, n, p, pas, d = nps_stats([r["nps_uk"] for r in subset])
    cs, csn = csat_pct([r["csat_uk_1_5"] for r in subset])
    totals[y] = {
        "n": n, "nps": v, "p": p, "pas": pas, "d": d,
        "p_pct":   round(p   / n * 100, 1) if n else None,
        "pas_pct": round(pas / n * 100, 1) if n else None,
        "d_pct":   round(d   / n * 100, 1) if n else None,
        "csat_pct": cs, "csat_n": csn,
        "zhk_count": len({r["zhk"] for r in subset}),
    }

# 2026 расширенно
sub26 = [r for r in rows if r["year"] == 2026]
dist_nps_2026, n_dist_nps   = dist_0_10([r["nps_uk"]     for r in sub26])
dist_opyt_2026, n_dist_opyt = dist_0_10([r["opyt_0_10"]  for r in sub26])
alt_v, alt_n, alt_p, alt_pas, alt_d = nps_stats([r["opyt_0_10"] for r in sub26])

# CSI (лежит рядом)
csi_data = {}
csi_path = f"{DB}/csi_breakdown.json"
if os.path.exists(csi_path):
    csi_data = json.load(open(csi_path, encoding="utf-8"))

totals[2026].update({
    "alt_nps": alt_v, "alt_n": alt_n,
    "csi_pct_2026": csi_data.get("csi_pct_2026"),
    "csi_avg_2026": csi_data.get("csi_avg_2026"),
    "csi_pct_2025": csi_data.get("csi_pct_2025"),
    "csi_pct_2024": csi_data.get("csi_pct_2024"),
    "csi_by_criterion": csi_data.get("csi_by_criterion", []),
})

# ============================================================
# ЖК × ГОД (с сегментацией)
# ============================================================
by_zhk = defaultdict(lambda: defaultdict(list))
for r in rows:
    if r["nps_uk"] is not None:
        by_zhk[r["zhk"]][r["year"]].append(r["nps_uk"])

# Справочник региона/мероприятия
peri_by_zhk = {p["ЖК"]: p for p in perimeter}

zhk_rows = []
n_total_26 = totals[2026]["n"] or 1

# Гарантируем строку на каждый ЖК периметра, даже если ответов 0
all_zhk_names = set(by_zhk.keys()) | PERIMETER_ZHK

for z in sorted(all_zhk_names):
    per_y = by_zhk.get(z, {})
    m = {y: nps_stats(per_y.get(y, [])) for y in [2024, 2025, 2026]}
    nps24, n24 = m[2024][0], m[2024][1]
    nps25, n25 = m[2025][0], m[2025][1]
    nps26, n26 = m[2026][0], m[2026][1]

    d_26_25 = None if (nps25 is None or nps26 is None) else round(nps26 - nps25, 1)
    d_25_24 = None if (nps24 is None or nps25 is None) else round(nps25 - nps24, 1)
    d_26_24 = None if (nps24 is None or nps26 is None) else round(nps26 - nps24, 1)

    # вклад в общий NPS 2026 (пп)
    vklad = round((n26 / n_total_26) * (nps26 or 0), 2) if n26 and nps26 is not None else 0.0

    # ДИНАМИЧЕСКАЯ СЕГМЕНТАЦИЯ
    # Недобор: n26<50 (порог метода)
    # Иначе:
    #   Звезды       — NPS26>=0
    #   Восходящие   — NPS26<0 и d_26_25>=+5
    #   Спящие       — NPS26<0 и -5<d_26_25<+5
    #   Кризис       — NPS26<0 и d_26_25<=-5
    if n26 < N_MIN:
        seg = "Недобор"
    elif nps26 is None:
        seg = "Недобор"
    elif nps26 >= 0:
        seg = "Звезды"
    else:
        dv = d_26_25 if d_26_25 is not None else 0
        if dv >= 5:
            seg = "Восходящие"
        elif dv <= -5:
            seg = "Кризис"
        else:
            seg = "Спящие"

    # % критиков
    vals26 = [v for v in per_y.get(2026, []) if v is not None]
    n_crit = sum(1 for v in vals26 if v <= 6)
    pct_crit = round(n_crit / len(vals26) * 100, 1) if vals26 else None
    n_prom = sum(1 for v in vals26 if v >= 9)
    pct_prom = round(n_prom / len(vals26) * 100, 1) if vals26 else None

    peri = peri_by_zhk.get(z, {})
    flag = "гипотеза (n<30)" if n26 < 30 and n26 > 0 else ""

    zhk_rows.append({
        "zhk": z,
        "region":  peri.get("Регион", ""),
        "meropr":  peri.get("Мероприятия", ""),
        "in_perimeter": z in PERIMETER_ZHK,
        "nps_2024": nps24, "n_2024": n24,
        "nps_2025": nps25, "n_2025": n25,
        "nps_2026": nps26, "n_2026": n26,
        "d_25_24": d_25_24, "d_26_25": d_26_25, "d_26_24": d_26_24,
        "vklad_2026": vklad,
        "pct_crit_2026": pct_crit,
        "n_crit_2026": n_crit if vals26 else 0,
        "pct_prom_2026": pct_prom,
        "seg": seg,
        "flag_26_25": flag,
    })

# сортировка: сначала по сегменту (Кризис, Спящие, Восходящие, Звезды, Недобор), внутри по вкладу
seg_order = {"Кризис": 0, "Спящие": 1, "Восходящие": 2, "Звезды": 3, "Недобор": 4}
zhk_rows.sort(key=lambda r: (seg_order.get(r["seg"], 9), r["vklad_2026"]))

# сколько ЖК в каждом сегменте
seg_counts = Counter(r["seg"] for r in zhk_rows if r["in_perimeter"])

# ============================================================
# LIKE-FOR-LIKE ЯДРО (n>=30 в 3 годах)
# ============================================================
set_3 = {r["zhk"] for r in zhk_rows if r["n_2024"] >= 30 and r["n_2025"] >= 30 and r["n_2026"] >= 30}
def subset_nps(zhk_set, year):
    vals = [r["nps_uk"] for r in rows if r["year"] == year and r["zhk"] in zhk_set and r["nps_uk"] is not None]
    return nps_stats(vals)
lfl_3 = {y: subset_nps(set_3, y) for y in [2024, 2025, 2026]}

# ============================================================
# ПУЛЬС ДНЯ — динамика NPS 2026 по датам
# ============================================================
by_date = defaultdict(list)
for r in sub26:
    d = parse_date(r.get("date"))
    if d and r["nps_uk"] is not None:
        by_date[d].append(r["nps_uk"])

pulse = []
cum_vals = []
for day in sorted(by_date.keys()):
    day_vals = by_date[day]
    cum_vals.extend(day_vals)
    day_nps, day_n, *_ = nps_stats(day_vals)
    cum_nps, cum_n, *_ = nps_stats(cum_vals)
    pulse.append({
        "date": day.strftime("%d.%m"),
        "date_iso": day.strftime("%Y-%m-%d"),
        "day_n": day_n, "day_nps": day_nps,
        "cum_n": cum_n, "cum_nps": cum_nps,
    })

# ============================================================
# ПРОГНОЗ ДО 30.09
# ============================================================
n_now = totals[2026]["n"]; nps_now = totals[2026]["nps"] or 0
# темп сбора по факту
if pulse:
    days_elapsed = max(1, (parse_date(sub26[-1].get("date")) - SURVEY_START).days + 1)
else:
    days_elapsed = max(1, (dt.date.today() - SURVEY_START).days + 1)
days_total = (SURVEY_END - SURVEY_START).days + 1
proj_n = int(round(n_now * days_total / days_elapsed)) if days_elapsed else n_now

# Сценарии прогноза NPS на 30.09
# A · динамика сохранится: nps_now
# B · возврат к 2025-lfl (если пассивы поднимутся)
# C · центральная = среднее A и KPI2025
# D · при 70% факт + 30% истории (историческая коррекция)
lfl25 = lfl_3[2025][0] if lfl_3.get(2025) and lfl_3[2025][0] is not None else -33.6
scen_A = nps_now
scen_B = lfl25
scen_C = round((nps_now + lfl25) / 2, 1)
scen_D = round(nps_now * 0.7 + YEAR_KPI * 0.3, 1)

prognoz = {
    "n_now": n_now, "nps_now": nps_now,
    "days_elapsed": days_elapsed, "days_total": days_total,
    "projected_n_30_09": proj_n,
    "scenario_A": scen_A, "scenario_B": scen_B,
    "scenario_C": scen_C, "scenario_D": scen_D,
    "kpi_2025": YEAR_KPI, "target_2026": YEAR_TARGET,
    "delta_to_kpi": round(scen_D - YEAR_KPI, 1),   # прогноз - KPI2025 (если >0 — держим планку)
    "delta_to_target": round(scen_D - YEAR_TARGET, 1),
    "risk_level": "green" if scen_D >= YEAR_KPI else ("yellow" if scen_D >= YEAR_KPI - 5 else "red"),
}

# ============================================================
# СПИСКИ ДЛЯ РАБОТЫ (динамически)
# ============================================================
# Просели: d_26_25 <= -5 (при n26 >= 10 — иначе шум)
proseli = [r for r in zhk_rows
           if r["d_26_25"] is not None and r["d_26_25"] <= -5 and r["n_2026"] >= 10]
proseli.sort(key=lambda r: r["d_26_25"])
# Недобор
nedobor = [r for r in zhk_rows if r["n_2026"] < N_MIN and r["in_perimeter"]]
nedobor.sort(key=lambda r: r["n_2026"])
# Лояльники: nps_2026 >= 0 ИЛИ d_26_25 >= +5
loyalniki = [r for r in zhk_rows
             if (r["nps_2026"] is not None and r["nps_2026"] >= 0) or
                (r["d_26_25"] is not None and r["d_26_25"] >= 5)]
loyalniki.sort(key=lambda r: -(r["nps_2026"] or -999))

# ============================================================
# ГОЛОС ЖИТЕЛЯ (кластеры)
# ============================================================
open_labeled = []
oap = f"{DB}/open_answers_labeled.jsonl"
if os.path.exists(oap):
    with open(oap, encoding="utf-8") as f:
        for line in f:
            try: open_labeled.append(json.loads(line))
            except: pass

cluster_counter = Counter()
for o in open_labeled:
    for c in o.get("clusters", []):
        cluster_counter[c] += 1
top_clusters = [{"cluster": c, "n": n} for c, n in cluster_counter.most_common(15)]

# по сегментам ответов
voice_by_seg = {"promo": [], "pass": [], "crit": []}
for o in open_labeled:
    nps = o.get("nps")
    if nps is None: continue
    if nps >= 9:   voice_by_seg["promo"].append(o)
    elif nps >= 7: voice_by_seg["pass"].append(o)
    else:          voice_by_seg["crit"].append(o)

voice = {
    "n_labeled": len(open_labeled),
    "clusters": top_clusters,
    "counts": {k: len(v) for k, v in voice_by_seg.items()},
    "examples": {
        k: [{"zhk": o["zhk"], "nps": o["nps"], "text": o["text"][:280]}
            for o in v[:6]]
        for k, v in voice_by_seg.items()
    },
}

# ============================================================
# ФИНАЛЬНЫЙ DATA.JSON
# ============================================================
data = {
    "updated_at_msk": dt.datetime.now(dt.timezone.utc).astimezone(dt.timezone(dt.timedelta(hours=3))).strftime("%d.%m.%Y %H:%M MSK"),
    "constants": {
        "N_MIN": N_MIN,
        "KPI_2025": YEAR_KPI,
        "TARGET_2026": YEAR_TARGET,
        "SURVEY_START": SURVEY_START.strftime("%d.%m.%Y"),
        "SURVEY_END":   SURVEY_END.strftime("%d.%m.%Y"),
        "perimeter_count": len(PERIMETER_ZHK),
    },
    "totals": totals,
    "zhk": zhk_rows,
    "seg_counts": dict(seg_counts),
    "lfl_3": {
        "count": len(set_3),
        "zhk_list": sorted(set_3),
        "series": {y: {"nps": lfl_3[y][0], "n": lfl_3[y][1]} for y in [2024, 2025, 2026]},
    },
    "pulse": pulse,
    "prognoz": prognoz,
    "lists": {
        "proseli":  [{"zhk": r["zhk"], "nps_2025": r["nps_2025"], "nps_2026": r["nps_2026"], "n_2026": r["n_2026"], "d_26_25": r["d_26_25"]} for r in proseli],
        "nedobor":  [{"zhk": r["zhk"], "n_2026": r["n_2026"], "nps_2026": r["nps_2026"], "seg": r["seg"]} for r in nedobor],
        "loyalniki":[{"zhk": r["zhk"], "nps_2025": r["nps_2025"], "nps_2026": r["nps_2026"], "n_2026": r["n_2026"], "d_26_25": r["d_26_25"]} for r in loyalniki],
    },
    "voice": voice,
    "dist_nps_2026": dist_nps_2026,
    "dist_opyt_2026": dist_opyt_2026,
    "source_note": "2024=NPS_2024_full.xlsx (18 778) · 2025=NPS_2025_full.xlsx (7 920) · 2026=Testograf API · обновление раз в час",
}
with open(f"{SITE}/data.json", "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

# ============================================================
# HTML — компактно, шаблон вынесен в site/index.html.tpl если хотим
# ============================================================
HTML_PATH = f"{SITE}/index.html"
HTML_TPL  = f"{os.path.dirname(__file__)}/site_template.html"

if os.path.exists(HTML_TPL):
    with open(HTML_TPL, encoding="utf-8") as f:
        html = f.read()
    with open(HTML_PATH, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"HTML взят из шаблона: {HTML_TPL}")
else:
    print(f"ВНИМАНИЕ: шаблон {HTML_TPL} не найден — index.html оставлен как есть")

with open(f"{SITE}/vercel.json", "w", encoding="utf-8") as f:
    json.dump({"cleanUrls": True}, f)

print(f"data.json пересобран. ЖК={len(zhk_rows)}, сегменты={dict(seg_counts)}")
print(f"Прогноз к 30.09: A={scen_A} B={scen_B} C={scen_C} D={scen_D} · KPI={YEAR_KPI} · Δ={prognoz['delta_to_kpi']} ({prognoz['risk_level']})")
