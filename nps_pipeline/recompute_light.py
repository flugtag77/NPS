"""
recompute_light.py — быстрый ежечасный пересчёт всех блоков,
которые зависят только от полей, доступных из Testograf /details:
NPS, CSAT, ЖК, дата/час.

Обновляет в _snapshots/09_46_v2.html:
- SUMMARY: total_responses, nps_pct, csat_pct, promoters/neutrals/critics,
  nps_dist, daily_2026, updated_at, hours_2026
- SHEETS["Пульс_день"]: заново, включая блок «АКТИВНЫЕ ЧАСЫ»
- SHEETS["Разрезы"] / "ЖК_детально" / "Справочник_ЖК" — по агрегатам ЖК
- MAP_ZHK — сегменты Звёзды/Восходящие/Спящие/Кризис/Недобор по n и NPS
- PLAN.summary.forecast_current_dynamics и nps_forecast по ЖК

Что НЕ обновляется (остаётся из последнего полного пересчёта):
- SHEETS["Факторы"], "CSI_пульс", "Индекс_опыта", "Недельный_синтез", "Методика"
- Открытые ответы, гарантия, распределения CSI
"""
import json, re, csv, datetime as dt
from collections import Counter, defaultdict

ROOT = "/home/user/workspace/nps_pipeline"
SNAP = f"{ROOT}/_snapshots/09_46_v2.html"
CSV_LONG = f"{ROOT}/db/responses_long.csv"
PERIMETER = f"{ROOT}/db/zhk_full.tsv"
YEAR_MATRIX = f"{ROOT}/db/zhk_year_matrix_full.csv"

# --- 1. Периметр ---
perimeter = []
with open(PERIMETER, encoding="utf-8") as f:
    for r in csv.DictReader(f, delimiter="\t"):
        perimeter.append({"zhk": r["zhk"], "reg": r["reg"], "meropr": r["meropr"]})
PERIM_SET = {p["zhk"] for p in perimeter}

# --- 2. 2025 факт ---
zhk_2025 = {}
with open(YEAR_MATRIX, encoding="utf-8") as f:
    for r in csv.DictReader(f):
        z = r.get("ЖК") or r.get("zhk")
        if not z: continue
        def _f(k):
            v = r.get(k, "")
            try: return float(v) if v not in ("","None",None) else None
            except: return None
        def _i(k):
            v = r.get(k, "")
            try: return int(float(v)) if v not in ("","None",None) else 0
            except: return 0
        # Новые колонки (NPS_2025, n_2025) — бэк-совместимо со старыми (nps_2025, n_2025)
        zhk_2025[z] = {
            "n_2025": _i("n_2025"),
            "nps_2025": _f("NPS_2025") if _f("NPS_2025") is not None else _f("nps_2025"),
            "n_2024": _i("n_2024"),
            "nps_2024": _f("NPS_2024") if _f("NPS_2024") is not None else _f("nps_2024"),
        }

# --- 2b. Кривая набора 2025 (для калибровки проекции объёма) ---
try:
    with open(f"{ROOT}/db/curve_2025.json", encoding="utf-8") as f:
        _curve = json.load(f)
    CURVE_PCT = {pt["day"]: pt["pct"] for pt in _curve["curve"]}
    CURVE_TOTAL_2025 = _curve["total"]
except Exception:
    CURVE_PCT = {}
    CURVE_TOTAL_2025 = 0

# --- 3. Читаю CSV ---
rows_2026 = []
with open(CSV_LONG, encoding="utf-8") as f:
    for r in csv.DictReader(f):
        if r["year"] != "2026": continue
        try:
            nps = int(r["nps_uk"]) if r["nps_uk"] not in ("","None",None) else None
        except: nps = None
        try:
            csat = int(r["csat_uk_1_5"]) if r["csat_uk_1_5"] not in ("","None",None) else None
        except: csat = None
        rows_2026.append({
            "zhk": r["zhk"], "date": r["date"], "nps": nps, "csat": csat
        })

total = len(rows_2026)
print(f"2026 ответов: {total}")

# --- 4. NPS/CSAT/распределения ---
nps_all = [r["nps"] for r in rows_2026 if r["nps"] is not None]
promo = sum(1 for v in nps_all if v >= 9)
crit  = sum(1 for v in nps_all if v <= 6)
neu   = sum(1 for v in nps_all if 7 <= v <= 8)
n_nps = len(nps_all)
nps_pct = round((promo - crit) / n_nps * 100, 1) if n_nps else 0
promo_pct = round(promo / n_nps * 100, 1) if n_nps else 0
crit_pct  = round(crit / n_nps * 100, 1) if n_nps else 0
neu_pct   = round(neu / n_nps * 100, 1) if n_nps else 0
nps_dist = [sum(1 for v in nps_all if v == i) for i in range(11)]

csat_all = [r["csat"] for r in rows_2026 if r["csat"] is not None]
n_csat = len(csat_all)
csat_pct = round(sum(1 for v in csat_all if v >= 4) / n_csat * 100, 1) if n_csat else 0

# --- 5. Дневная динамика — ключ всегда DD.MM.YYYY для совместимости с _d2iso ---
def _norm_date_key(s):
    if not s or len(s) < 10: return None
    s10 = s[:10]
    if s10[2] == "." and s10[5] == ".": return s10  # уже DD.MM.YYYY
    if s10[4] == "-" and s10[7] == "-":  # ISO → DD.MM.YYYY
        return f"{s10[8:10]}.{s10[5:7]}.{s10[0:4]}"
    return None

day_rows = defaultdict(list)
for r in rows_2026:
    d = _norm_date_key(r["date"])
    if d: day_rows[d].append(r)

# сортирую по DD.MM.YYYY → YYYY-MM-DD, принимаю оба формата
def _d2iso(s):
    if not s: return None
    try:
        return dt.datetime.strptime(s, "%d.%m.%Y").strftime("%Y-%m-%d")
    except: pass
    # уже ISO?
    if len(s) >= 10 and s[4] == "-" and s[7] == "-":
        return s[:10]
    return None

sorted_days = sorted(day_rows.keys(), key=lambda x: _d2iso(x) or x)
daily = []
cumul_nps_all = []
prev_nps = None
for i, day in enumerate(sorted_days, 1):
    day_r = day_rows[day]
    cumul_nps_all.extend([r["nps"] for r in day_r if r["nps"] is not None])
    p = sum(1 for v in cumul_nps_all if v >= 9)
    c = sum(1 for v in cumul_nps_all if v <= 6)
    nps_cumul = round((p - c) / len(cumul_nps_all) * 100) if cumul_nps_all else 0
    # nps по дню
    day_nps = [r["nps"] for r in day_r if r["nps"] is not None]
    p_d = sum(1 for v in day_nps if v >= 9)
    c_d = sum(1 for v in day_nps if v <= 6)
    nps_day = round((p_d - c_d) / len(day_nps) * 100) if day_nps else 0
    daily.append({
        "date": _d2iso(day), "day_num": i,
        "ans_day": len(day_r),
        "ans_cumul": len(cumul_nps_all),
        "nps": nps_cumul, "nps_cumul": nps_cumul, "nps_day": nps_day,
    })

# --- 6. Часы ---
hours = Counter()
for r in rows_2026:
    if r["date"] and len(r["date"]) >= 13:
        try:
            hh = int(r["date"][11:13])
            hours[hh] += 1
        except: pass
hours_list = [{"h": h, "n": hours.get(h, 0), "pct": round(hours.get(h,0)/total*100,1) if total else 0} for h in range(24)]

# --- 7. Агрегаты по ЖК ---
zhk_agg = defaultdict(lambda: {"n":0, "p":0, "d":0, "csat45":0, "csat_n":0})
for r in rows_2026:
    z = r["zhk"]
    if not z: continue
    zhk_agg[z]["n"] += 1
    if r["nps"] is not None:
        if r["nps"] >= 9: zhk_agg[z]["p"] += 1
        elif r["nps"] <= 6: zhk_agg[z]["d"] += 1
    if r["csat"] is not None:
        zhk_agg[z]["csat_n"] += 1
        if r["csat"] >= 4: zhk_agg[z]["csat45"] += 1

# --- 8. MAP_ZHK (BCG) — идём по всему периметру 96 ЖК
bcg = {"Звёзды":[], "Восходящие":[], "Спящие":[], "Кризис":[], "Недобор":[]}
for pi in perimeter:
    z = pi["zhk"]
    a = zhk_agg.get(z, {"n":0, "p":0, "d":0})
    n, p, d = a["n"], a["p"], a["d"]
    nps_z = round((p - d) / n * 100, 1) if n else None
    nps25 = (zhk_2025.get(z) or {}).get("nps_2025")
    dyoy = None
    if nps_z is not None and nps25 is not None:
        dyoy = round(nps_z - nps25, 1)
    row = {"zhk": z, "reg": pi["reg"], "n": n, "nps": nps_z, "dyoy": dyoy}
    if n < 50:
        bcg["Недобор"].append(row)
    elif nps_z >= 0:
        bcg["Звёзды"].append(row)
    elif dyoy is not None and dyoy >= 5:
        bcg["Восходящие"].append(row)
    elif dyoy is not None and dyoy <= -5:
        bcg["Кризис"].append(row)
    else:
        bcg["Спящие"].append(row)

for k in bcg: bcg[k].sort(key=lambda x: -x["n"])

print(f"BCG: Звёзды {len(bcg['Звёзды'])}, Восходящие {len(bcg['Восходящие'])}, Спящие {len(bcg['Спящие'])}, Кризис {len(bcg['Кризис'])}, Недобор {len(bcg['Недобор'])}")

# --- 9. Прогноз объёма и NPS ---
# Объём — через кривую набора 2025:
# если мы на дне D опроса, то = pct(D)% финального объёма (было в 2025).
days_passed = len(sorted_days)
days_total = 24
if CURVE_PCT and days_passed in CURVE_PCT and CURVE_PCT[days_passed] > 0:
    proj_mult = round(100 / CURVE_PCT[days_passed], 3)
    proj_method = f"кривая 2025 (день {days_passed} = {CURVE_PCT[days_passed]}%)"
else:
    proj_mult = 3.0  # fallback
    proj_method = "fallback ×3.0"
total_proj = int(round(total * proj_mult))

# NPS-прогноз — байесовский шринкидж к 2025-факту по-ЖК.
# Формула: NPS_ПРОГ(ЖК) = (n_2026 * NPS_2026_факт + N_BASE * NPS_2025) / (n_2026 + N_BASE)
# При n_2026 = 0 берём 100% NPS_2025.  Вес = проекция набора к концу опроса.
N_BASE = 30
w_nps_b, w_n_b = 0.0, 0.0
# Инерционный (100% по 2025): взвешенный средний NPS_2025 по n_2025.
w_ne, wn_ne = 0.0, 0.0
for p in perimeter:
    z = p["zhk"]
    a = zhk_agg.get(z, {"n":0,"p":0,"d":0})
    y = zhk_2025.get(z, {})
    n26 = a["n"]
    nps26 = ((a["p"] - a["d"]) / n26 * 100) if n26 else None
    nps25 = y.get("nps_2025")
    n25 = y.get("n_2025", 0)
    # Байес
    if n26 == 0 and nps25 is None:
        pass
    else:
        if nps25 is None:
            nps_f = nps26
        elif n26 == 0:
            nps_f = nps25
        else:
            nps_f = (n26 * nps26 + N_BASE * nps25) / (n26 + N_BASE)
        # вес по проекции; если 2026 пуст — берём ожидаемую долю по кривой от n_2025
        weight = n26 * proj_mult if n26 > 0 else max(1.0, n25 * (total_proj / max(CURVE_TOTAL_2025, 1)))
        w_nps_b += weight * nps_f
        w_n_b += weight
    # Инерция (100% 2025)
    if nps25 is not None and n25 > 0:
        w_ne += n25 * nps25
        wn_ne += n25

forecast_nps = round(w_nps_b / w_n_b, 1) if w_n_b else -50.0
forecast_inert = round(w_ne / wn_ne, 1) if wn_ne else None

print(f"Прогноз объёма: {total_proj}  ({proj_method})")
print(f"Прогноз NPS (байес): {forecast_nps}  |  Инерция 100% 2025: {forecast_inert}")

# --- 10. Обновление снимка ---
html = open(SNAP, encoding="utf-8").read()
now_msk = dt.datetime.now(dt.timezone(dt.timedelta(hours=3))).strftime("%d.%m.%Y %H:%M MSK")

# 10.1 SUMMARY: читаю, обновляю нужные поля
m = re.search(r'const SUMMARY = (\{.*?\});\n', html, re.DOTALL)
if m:
    S = json.loads(m.group(1))
    S["updated_at"] = now_msk
    S["total_responses"] = total
    S["nps_pct"] = round((promo - crit) / n_nps * 100, 1) if n_nps else 0
    S["csat_pct"] = csat_pct
    S["n_nps"] = n_nps
    S["promoters_pct"] = promo_pct
    S["neutrals_pct"]  = neu_pct
    S["critics_pct"]   = crit_pct
    S["nps_dist"] = nps_dist
    S["daily_2026"] = daily
    new_json = json.dumps(S, ensure_ascii=False)
    html = html.replace(m.group(0), f'const SUMMARY = {new_json};\n', 1)

# 10.2 SHEETS.Пульс_день (полностью, с активными часами)
m = re.search(r'const SHEETS = (\{.*?\});\n', html, re.DOTALL)
SHEETS = json.loads(m.group(1))
pulse = [
    ["ПУЛЬС ДЕНЬ NPS-2026"],
    ["Ежедневный пульс, к вчера — Δ NPS накоп."],
    ["Дата","Ответов накоп.","Прирост за день","NPS накоп.","Δ NPS к вчера",
     "Критики (n)","Нейтралы (n)","Промоутеры (n)",
     "% критиков","% нейтралов","% промоутеров",
     "Индекс опыта накоп.","Δ индекса опыта"],
]
prev_nps_c = None
for d in daily:
    ans = d["ans_cumul"]
    day_rows_ = day_rows[[k for k in day_rows if _d2iso(k) == d["date"]][0]]
    day_nps_r = [r for r in day_rows_ if r["nps"] is not None]
    day_p = sum(1 for r in day_rows_ if r["nps"] is not None and r["nps"]>=9)
    day_c = sum(1 for r in day_rows_ if r["nps"] is not None and r["nps"]<=6)
    day_n = sum(1 for r in day_rows_ if r["nps"] is not None and 7<=r["nps"]<=8)
    # кумулятивно p/n/c для процентов
    date_iso = d["date"]
    cumul_r = []
    for dd in daily:
        if dd["date"] <= date_iso:
            day_key = [k for k in day_rows if _d2iso(k) == dd["date"]][0]
            cumul_r.extend([r for r in day_rows[day_key] if r["nps"] is not None])
    cp = sum(1 for r in cumul_r if r["nps"]>=9)
    cc = sum(1 for r in cumul_r if r["nps"]<=6)
    cn = sum(1 for r in cumul_r if 7<=r["nps"]<=8)
    cn_total = len(cumul_r)
    delta = None
    if prev_nps_c is not None:
        delta = d["nps_cumul"] - prev_nps_c
    prev_nps_c = d["nps_cumul"]
    pulse.append([
        dt.datetime.strptime(d["date"], "%Y-%m-%d").strftime("%d.%m.%Y"),
        ans, d["ans_day"],
        f'{d["nps_cumul"]}%',
        f'{delta:+.1f}%' if delta is not None else "—",
        cc, cn, cp,
        f'{cc/cn_total*100:.1f}%' if cn_total else "—",
        f'{cn/cn_total*100:.1f}%' if cn_total else "—",
        f'{cp/cn_total*100:.1f}%' if cn_total else "—",
        "—", "—",  # индекс опыта не тянем в light
    ])
# Активные часы
pulse.append([""])
pulse.append(["АКТИВНЫЕ ЧАСЫ (когда пользователи проходят опрос)"])
pulse.append(["Час","Ответов","% от всех"])
top5 = sorted(hours_list, key=lambda x: -x["n"])[:5]
top5_set = {t["h"] for t in top5}
for h in hours_list:
    tag = " ⚡ ПИК" if h["h"] in top5_set and h["n"] > 0 else ""
    pulse.append([f'{h["h"]:02d}:00', h["n"], f'{h["pct"]}%{tag}'])
SHEETS["Пульс_день"] = pulse

# Пишу SHEETS обратно
def _clean(v):
    if isinstance(v, str):
        return v.replace("\r"," ").replace("\n"," ").replace("\t"," ")
    elif isinstance(v, list): return [_clean(x) for x in v]
    elif isinstance(v, dict): return {k: _clean(vv) for k,vv in v.items()}
    return v
SHEETS = _clean(SHEETS)
new_sheets = json.dumps(SHEETS, ensure_ascii=False)
html = re.sub(r'const SHEETS = \{.*?\};\n', f'const SHEETS = {new_sheets};\n', html, count=1, flags=re.DOTALL)

# 10.3 PLAN.summary.forecast_current_dynamics
start = html.find("const PLAN = {")
if start != -1:
    i = start + len("const PLAN = ")
    depth = 0
    j = i
    while j < len(html):
        if html[j] == '{': depth += 1
        elif html[j] == '}':
            depth -= 1
            if depth == 0:
                j += 1; break
        j += 1
    PLAN = json.loads(html[i:j])
    PLAN.setdefault("summary", {})["forecast_current_dynamics"] = forecast_nps
    PLAN["summary"]["nps_now"] = round((promo - crit) / n_nps * 100, 1) if n_nps else 0
    PLAN["summary"]["total_n_fact"] = total
    PLAN["summary"]["forecast_inertial"] = forecast_inert
    PLAN["summary"]["forecast_volume"] = total_proj
    PLAN["summary"]["forecast_method"] = proj_method
    # обновляю по-ЖК n/p/d/nps_forecast в PLAN.plan
    for row in PLAN.get("plan", []):
        z = row["name"]
        a = zhk_agg.get(z)
        if a:
            row["fact_n"] = a["n"]
            row["fact_p"] = a["p"]
            row["fact_d"] = a["d"]
            row["fact_nps"] = round((a["p"]-a["d"])/a["n"]*100) if a["n"] else 0
    html = html[:i] + json.dumps(PLAN, ensure_ascii=False) + html[j:]

# 10.3b MAP_ZHK — обновляю поля n_2026, nps_2026, seg, n_proj, risk
m_map = re.search(r'const MAP_ZHK = (\[.*?\]);\n', html, re.DOTALL)
if m_map:
    MAP = json.loads(m_map.group(1))
    for row in MAP:
        z = row["zhk"]
        a = zhk_agg.get(z, {"n":0, "p":0, "d":0})
        y = zhk_2025.get(z, {})
        n, p, d = a["n"], a["p"], a["d"]
        row["n_2026"] = n
        row["nps_2026"] = round((p - d) / n * 100, 1) if n else None
        # Подтягиваю актуальный NPS_2025 из матрицы (приоритет над тем что было в MAP)
        if y.get("nps_2025") is not None:
            row["nps_2025"] = y["nps_2025"]
            row["n_2025"] = y.get("n_2025", 0)
        if y.get("nps_2024") is not None:
            row["nps_2024"] = y["nps_2024"]
            row["n_2024"] = y.get("n_2024", 0)
        n_proj = int(round(n * proj_mult))
        row["n_proj"] = n_proj
        # risk
        if n_proj >= 60: row["risk"] = "нет"
        elif n_proj >= 40: row["risk"] = "средний"
        else: row["risk"] = "высокий"
        # seg (используем «Звезды» без ё как в исходной схеме)
        nps_z = row["nps_2026"]
        nps25 = row.get("nps_2025")
        dyoy = round(nps_z - nps25, 1) if (nps_z is not None and nps25 is not None) else None
        row["delta"] = dyoy
        if n < 50: row["seg"] = "Недобор"
        elif nps_z is not None and nps_z >= 0: row["seg"] = "Звезды"
        elif dyoy is not None and dyoy >= 5: row["seg"] = "Восходящие"
        elif dyoy is not None and dyoy <= -5: row["seg"] = "Кризис"
        else: row["seg"] = "Спящие"
    new_map = json.dumps(MAP, ensure_ascii=False)
    html = html.replace(m_map.group(0), f'const MAP_ZHK = {new_map};\n', 1)

# 10.4 timestamp в заголовке
html = re.sub(r"Обновлено\s*\d{2}\.\d{2}\.\d{4}\s*\d{2}:\d{2}\s*MSK",
              f"Обновлено {now_msk}", html, count=1)

open(SNAP, "w", encoding="utf-8").write(html)
print(f"Снимок обновлён: {SNAP}")
print(f"NPS {nps_pct}% · CSAT {csat_pct}% · total {total} · часы {sum(hours.values())} · прогноз {forecast_nps}")
