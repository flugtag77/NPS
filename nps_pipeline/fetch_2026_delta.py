"""
fetch_2026_delta.py
Забирает все актуальные ответы 2026 из Testograf API и мержит с базой.
Ловушки:
 - Список /responses возвращает только конверты (без ANSWERS).
 - Для каждого нового ответа нужен /responses/{id}/details.
 - Лимит 500 req/day и 120 req/min — уважаем.

Логика:
 1. Читаем текущий responses_long.csv (2024+2025+2026 snapshot).
 2. Список 2026 ответов из snapshot XLSX = базовая когорта resp_ids.
 3. Тянем /responses с 07.09.2026 (SERIAL_NUM≥2, NOT_COMPLETED=0).
 4. Дельта = ID из API, которых нет в базе.
 5. Для дельты — /responses/{id}/details по одному, с throttle.
 6. Считаем NPS Q9 (QUESTION_ID для «останетесь клиентом»), CSAT Q7,
    ЖК Q1 — по mapping ЖК-названий (используем ALIAS_NAME из details).
 7. Мержим в responses_long.csv, перезаписываем.
"""
import json, csv, time, urllib.request, urllib.error
import os
from urllib.parse import urlencode

API = os.environ["TESTOGRAF_API_KEY"]
BASE = "https://core.testograf.ru/api/private/v1"
SURVEY = int(os.environ.get("TESTOGRAF_SURVEY_ID", "712624"))

DB = os.environ.get("NPS_DB", os.path.join(os.path.dirname(__file__), "db"))
LONG = f"{DB}/responses_long.csv"

def get(path, params=None):
    u = f"{BASE}{path}"
    if params: u += "?" + urlencode(params)
    req = urllib.request.Request(u, headers={"ApiKey": API})
    try:
        return json.loads(urllib.request.urlopen(req, timeout=30).read().decode())
    except urllib.error.HTTPError as e:
        return {"HTTP_ERROR": e.code, "body": e.read()[:300].decode('utf-8','ignore')}

# ---- 1) Читаем базу ----
existing = []
with open(LONG, encoding="utf-8") as f:
    for r in csv.DictReader(f):
        existing.append(r)
have_2026 = {r["resp_id"] for r in existing if r["year"]=="2026"}
print(f"База: {len(existing)} строк. 2026 уже есть: {len(have_2026)} ответов.")

# ---- 2) Полный список 2026 из API ----
all_resp = []
page = 1
while True:
    r = get(f"/surveys/{SURVEY}/responses",
            {"start":"07.09.2026","end":"30.09.2026","limit":500,"page":page})
    if "HTTP_ERROR" in r:
        print("ERR list:", r); break
    batch = r.get("responses", [])
    all_resp += batch
    if len(batch) < 500: break
    page += 1
    time.sleep(0.6)
print(f"API вернул: {len(all_resp)} ответов")

# Фильтр: SERIAL_NUM>=2 и статус completed
valid = [x for x in all_resp if x.get("SERIAL_NUM",0)>=2 and not x.get("NOT_COMPLETED",0)]
print(f"После фильтра SERIAL>=2 и completed: {len(valid)}")

# Дельта
new_ids = [x["ID"] for x in valid if str(x["ID"]) not in have_2026]
print(f"Дельта (новых после snapshot): {len(new_ids)}")

# ---- 3) Тянем details ----
# Найдем QUESTION_ID для сквозного NPS и CSAT УК из survey details (один запрос)
survey_meta = get(f"/surveys/{SURVEY}/details")
qs = survey_meta.get("questions", [])
print(f"Вопросов в опросе: {len(qs)}")
# По заголовкам: NPS = «останетесь клиентом ПИК-Комфорт», CSAT УК = «Оцените ... работу управляющей компании в целом», ЖК = «Укажите название вашего жилого квартала»
def find_q(subs):
    for q in qs:
        t = (q.get("QUESTION_TEXT") or "").lower()
        if all(s.lower() in t for s in subs):
            return q["ID"], q.get("QUESTION_TEXT")[:80]
    return None, None

Q_NPS, t1 = find_q(["останетесь клиентом", "пик-комфорт"])
Q_CSAT, t2 = find_q(["работу управляющей компании в целом"])
Q_ZHK, t3 = find_q(["название вашего жилого квартала"])
Q_OPYT, t4 = find_q(["оцените от 0 до 10"])
print(f"Q_NPS  ID={Q_NPS} · {t1}")
print(f"Q_CSAT ID={Q_CSAT} · {t2}")
print(f"Q_ZHK  ID={Q_ZHK} · {t3}")
print(f"Q_OPYT ID={Q_OPYT} · {t4}")

# Быстрый мап ANSWER_ID → название ЖК (для варианта Q_ZHK — выпадающий список)
zhk_map = {}
for q in qs:
    if q["ID"] == Q_ZHK:
        for a in q.get("answers", []):
            zhk_map[a["ID"]] = a.get("ANSWER_STR") or a.get("ANSWER_TEXT")
print(f"ЖК в справочнике вопроса: {len(zhk_map)}")

# ---- 4) Забираем details ----
new_rows = []
errors = 0
for i, rid in enumerate(new_ids):
    if i % 20 == 0:
        print(f"  {i}/{len(new_ids)}...")
    r = get(f"/surveys/{SURVEY}/responses/{rid}/details")
    if "HTTP_ERROR" in r:
        errors += 1
        if errors > 5:
            print("Много ошибок, стоп."); break
        continue
    details = r.get("details", [])
    # найдем нужные значения
    nps_val = csat_val = opyt_val = zhk_name = None
    for d in details:
        qid = d.get("id")
        for a in d.get("responses", []):
            if qid == Q_NPS and a.get("ANSWER_INT") is not None:
                nps_val = a["ANSWER_INT"]
            elif qid == Q_CSAT and a.get("ANSWER_INT") is not None:
                csat_val = a["ANSWER_INT"]
            elif qid == Q_OPYT and a.get("ANSWER_INT") is not None:
                opyt_val = a["ANSWER_INT"]
            elif qid == Q_ZHK and a.get("ANSWER_BOOL") == 1:
                aid = a.get("ANSWER_ID")
                if aid in zhk_map: zhk_name = zhk_map[aid]
    if zhk_name is None:
        # попробуем из ALIAS_NAME конверта
        for x in valid:
            if x["ID"] == rid:
                zhk_name = x.get("ALIAS_NAME")
                break
    new_rows.append({
        "year": "2026",
        "resp_id": rid,
        "date": next((x.get("START_DATE","") for x in valid if x["ID"]==rid),""),
        "zhk": zhk_name or "",
        "nps_uk": nps_val,
        "nps_extra_recommend_zhk_2024": None,
        "csat_uk_1_5": csat_val,
        "opyt_0_10": opyt_val,
    })
    time.sleep(0.55)  # 120 req/min safety

print(f"Собрано новых: {len(new_rows)}, ошибок: {errors}")

# ---- 5) Мержим и перезаписываем ----
if new_rows:
    all_rows = existing + new_rows
    fields = ["year","resp_id","date","zhk","nps_uk","nps_extra_recommend_zhk_2024","csat_uk_1_5","opyt_0_10"]
    with open(LONG, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in all_rows: w.writerow(r)
    print(f"База обновлена: {len(all_rows)} строк.")
else:
    print("Дельта пустая или добор не удался.")
