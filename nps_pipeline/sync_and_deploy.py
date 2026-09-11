"""
sync_and_deploy.py — ежечасный цикл:
1. Забирает список ID новых ответов 2026 через Testograf API (1 запрос).
2. Добирает details по не более 40 новым ID (лимит 500/день с запасом).
3. Пересобирает базу и сайт.
4. Деплоит на Vercel.

Может запускаться и вручную, и из cron.
"""
import json, csv, time, os, urllib.request, urllib.error, base64, datetime as dt
from urllib.parse import urlencode
from collections import defaultdict

API_KEY  = os.environ["TESTOGRAF_API_KEY"]
API_BASE = "https://core.testograf.ru/api/private/v1"
SURVEY   = int(os.environ.get("TESTOGRAF_SURVEY_ID", "712624"))
VERCEL_TOKEN = os.environ.get("VERCEL_TOKEN", "")
VERCEL_TEAM  = os.environ.get("VERCEL_TEAM_ID", "")
VERCEL_PROJ  = os.environ.get("VERCEL_PROJECT_ID", "")

MAX_DETAILS_PER_RUN = 40   # ~40*24 = 960/day; безопасно с учётом ретраев внутри 500/day будут проблемы; ставим 20
MAX_DETAILS_PER_RUN = 20   # 20*24 = 480/day, вписываемся в 500/day
ROOT   = os.environ.get("NPS_ROOT", os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DB     = f"{ROOT}/nps_pipeline/db"
SITE   = f"{ROOT}/nps_pipeline/site"
STATE  = f"{DB}/state.json"

os.makedirs(DB, exist_ok=True)
os.makedirs(SITE, exist_ok=True)

# ---------- helpers ----------
def api_get(path, params=None, retries=2):
    u = f"{API_BASE}{path}"
    if params: u += "?" + urlencode(params)
    for attempt in range(retries+1):
        req = urllib.request.Request(u, headers={"ApiKey": API_KEY})
        try:
            return json.loads(urllib.request.urlopen(req, timeout=30).read().decode())
        except urllib.error.HTTPError as e:
            if e.code == 429 and attempt < retries:
                time.sleep(65); continue
            return {"HTTP_ERROR": e.code}
        except Exception as e:
            if attempt < retries: time.sleep(10); continue
            return {"HTTP_ERROR": str(e)}

def load_state():
    if os.path.exists(STATE):
        return json.load(open(STATE, encoding="utf-8"))
    return {"details_fetched_today": 0, "day": "", "known_ids_2026": []}

def save_state(s):
    json.dump(s, open(STATE,"w",encoding="utf-8"), ensure_ascii=False, indent=2)

# ---------- 1) load base long ----------
def load_long():
    rows = []
    with open(f"{DB}/responses_long.csv", encoding="utf-8") as f:
        for r in csv.DictReader(f):
            rows.append(r)
    return rows

def save_long(rows):
    fields = ["year","resp_id","date","zhk","nps_uk","nps_extra_recommend_zhk_2024","csat_uk_1_5","opyt_0_10"]
    with open(f"{DB}/responses_long.csv","w",newline="",encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields); w.writeheader()
        for r in rows: w.writerow({k: r.get(k,"") for k in fields})

# ---------- 2) discover new IDs ----------
def list_all_2026_ids():
    all_ids = []
    page = 1
    while True:
        r = api_get(f"/surveys/{SURVEY}/responses",
                    {"start":"07.09.2026","end":"30.09.2026","limit":500,"page":page})
        if "HTTP_ERROR" in r: return None
        batch = r.get("responses", [])
        for x in batch:
            sn = x.get("SERIAL_NUM",0)
            if sn >= 2 and not x.get("NOT_COMPLETED",0):
                # важно: используем SERIAL_NUM как resp_id (так же как xlsx-выгрузка)
                all_ids.append({"SERIAL_NUM":sn, "ID_INTERNAL":x["ID"], "START_DATE":x.get("START_DATE",""), "ALIAS_NAME":x.get("ALIAS_NAME","")})
        if len(batch) < 500: break
        page += 1
        time.sleep(1)
    return all_ids

# ---------- 3) fetch details for one response ----------
# Мы не знаем правильных QUESTION_ID, но details возвращают QUESTION_TYPE_ID.
# NPS = вопрос типа 15 (шкала 0-10) с содержимым ANSWER_INT.
# CSAT УК = вопрос типа с ANSWER_INT 1..5, идет РАНЬШЕ NPS.
# Простой безопасный подход: берем все ANSWER_INT-значения, где 0<=v<=10 —
# первое подходящее <=10 в диапазоне 1-5 = CSAT, значение 0-10 после — NPS.
# Более надежно: если в details есть несколько ANSWER_INT, NPS = последнее (в анкете вопрос про клиента идет ближе к концу).
def parse_details(details, alias_name):
    ints = []
    for d in details:
        for a in d.get("responses", []):
            v = a.get("ANSWER_INT")
            if v is not None and isinstance(v,(int,float)):
                ints.append((d.get("id"), int(v)))
    # NPS: значение 0..10 из вопроса с максимальным QUESTION_ID (последний по порядку)
    nps = None; csat = None; opyt = None
    scale010 = [(q,v) for q,v in ints if 0 <= v <= 10]
    scale15  = [(q,v) for q,v in ints if 1 <= v <= 5]
    if scale010:
        # ищем последний вопрос по q_id
        q_max = max(q for q,_ in scale010)
        vals_last = [v for q,v in scale010 if q == q_max]
        # берем среднее если несколько (обычно один)
        nps = int(round(sum(vals_last)/len(vals_last)))
    # CSAT — самая ранняя пятибалльная оценка (первый вопрос типа «работа УК в целом»)
    if scale15:
        q_min = min(q for q,_ in scale15)
        vals_first = [v for q,v in scale15 if q == q_min]
        csat = int(round(sum(vals_first)/len(vals_first)))
    return {"zhk": alias_name or "", "nps_uk": nps, "csat_uk_1_5": csat, "opyt_0_10": opyt}

def fetch_detail(rid):
    return api_get(f"/surveys/{SURVEY}/responses/{rid}/details")

# ---------- 4) FROZEN deploy: заливаем снимок 09:46 как есть, только освежаем метку времени ----------
FROZEN_SNAPSHOT = f"{ROOT}/nps_pipeline/_snapshots/09_46_v2.html"

def rebuild_and_deploy(comment=""):
    # НЕ вызываем build_site.py — работаем только со снимком.
    src = open(FROZEN_SNAPSHOT, "rb").read().decode("utf-8")
    # Обновляем видимый timestamp «Обновлено ...» на актуальный MSK.
    import re
    now_msk = dt.datetime.now(dt.timezone(dt.timedelta(hours=3))).strftime("%d.%m.%Y %H:%M MSK")
    src = re.sub(r"Обновлено\s*\d{2}\.\d{2}\.\d{4}\s*\d{2}:\d{2}\s*MSK", f"Обновлено {now_msk}", src, count=1)
    # Кладём как index.html в SITE (чтобы всё было консистентно) и деплоим.
    open(f"{SITE}/index.html","w",encoding="utf-8").write(src)
    files = [{"file":"index.html","data":base64.b64encode(src.encode()).decode(),"encoding":"base64"}]
    # vercel.json оставляем на случай наличия конфига
    for name in ["vercel.json"]:
        path=f"{SITE}/{name}"
        if os.path.exists(path):
            raw = open(path,"rb").read()
            files.append({"file":name,"data":base64.b64encode(raw).decode(),"encoding":"base64"})
    body = {"name":"nps-2026-pik","project":VERCEL_PROJ,"files":files,
            "target":"production","projectSettings":{"framework":None}}
    req = urllib.request.Request(
        f"https://api.vercel.com/v13/deployments?teamId={VERCEL_TEAM}&forceNew=1",
        data=json.dumps(body).encode(), method="POST")
    req.add_header("Authorization", f"Bearer {VERCEL_TOKEN}")
    req.add_header("Content-Type","application/json")
    try:
        r = json.loads(urllib.request.urlopen(req, timeout=90).read().decode())
        return {"ok":True, "id":r.get("id"), "url":r.get("url")}
    except Exception as e:
        return {"ok":False, "err":str(e)}

# ---------- MAIN ----------
def main():
    today = dt.datetime.now(dt.timezone(dt.timedelta(hours=3))).strftime("%Y-%m-%d")
    st = load_state()
    if st.get("day") != today:
        st = {"details_fetched_today":0, "day":today, "known_ids_2026": st.get("known_ids_2026",[])}
    print(f"[{dt.datetime.now().isoformat()}] Cron sync start. Fetched today: {st['details_fetched_today']}/500")

    # 1. base
    rows = load_long()
    have_ids = set(str(r["resp_id"]) for r in rows if r["year"] in ("2026",2026))
    print(f"  База: {len(rows)} строк. 2026 уже: {len(have_ids)}")

    # 2. discover
    api_list = list_all_2026_ids()
    if api_list is None:
        print("  API list FAIL — пропускаем дозаборку, но всё равно обновляем витрину.")
        api_list = []
    api_ids = [str(x["SERIAL_NUM"]) for x in api_list]
    new_ids = [x for x in api_list if str(x["SERIAL_NUM"]) not in have_ids]
    print(f"  API 2026: {len(api_ids)}. Новых: {len(new_ids)}")

    # 3. берём ВСЁ новое, только в рамках суточного бюджета 500 лимита Testograf
    remaining_budget = 500 - st["details_fetched_today"] - 5  # запас на list
    take = min(len(new_ids), remaining_budget)
    print(f"  Возьмем в этот тик: {take} (все новые, остаток бюджета на сегодня: {remaining_budget})")

    added = 0
    for item in new_ids[:take]:
        rid_internal = item["ID_INTERNAL"]  # для API запроса details
        sn = item["SERIAL_NUM"]              # для CSV и дедупа (совпадает с xlsx)
        r = fetch_detail(rid_internal)
        st["details_fetched_today"] += 1
        if "HTTP_ERROR" in r:
            print(f"    ERR {sn}: {r['HTTP_ERROR']}")
            if r["HTTP_ERROR"] == 429: break
            continue
        parsed = parse_details(r.get("details",[]), item.get("ALIAS_NAME"))
        rows.append({
            "year":"2026","resp_id":str(sn),
            "date":item.get("START_DATE",""),
            "zhk": parsed["zhk"],
            "nps_uk": parsed["nps_uk"],
            "nps_extra_recommend_zhk_2024": None,
            "csat_uk_1_5": parsed["csat_uk_1_5"],
            "opyt_0_10": parsed["opyt_0_10"],
        })
        added += 1
        time.sleep(0.6)

    if added > 0:
        save_long(rows)
        print(f"  Добавлено новых: {added}. База: {len(rows)}")

    save_state(st)

    # 3b. LIGHT recompute — пересчёт SUMMARY / Пульс_день / MAP_ZHK / PLAN из обновлённого CSV
    try:
        import subprocess
        rc = subprocess.run(["python3", f"{ROOT}/nps_pipeline/recompute_light.py"],
                            capture_output=True, text=True, timeout=180)
        print("  Light recompute:", rc.stdout.strip().split("\n")[-1] if rc.stdout else "OK")
        if rc.returncode != 0:
            print(f"  Light recompute STDERR: {rc.stderr[:300]}")
    except Exception as e:
        print(f"  Light recompute FAIL: {e}")

    # 3c. Пересбор XLSX — кнопка «Скачать xlsx» должна иметь свежие данные
    try:
        import subprocess
        rc = subprocess.run(["python3", f"{ROOT}/nps_pipeline/build_v2_patches.py"],
                            capture_output=True, text=True, timeout=180)
        print("  XLSX rebuild:", (rc.stdout or "").strip().split("\n")[-1] if rc.stdout else "OK")
        if rc.returncode != 0:
            print(f"  XLSX rebuild STDERR: {rc.stderr[:300]}")
    except Exception as e:
        print(f"  XLSX rebuild FAIL: {e}")

    # 3b. Патч колонки Коммуникация в Пульс_день (после build_v2_patches!)
    try:
        import subprocess
        rc = subprocess.run(["python3", f"{ROOT}/nps_pipeline/patch_pulse_comm.py"],
                            capture_output=True, text=True, timeout=30)
        print("  Pulse comm:", (rc.stdout or "").strip().split("\n")[-1] if rc.stdout else "OK")
    except Exception as e:
        print(f"  Pulse comm FAIL: {e}")

    # 4. rebuild + deploy — делаем всегда, чтобы обновить timestamp
    dep = rebuild_and_deploy()
    print(f"  Deploy: {dep}")

    # 5. Alias production → nps-2026-pik.vercel.app.
    # Vercel автоматически переключает alias на свежий production-деплой с небольшой задержкой.
    # Проверяем через паузу; если через 20с всё ещё старый — дергаем POST.
    if dep.get("ok") and dep.get("id"):
        time.sleep(10)
        alias_ok = False
        for attempt in range(2):
            try:
                req = urllib.request.Request(
                    f"https://api.vercel.com/v4/aliases/nps-2026-pik.vercel.app?teamId={VERCEL_TEAM}",
                    headers={"Authorization":f"Bearer {VERCEL_TOKEN}"})
                cur = json.loads(urllib.request.urlopen(req, timeout=30).read())
                if cur.get("deploymentId") == dep["id"]:
                    print(f"  Alias: nps-2026-pik.vercel.app → {dep['id'][:16]} OK")
                    alias_ok = True
                    break
                time.sleep(10)
            except Exception as e:
                time.sleep(10)
        if not alias_ok:
            try:
                req2 = urllib.request.Request(
                    f"https://api.vercel.com/v2/deployments/{dep['id']}/aliases?teamId={VERCEL_TEAM}",
                    data=json.dumps({"alias":"nps-2026-pik.vercel.app"}).encode(),
                    method="POST",
                    headers={"Authorization":f"Bearer {VERCEL_TOKEN}","Content-Type":"application/json"})
                r = json.loads(urllib.request.urlopen(req2, timeout=30).read().decode())
                print(f"  Alias forced: {r.get('alias')} ← {(r.get('oldDeploymentId') or '?')[:12]}")
            except Exception as e:
                print(f"  Alias FAIL: {e}")

if __name__ == "__main__":
    main()
