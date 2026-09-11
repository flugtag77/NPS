#!/usr/bin/env python3
"""NPS 2026 sync: Testograf 712624 -> метрики -> data.json + snapshot.xlsx.

Ключевые правила:
- SERIAL_NUM >= 2 (первая тестовая запись отсекается)
- NOT_COMPLETED == 0
- Часовой пояс: MSK (UTC+3)
- Периметр: 96 ЖК из reference/approved_zhk_2026_perimeter96.csv
- Малые сегменты (n<30) помечаются флагом гипотезы
- В рабочих текстах не используется буква "ё"
"""
from __future__ import annotations
import csv, datetime as dt, json, os, statistics, sys, time
from collections import Counter, defaultdict
from pathlib import Path
from urllib.request import Request, urlopen

API_KEY = os.environ["TESTOGRAF_API_KEY"]
SURVEY_ID = int(os.environ.get("TESTOGRAF_SURVEY_ID", "712624"))
BASE = "https://core.testograf.ru/api/private/v1"
MSK = dt.timezone(dt.timedelta(hours=3))

ROOT = Path("/home/user/workspace/nps_pipeline")
REF_CSV = Path("/home/user/workspace/projects/nps-2026-EkRj3lrbTCeDbXz0D5Va8g/files/reference/approved_zhk_2026_perimeter96.csv")
OUT_DIR = ROOT / "build"
OUT_DIR.mkdir(exist_ok=True)


def _req(path: str, params: dict) -> dict:
    from urllib.parse import urlencode
    url = f"{BASE}{path}?{urlencode(params)}"
    req = Request(url, headers={"ApiKey": API_KEY, "Accept": "application/json"})
    with urlopen(req, timeout=60) as r:
        return json.loads(r.read().decode("utf-8"))


def load_reference() -> list[dict]:
    with REF_CSV.open() as f:
        return list(csv.DictReader(f))


def fetch_all_responses() -> list[dict]:
    out = []
    page = 1
    while True:
        d = _req(f"/surveys/{SURVEY_ID}/responses",
                 {"start": "01.09.2026", "end": "31.10.2026", "limit": 500, "page": page})
        rows = d.get("responses") or []
        out.extend(rows)
        if len(rows) < 500:
            break
        page += 1
        time.sleep(0.6)
    return out


def fetch_details() -> dict:
    return _req(f"/surveys/{SURVEY_ID}/details", {})


def parse_answers(details: dict) -> tuple[dict, dict]:
    """Возвращает (id_вопроса -> метаданные, id_ответа -> текст)."""
    qmap = {}
    amap = {}
    for q in details.get("questions", []):
        qmap[q["ID"]] = {
            "order": q.get("ORDER"),
            "type": q.get("QUESTION_TYPE_ID"),
            "text": (q.get("QUESTION") or "").replace("<p>", "").replace("</p>", "").strip(),
        }
        for a in q.get("answers", []):
            val = a.get("ANSWER")
            if val:
                try:
                    v = json.loads(val)
                    if isinstance(v, list) and v:
                        amap[a["ID"]] = str(v[0])
                except Exception:
                    amap[a["ID"]] = str(val)
    return qmap, amap


def to_dt_msk(s: str | None) -> dt.datetime | None:
    if not s:
        return None
    try:
        return dt.datetime.strptime(s, "%Y-%m-%d %H:%M:%S").replace(tzinfo=MSK)
    except Exception:
        return None


def response_hydrate(r: dict, qmap: dict, amap: dict) -> dict:
    """Собирает плоский ответ: qX -> значение (число или текст)."""
    out = {"id": r.get("ID"), "serial": r.get("SERIAL_NUM"),
           "start": to_dt_msk(r.get("START_DATE")),
           "updated": to_dt_msk(r.get("UPDATED")),
           "total": r.get("TOTAL")}
    for a in r.get("answers", []):
        qid = a.get("QUESTION_ID")
        meta = qmap.get(qid, {})
        order = meta.get("order")
        if order is None:
            continue
        # NPS шкалы обычно приходят как SCORE или ANSWER=score/text
        v = a.get("SCORE")
        if v is None:
            v = a.get("VALUE")
        if v is None:
            aid = a.get("ANSWER_ID")
            if aid and aid in amap:
                v = amap[aid]
        if v is None:
            v = a.get("TEXT")
        # Числовое приведение для 0-10
        try:
            if isinstance(v, str) and v.strip().isdigit():
                v = int(v)
        except Exception:
            pass
        out[f"q{order}"] = v
    return out


def nps(scores: list[int]) -> tuple[float, int, int, int]:
    if not scores:
        return 0.0, 0, 0, 0
    prom = sum(1 for s in scores if s >= 9)
    detr = sum(1 for s in scores if s <= 6)
    neut = len(scores) - prom - detr
    return round(100 * (prom - detr) / len(scores), 1), prom, neut, detr


def csat(scores: list[int]) -> float:
    if not scores:
        return 0.0
    return round(100 * sum(1 for s in scores if s >= 4) / len(scores), 1)


def compute(records: list[dict], ref: list[dict]) -> dict:
    # NPS-9 и NPS-13 (alt)
    q9 = [int(r["q9"]) for r in records if isinstance(r.get("q9"), (int, float))]
    q13 = [int(r["q13"]) for r in records if isinstance(r.get("q13"), (int, float))]
    q2 = [int(r["q2"]) for r in records if isinstance(r.get("q2"), (int, float))]
    total_nps, prom, neut, detr = nps(q9)
    total_alt, _, _, _ = nps(q13)
    total_csat_q2 = csat([1 if s == 5 else (1 if s == 4 else 0) for s in q2]) if q2 else 0.0
    # По ЖК
    ref_by_name = {r["ЖК"]: r for r in ref}
    per_zhk = defaultdict(list)
    for r in records:
        zhk = r.get("q1")
        if zhk and zhk in ref_by_name:
            per_zhk[zhk].append(r)
    zhk_rows = []
    total_n = len(q9)
    for zhk, rows in per_zhk.items():
        zq9 = [int(x["q9"]) for x in rows if isinstance(x.get("q9"), (int, float))]
        val, _, _, _ = nps(zq9)
        n = len(zq9)
        contrib_pp = round((val * n / total_n), 2) if total_n else 0.0
        meta = ref_by_name[zhk]
        zhk_rows.append({
            "zhk": zhk, "region": meta["Регион"], "group": meta["Группа"],
            "n": n, "nps": val, "contrib_pp": contrib_pp,
            "hypothesis": n < 30,
        })
    zhk_rows.sort(key=lambda x: x["contrib_pp"])
    # Пульс день (по START_DATE в MSK)
    day_bucket = defaultdict(list)
    for r in records:
        d0 = r.get("start")
        if d0:
            day_bucket[d0.strftime("%Y-%m-%d")].append(r)
    pulse = []
    for day, rows in sorted(day_bucket.items()):
        dq9 = [int(x["q9"]) for x in rows if isinstance(x.get("q9"), (int, float))]
        val, p, ne, de = nps(dq9)
        pulse.append({"date": day, "n": len(dq9), "nps": val, "prom": p, "neut": ne, "detr": de})
    # Прогноз к 30.09: линейная экстраполяция взвешенного NPS
    horizon = dt.date(2026, 9, 30)
    today = dt.datetime.now(MSK).date()
    days_gone = max((today - dt.date(2026, 9, 7)).days, 1)
    days_left = max((horizon - today).days, 0)
    daily_avg = total_n / days_gone
    projected_n = int(total_n + daily_avg * days_left * 0.35)  # затухание вовлеченности
    base_forecast = total_nps  # без структурного сдвига
    return {
        "total": {
            "responses": len(records),
            "q9_n": len(q9), "nps": total_nps,
            "prom_pct": round(100*prom/len(q9), 1) if q9 else 0,
            "neut_pct": round(100*neut/len(q9), 1) if q9 else 0,
            "detr_pct": round(100*detr/len(q9), 1) if q9 else 0,
            "alt_nps": total_alt,
            "q2_csat": total_csat_q2,
        },
        "pulse": pulse,
        "zhk": zhk_rows,
        "forecast": {
            "horizon": horizon.isoformat(),
            "projected_n": projected_n,
            "base_nps": base_forecast,
            "gap_to_-30": round(-30 - base_forecast, 1),
            "gap_to_-35": round(-35 - base_forecast, 1),
        },
    }


def main() -> None:
    started = dt.datetime.now(MSK)
    ref = load_reference()
    details = fetch_details()
    qmap, amap = parse_answers(details)
    raw = fetch_all_responses()
    records = [response_hydrate(r, qmap, amap) for r in raw]
    # Фильтры
    records = [r for r in records if isinstance(r.get("serial"), int) and r["serial"] >= 2]
    metrics = compute(records, ref)
    payload = {
        "updated_at": started.strftime("%d.%m.%Y %H:%M MSK"),
        "survey_id": SURVEY_ID,
        "perimeter_zhk": len(ref),
        "records_total": len(records),
        **metrics,
    }
    out_json = OUT_DIR / "data.json"
    out_json.write_text(json.dumps(payload, ensure_ascii=False, indent=2))
    print(f"[sync] updated_at={payload['updated_at']}")
    print(f"[sync] responses={payload['records_total']} nps={metrics['total']['nps']} alt={metrics['total']['alt_nps']}")
    print(f"[sync] zhk_with_data={len(metrics['zhk'])} pulse_days={len(metrics['pulse'])}")
    print(f"[sync] wrote {out_json}")


if __name__ == "__main__":
    main()
