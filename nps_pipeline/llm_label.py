"""LLM-разметка открытых ответов NPS-2026 по финальным 22+3 кластерам через pplx_sdk.llm.extract."""
import json, sys, os
import pplx_sdk

INPUT  = "/home/user/workspace/nps_pipeline/db/open_answers_2026.jsonl"
OUTPUT = "/home/user/workspace/nps_pipeline/db/open_answers_labeled.jsonl"

CLUSTERS = [
    "Чистота подъездов и МОП",
    "Чистота двора и мусор",
    "Работа лифтов",
    "Инженерные системы и коммунальные ресурсы",
    "Пожарная система и безопасность",
    "Домофония и контроль доступа",
    "Благоустройство территории",
    "Парковка",
    "Оперативность заявок",
    "Качество исполнения работ",
    "Информирование по заявкам и статусам",
    "Диспетчерская и колл-центр",
    "Личный кабинет и приложение",
    "Персонал: вежливость и компетентность",
    "Кадровые решения по управляющему ЖК",
    "Смена управляющей компании",
    "Собрания собственников и легитимность",
    "Финансы: тарифы, платежки, прозрачность расходов",
    "Соседи и коммьюнити",
    "Работа с нарушителями и вандализмом",
    "Гарантийные случаи и претензии к застройщику",
    "Внешние сервисы (интернет-провайдер, ПИК-Сервис, ПИК-Аренда)",
    "Положительная обратная связь",
    "Ничего не менять / всё устраивает",
    "Низкоинформативные",
]

INSTRUCTION = (
    "Ты аналитик клиентского опыта в управляющей компании жилой недвижимости ПИК-Комфорт. "
    "Разметь свободный ответ жителя (поле text) по кластерам болей и потребностей. "
    "Мультитэг: если ответ затрагивает несколько тем — верни все подходящие. "
    "Используй ТОЛЬКО значения из закрытого списка. Список кластеров:\n"
    + "\n".join(f"- {c}" for c in CLUSTERS)
    + "\n\nВерни массив выбранных кластеров в поле clusters."
)

SCHEMA = {
    "type": "object",
    "properties": {
        "clusters": {
            "type": "array",
            "items": {"type":"string", "enum": CLUSTERS},
            "minItems": 1
        }
    },
    "required": ["clusters"]
}

def main(limit=None):
    rows = [json.loads(l) for l in open(INPUT, encoding="utf-8")]
    if limit: rows = rows[:limit]
    print(f"Разметка {len(rows)} ответов через pplx_sdk.llm.extract")
    items = [{"text": r["text"][:1500]} for r in rows]
    results = pplx_sdk.llm.extract(items, INSTRUCTION, SCHEMA)
    print(f"Получено ответов: {len(results)}")
    with open(OUTPUT,"w",encoding="utf-8") as f:
        for src, res in zip(rows, results):
            clusters = []
            if hasattr(res,'result') and res.result:
                clusters = res.result.get("clusters", [])
            elif isinstance(res, dict):
                clusters = res.get("clusters", [])
            f.write(json.dumps({**src, "clusters": clusters or ["Низкоинформативные"]}, ensure_ascii=False)+"\n")
    print(f"Готово: {OUTPUT}")

if __name__ == "__main__":
    lim = int(sys.argv[1]) if len(sys.argv)>1 else None
    main(lim)
