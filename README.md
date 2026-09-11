# NPS-2026 · локальная копия проекта

## Структура

- `nps_pipeline/` — весь пайплайн витрины
  - `sync_and_deploy.py` — главный скрипт: пересчёт → пересборка → деплой Vercel
  - `patch_pulse_comm.py` — колонка «Коммуникация» на «Пульс_день»
  - `build_*.py` — сборщики отдельных листов
  - `deploy_vercel.py` — загрузка на Vercel (вызывается изнутри sync)
  - `db/` — данные: responses_long.csv (база ответов), zhk_meta.json (96 ЖК), factors.json, pulse.json, zhk_detail.json и др.
  - `_snapshots/09_46_v2.html` — мастер-HTML витрины
  - `site/` — то, что реально уходит на Vercel (index.html + vercel.json)
  - `build/` — XLSX-скачивание витрины (dashboard_prognoz_*.xlsx)
- `nps-2026-pipeline.zip` — скилл проекта для Perplexity (правила пайплайна)

## Как обновить витрину локально

```bash
cd nps_pipeline
python3 sync_and_deploy.py 2>&1 | tail -30
```

Перед запуском создайте `.env` по образцу `.env.example` и задайте
перевыпущенный ключ Testograf. Vercel-настройки требуются только для
устаревшего деплоя на Vercel.

## Конфиденциальные данные

- Не публикуйте `.env`, ключи API или учётные данные.
- Исходные ответы опроса и сгенерированные локальные файлы исключены из Git.
- После клонирования восстановите защищённые данные из отдельного хранилища.

## Витрина

- Локальная публикация строится из `nps_pipeline/site/`.
