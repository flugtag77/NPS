#!/usr/bin/env python3
"""Сборка PPTX-playbook по клиентским опросам Testograf.

Стиль: палитра из make_mockup_pdf.py (ПИК-Комфорт).
В рабочих текстах не используется буква "е" с точками (е вместо ё).
"""
from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

# --- Brand ---
ORANGE = RGBColor(0xFC, 0x4C, 0x02)
VIOLET = RGBColor(0x96, 0x78, 0xD2)
VIOLET_DARK = RGBColor(0x7C, 0x5E, 0xB8)
INK = RGBColor(0x1A, 0x1A, 0x1A)
GRAY = RGBColor(0x68, 0x68, 0x68)
GRAY_LIGHT = RGBColor(0xB2, 0xB2, 0xB2)
BG = RGBColor(0xF7, 0xF5, 0xFB)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREEN = RGBColor(0x2F, 0xA8, 0x4F)
RED = RGBColor(0xE2, 0x4C, 0x4B)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
MARGIN = Inches(0.45)

OUT = Path(__file__).resolve().parent / "build" / "Testograf_playbook_client_surveys.pptx"


def _set_run(run, *, size=14, bold=False, color=INK, font="Arial"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font
    # East Asian / complex script fallback
    rPr = run._r.get_or_add_rPr()
    for tag in ("latin", "ea", "cs"):
        el = rPr.find(qn(f"a:{tag}"))
        if el is None:
            el = rPr.makeelement(qn(f"a:{tag}"), {})
            rPr.append(el)
        el.set("typeface", font)


def _add_text(tf, text, *, size=14, bold=False, color=INK, align=PP_ALIGN.LEFT, space_after=6):
    p = tf.paragraphs[0] if not tf.paragraphs[0].text else tf.add_paragraph()
    if not tf.paragraphs[0].text and tf.paragraphs[0] is p:
        pass
    else:
        if tf.paragraphs[0].text == "" and len(tf.paragraphs) == 1:
            p = tf.paragraphs[0]
        elif tf.paragraphs[0].text and p is tf.paragraphs[0]:
            p = tf.add_paragraph()
    p.clear() if hasattr(p, "clear") else None
    # Safer: always use first empty or new
    return p


def write_paragraph(tf, text, *, size=14, bold=False, color=INK, align=PP_ALIGN.LEFT, space_after=Pt(4), level=0):
    if tf.paragraphs[0].text == "" and len(tf.paragraphs) == 1:
        p = tf.paragraphs[0]
    else:
        p = tf.add_paragraph()
    p.text = text
    p.alignment = align
    p.level = level
    p.space_after = space_after
    for run in p.runs:
        _set_run(run, size=size, bold=bold, color=color)
    return p


def add_shape(slide, left, top, width, height, fill=None, line=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.line.fill.background()
    if fill is not None:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()
    if line is not None:
        shape.line.color.rgb = line
    return shape


def add_round(slide, left, top, width, height, fill):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.adjustments[0] = 0.08
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    return shape


def textbox(slide, left, top, width, height):
    return slide.shapes.add_textbox(left, top, width, height)


def header_bar(slide, title, subtitle=""):
    add_shape(slide, 0, 0, SLIDE_W, Inches(0.08), fill=ORANGE)
    add_shape(slide, 0, Inches(0.08), SLIDE_W, Inches(0.72), fill=BG)
    box = textbox(slide, MARGIN, Inches(0.18), Inches(3.2), Inches(0.5))
    tf = box.text_frame
    tf.word_wrap = True
    write_paragraph(tf, "ПИК Комфорт", size=16, bold=True, color=ORANGE)
    tbox = textbox(slide, Inches(3.6), Inches(0.16), Inches(9.2), Inches(0.55))
    tf = tbox.text_frame
    tf.word_wrap = True
    write_paragraph(tf, title, size=18, bold=True, color=INK, space_after=Pt(0))
    if subtitle:
        write_paragraph(tf, subtitle, size=11, color=GRAY, space_after=Pt(0))


def footer(slide, page_no, total):
    box = textbox(slide, MARGIN, Inches(7.1), Inches(10), Inches(0.3))
    tf = box.text_frame
    write_paragraph(tf, "Testograf PRO+ · Playbook клиентских опросов", size=10, color=GRAY_LIGHT)
    nbox = textbox(slide, Inches(11.5), Inches(7.1), Inches(1.3), Inches(0.3))
    tf = nbox.text_frame
    write_paragraph(tf, f"{page_no} / {total}", size=10, color=GRAY_LIGHT, align=PP_ALIGN.RIGHT)


def new_slide(prs):
    blank = prs.slide_layouts[6]
    return prs.slides.add_slide(blank)


def card(slide, left, top, width, height, title, lines, accent=VIOLET):
    add_round(slide, left, top, width, height, WHITE)
    # left accent
    add_shape(slide, left, top, Inches(0.08), height, fill=accent)
    box = textbox(slide, left + Inches(0.2), top + Inches(0.12), width - Inches(0.3), height - Inches(0.2))
    tf = box.text_frame
    tf.word_wrap = True
    write_paragraph(tf, title, size=13, bold=True, color=INK, space_after=Pt(6))
    for line in lines:
        write_paragraph(tf, line, size=11, color=GRAY, space_after=Pt(3))


def build():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    slides_meta = []  # filled as we go; footer applied in second pass via list of (slide, idx)

    built = []

    def finish(slide, title, subtitle=""):
        built.append((slide, title, subtitle))

    # --- 1 Title ---
    s = new_slide(prs)
    add_shape(s, 0, 0, SLIDE_W, SLIDE_H, fill=BG)
    add_shape(s, 0, 0, SLIDE_W, Inches(0.12), fill=ORANGE)
    add_shape(s, 0, Inches(6.9), SLIDE_W, Inches(0.6), fill=VIOLET_DARK)
    box = textbox(s, MARGIN, Inches(2.0), Inches(12), Inches(3))
    tf = box.text_frame
    tf.word_wrap = True
    write_paragraph(tf, "ПИК Комфорт", size=18, bold=True, color=ORANGE)
    write_paragraph(tf, "Playbook Testograf", size=36, bold=True, color=INK, space_after=Pt(10))
    write_paragraph(
        tf,
        "Клиентские опросы: NPS · CSI · CSAT · голосования · послесервис",
        size=16,
        color=GRAY,
        space_after=Pt(14),
    )
    write_paragraph(
        tf,
        "Процесс запуска · настройки сервиса · контроль качества ответов",
        size=14,
        color=VIOLET_DARK,
    )
    f = textbox(s, MARGIN, Inches(7.05), Inches(12), Inches(0.35))
    write_paragraph(f.text_frame, "Тариф PRO+ · Согласование запуска: Голубева Екатерина Анатольевна · 2026", size=12, color=WHITE)
    finish(s, "title")

    # --- 2 Agenda ---
    s = new_slide(prs)
    header_bar(s, "Содержание", "Операционный playbook")
    items = [
        "01  Как мы работаем с опросами (шаги 1–9)",
        "02  Типы опросов и риски качества",
        "03  Ключевые настройки Testograf: вкл / выкл",
        "04  Матрица: тип опроса × настройка",
        "05  Постобработка и «блуждающие» ответы",
        "06  Чеклисты и типичные ошибки",
    ]
    box = textbox(s, MARGIN, Inches(1.2), Inches(12), Inches(5.5))
    tf = box.text_frame
    tf.word_wrap = True
    for it in items:
        write_paragraph(tf, it, size=20, color=INK, space_after=Pt(14))
    finish(s, "Содержание")

    # --- 3 Process 1-5 ---
    s = new_slide(prs)
    header_bar(s, "Как мы работаем с опросами", "Шаги 1–5 · до согласования заказчиком")
    steps = [
        ("1", "Бриф", "Заказчик приносит бриф. Если заказчик — мы, бриф готовим сами."),
        ("2", "Заводим опрос", "Создание → Логика → Настройка → Дизайн в Testograf."),
        ("3", "Снимаем ограничение повторов", "Выключаем «нельзя пройти дважды» — иначе нельзя тестировать."),
        ("4", "Внутренняя проверка", "Логика, ошибки, тексты, мобильный вид."),
        ("5", "Проверка заказчиком", "Фиксируем список правок."),
    ]
    y = Inches(1.15)
    for num, title, body in steps:
        add_round(s, MARGIN, y, Inches(12.4), Inches(0.95), WHITE)
        badge = add_round(s, MARGIN + Inches(0.15), y + Inches(0.2), Inches(0.55), Inches(0.55), ORANGE)
        bt = textbox(s, MARGIN + Inches(0.15), y + Inches(0.28), Inches(0.55), Inches(0.45))
        write_paragraph(bt.text_frame, num, size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        t = textbox(s, MARGIN + Inches(0.9), y + Inches(0.15), Inches(11), Inches(0.7))
        tf = t.text_frame
        tf.word_wrap = True
        write_paragraph(tf, title, size=14, bold=True, color=INK, space_after=Pt(2))
        write_paragraph(tf, body, size=12, color=GRAY, space_after=Pt(0))
        y += Inches(1.05)
    finish(s, "Процесс 1–5")

    # --- 4 Process 6-9 ---
    s = new_slide(prs)
    header_bar(s, "Как мы работаем с опросами", "Шаги 6–9 · согласование и запуск")
    steps = [
        ("6", "Коммуникация", "Параллельно готовим текст push / ТГ / email / ЛК."),
        ("7", "Согласование", "Опрос и коммуникацию утверждает Голубева Екатерина Анатольевна. Без этого — запуск запрещен."),
        ("8", "Зачистка + защита", "Удаляем тестовые ответы. Включаем ограничение повторного прохождения."),
        ("9", "Запуск", "Период проведения, каналы, фиксация даты старта."),
    ]
    y = Inches(1.15)
    for num, title, body in steps:
        accent = VIOLET if num != "7" else ORANGE
        add_round(s, MARGIN, y, Inches(12.4), Inches(1.05), WHITE)
        add_round(s, MARGIN + Inches(0.15), y + Inches(0.25), Inches(0.55), Inches(0.55), accent)
        bt = textbox(s, MARGIN + Inches(0.15), y + Inches(0.33), Inches(0.55), Inches(0.45))
        write_paragraph(bt.text_frame, num, size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        t = textbox(s, MARGIN + Inches(0.9), y + Inches(0.18), Inches(11), Inches(0.75))
        tf = t.text_frame
        tf.word_wrap = True
        write_paragraph(tf, title, size=14, bold=True, color=INK, space_after=Pt(2))
        write_paragraph(tf, body, size=12, color=GRAY, space_after=Pt(0))
        y += Inches(1.15)

    note = textbox(s, MARGIN, Inches(5.9), Inches(12.4), Inches(0.9))
    tf = note.text_frame
    tf.word_wrap = True
    write_paragraph(tf, "Правило про повторы", size=13, bold=True, color=ORANGE, space_after=Pt(4))
    write_paragraph(
        tf,
        "Шаги 3–7: ограничение ВЫКЛ. Шаг 8 перед запуском: ограничение ВКЛ. "
        "После бага: выкл → фикс → зачистка тестов → снова вкл → продолжить сбор.",
        size=12,
        color=GRAY,
    )
    finish(s, "Процесс 6–9")

    # --- 5 Types ---
    s = new_slide(prs)
    header_bar(s, "Типы опросов и риски", "Что защищаем в первую очередь")
    rows = [
        ("NPS", "Лояльность 0–10", "Повторы, незавершенные, случайные клики, малые n", "1 чел = 1 ответ; только завершенные"),
        ("CSI", "Удовлетворенность по факторам", "Straight-line по матрицам", "Обязательные шкалы; скринер ЦА"),
        ("CSAT", "Оценка касания", "Дубли по повторным касаниям", "Перс. ссылка / ID сделки"),
        ("Голосование", "Выбор варианта", "Накрутки, боты", "SMS / одноразовые ссылки; без лидерборда"),
        ("Послесервис", "Обратная связь по обращению", "Не та ЦА, повтор одной заявки", "ID обращения; дисквал «не обращался»"),
    ]
    # table header
    headers = ["Тип", "Цель", "Риски", "Защита"]
    col_w = [Inches(1.8), Inches(3.0), Inches(3.8), Inches(3.6)]
    x0 = MARGIN
    y0 = Inches(1.15)
    # header row
    x = x0
    for i, h in enumerate(headers):
        add_shape(s, x, y0, col_w[i], Inches(0.42), fill=VIOLET_DARK)
        tb = textbox(s, x + Inches(0.08), y0 + Inches(0.08), col_w[i] - Inches(0.1), Inches(0.3))
        write_paragraph(tb.text_frame, h, size=11, bold=True, color=WHITE)
        x += col_w[i]
    y = y0 + Inches(0.42)
    for ri, row in enumerate(rows):
        fill = WHITE if ri % 2 == 0 else BG
        x = x0
        for i, cell in enumerate(row):
            add_shape(s, x, y, col_w[i], Inches(0.95), fill=fill)
            tb = textbox(s, x + Inches(0.08), y + Inches(0.12), col_w[i] - Inches(0.12), Inches(0.75))
            tf = tb.text_frame
            tf.word_wrap = True
            write_paragraph(tf, cell, size=11, bold=(i == 0), color=INK if i == 0 else GRAY)
            x += col_w[i]
        y += Inches(0.95)
    finish(s, "Типы опросов")

    # --- 6 Repeats ---
    s = new_slide(prs)
    header_bar(s, "Ограничение повторного прохождения", "Устройство / cookies · IP · персональные ссылки")
    card(
        s,
        MARGIN,
        Inches(1.15),
        Inches(4.0),
        Inches(5.4),
        "Устройство / cookies",
        [
            "Где: Настройка → Основное → Вход",
            "Зачем: 1 человек ≈ 1 ответ",
            "Вкл: шаг 8 перед запуском",
            "Выкл: шаги 3–7 (тест)",
            "Риск: обход через инкогнито / другое устройство",
            "Наш default для NPS/CSI/CSAT",
        ],
        accent=ORANGE,
    )
    card(
        s,
        Inches(4.65),
        Inches(1.15),
        Inches(4.0),
        Inches(5.4),
        "Ограничение по IP",
        [
            "Где: Настройка → Основное → Вход",
            "Зачем: умеренная защита от накруток",
            "Вкл: открытые голосования без перс. ссылок",
            "Выкл: жители ЖК, общий Wi‑Fi / NAT",
            "Риск: один IP = много честных людей",
            "Для NPS ЖК — обычно ВЫКЛ",
        ],
        accent=VIOLET,
    )
    card(
        s,
        Inches(8.85),
        Inches(1.15),
        Inches(4.0),
        Inches(5.4),
        "Персональные ссылки",
        [
            "Где: параметры ссылки / рассылка",
            "Зачем: максимальный контроль участника",
            "Вкл: CSAT, послесервис с ID",
            "Выкл: широкая открытая ссылка в ТГ/QR",
            "Риск: пересылка ссылки другому",
            "PRO+: до 50 параметров ссылки",
        ],
        accent=GREEN,
    )
    finish(s, "Повторы")

    # --- 7 Incomplete + random ---
    s = new_slide(prs)
    header_bar(s, "Незавершенные и рандомизация", "«Блуждающие» ответы — два разных смысла")
    card(
        s,
        MARGIN,
        Inches(1.15),
        Inches(6.1),
        Inches(5.4),
        "Незавершенные ответы (PRO+)",
        [
            "Где: Настройка → Основное",
            "Зачем: вернуться и дописать длинный опрос",
            "Вкл: длинные NPS/CSI",
            "Выкл: короткие CSAT и голосования",
            "Риск: частичные попадут в выгрузку",
            "В индекс — только завершенные (NOT_COMPLETED = 0)",
            "Не пересылать уникальную ссылку с черновиком",
        ],
        accent=ORANGE,
    )
    card(
        s,
        Inches(6.8),
        Inches(1.15),
        Inches(6.0),
        Inches(5.4),
        "Рандомизация (случайный порядок)",
        [
            "Где: страница / вопрос → случайный порядок",
            "Зачем: снять эффект первого варианта",
            "Вкл: равноправные списки факторов/опций",
            "Выкл: шкала NPS 0–10; логические цепочки страниц",
            "Риск: сложнее тестировать логику",
            "Быстрые случайные клики ≠ рандомизация:",
            "их ловим постобработкой (скорость, straight-line)",
        ],
        accent=VIOLET,
    )
    finish(s, "Незавершенные / рандом")

    # --- 8 More settings ---
    s = new_slide(prs)
    header_bar(s, "Еще настройки качества", "Логика · SMS · таймер · итоги · API")
    cards = [
        (
            "Логика и дисквал",
            [
                "Скринер ЦА, ветки «пользовался / нет»",
                "Дисквалифицированных не в индекс",
                "Все ветки прогнать на шаге 4",
            ],
        ),
        (
            "SMS-подтверждение",
            [
                "Для голосований с риском накрутки",
                "NPS/CSI жителей — обычно выкл",
                "Нужен кабинет SMS Профи",
            ],
        ),
        (
            "Таймер",
            [
                "Для UX-тестов — ок",
                "Для NPS/CSI/CSAT — ВЫКЛ",
                "Давление по времени → случайные клики",
            ],
        ),
        (
            "Показ итогов",
            [
                "Во время сбора — ВЫКЛ",
                "Иначе бандвагон-эффект",
                "Публиковать после закрытия",
            ],
        ),
        (
            "Обязательность",
            [
                "Ключ метрики и сегмент — обязательны",
                "Открытые комментарии — по желанию",
                "Слишком много обязательных → отказы",
            ],
        ),
        (
            "API / веб-хуки",
            [
                "PRO+: 120/мин, 500/сутки",
                "Учитывать HTTP 429 в синке",
                "Ключ только в секретах",
            ],
        ),
    ]
    positions = [
        (MARGIN, Inches(1.15)),
        (Inches(4.65), Inches(1.15)),
        (Inches(8.85), Inches(1.15)),
        (MARGIN, Inches(4.0)),
        (Inches(4.65), Inches(4.0)),
        (Inches(8.85), Inches(4.0)),
    ]
    for (left, top), (title, lines) in zip(positions, cards):
        card(s, left, top, Inches(4.0), Inches(2.6), title, lines)
    finish(s, "Прочие настройки")

    # --- 9 Matrix ---
    s = new_slide(prs)
    header_bar(s, "Матрица: тип × настройка", "Вкл / Выкл / Ситуативно")
    headers = ["Настройка", "NPS", "CSI", "CSAT", "Голос.", "Послесервис"]
    data = [
        ["Повтор с устройства*", "Вкл", "Вкл", "Вкл", "Вкл", "Вкл"],
        ["Ограничение по IP", "Выкл", "Выкл", "Сит.", "Сит.", "Выкл"],
        ["Перс. ссылка / ID", "Сит.", "Сит.", "Вкл", "Сит.", "Вкл"],
        ["SMS-подтверждение", "Выкл", "Выкл", "Выкл", "Сит.", "Выкл"],
        ["Незавершенные", "Сит.", "Сит.", "Выкл", "Выкл", "Выкл"],
        ["Рандом шкалы NPS", "Выкл", "—", "—", "—", "—"],
        ["Рандом списков", "Сит.", "Сит.", "Сит.", "Сит.", "Сит."],
        ["Таймер опроса", "Выкл", "Выкл", "Выкл", "Выкл", "Выкл"],
        ["Итоги во время сбора", "Выкл", "Выкл", "Выкл", "Выкл", "Выкл"],
    ]
    col_w = [Inches(3.2)] + [Inches(1.8)] * 5
    x0, y0 = MARGIN, Inches(1.1)
    row_h = Inches(0.48)
    x = x0
    for i, h in enumerate(headers):
        add_shape(s, x, y0, col_w[i], row_h, fill=VIOLET_DARK)
        tb = textbox(s, x + Inches(0.06), y0 + Inches(0.1), col_w[i] - Inches(0.08), Inches(0.3))
        write_paragraph(tb.text_frame, h, size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER if i else PP_ALIGN.LEFT)
        x += col_w[i]
    y = y0 + row_h
    for ri, row in enumerate(data):
        x = x0
        fill = WHITE if ri % 2 == 0 else BG
        for i, cell in enumerate(row):
            add_shape(s, x, y, col_w[i], row_h, fill=fill)
            color = INK
            if cell == "Вкл":
                color = GREEN
            elif cell == "Выкл":
                color = RED
            elif cell == "Сит.":
                color = ORANGE
            tb = textbox(s, x + Inches(0.06), y + Inches(0.1), col_w[i] - Inches(0.08), Inches(0.3))
            write_paragraph(
                tb.text_frame,
                cell,
                size=11,
                bold=(i == 0 or cell in ("Вкл", "Выкл", "Сит.")),
                color=color,
                align=PP_ALIGN.CENTER if i else PP_ALIGN.LEFT,
            )
            x += col_w[i]
        y += row_h
    note = textbox(s, MARGIN, Inches(6.5), Inches(12), Inches(0.4))
    write_paragraph(
        note.text_frame,
        "* Повтор с устройства: Вкл только перед боевым запуском (шаг 8). На тесте (шаги 3–7) — Выкл.",
        size=11,
        color=GRAY,
    )
    finish(s, "Матрица")

    # --- 10 Postprocess ---
    s = new_slide(prs)
    header_bar(s, "Постобработка", "До публикации индексов")
    blocks = [
        ("Базовый фильтр", ["Удалить тестовые", "Только завершенные", "Валидный сегмент (ЖК и т.п.)", "Учет времени — MSK"]),
        ("Скорость и «блуждания»", ["Слишком короткое время", "Straight-line по матрице", "Мусор в открытых", "Флаг low_quality, не в индекс"]),
        ("Дубли и малые n", ["1 ID / обращение = 1 ответ", "Дедуп между каналами", "n < 30 — пометка «гипотеза»", "Голосования: только подтвержденные"]),
    ]
    x = MARGIN
    for title, lines in blocks:
        card(s, x, Inches(1.2), Inches(4.0), Inches(5.2), title, lines, accent=ORANGE if title.startswith("Б") else VIOLET)
        x += Inches(4.2)
    finish(s, "Постобработка")

    # --- 11 Checklists ---
    s = new_slide(prs)
    header_bar(s, "Чеклисты", "Перед проверкой · перед Голубевой Е. А. · перед запуском")
    cols = [
        (
            "Перед внутренней проверкой",
            [
                "□ Бриф есть",
                "□ Опрос создан",
                "□ Повторы ВЫКЛ",
                "□ Ветки в предпросмотре",
                "□ Мобильный вид",
                "□ NPS-шкала без рандома",
            ],
        ),
        (
            "Перед согласованием Е. А.",
            [
                "□ Внутренняя команда ОК",
                "□ Заказчик ОК",
                "□ Черновик коммуникации",
                "□ Каналы и даты",
                "□ Опрос + комм. вместе",
            ],
        ),
        (
            "Перед запуском",
            [
                "□ Согласование Голубевой Е. А.",
                "□ Тесты удалены",
                "□ Повторы ВКЛ",
                "□ Итоги респондентам ВЫКЛ",
                "□ Боевая ссылка в коммуникации",
                "□ API-бюджет 500/сутки учтен",
            ],
        ),
    ]
    x = MARGIN
    for title, lines in cols:
        card(s, x, Inches(1.15), Inches(4.0), Inches(5.5), title, lines, accent=ORANGE)
        x += Inches(4.2)
    finish(s, "Чеклисты")

    # --- 12 Errors ---
    s = new_slide(prs)
    header_bar(s, "Типичные ошибки", "Что ломает репрезентативность")
    errs = [
        ("Забыли выкл повторы на тесте", "Команда не может перепройти → шаг 3"),
        ("Забыли вкл повторы перед запуском", "Дубли и накрутки → шаг 8"),
        ("Не зачистили тестовые", "Искажение NPS/CSI → зачистка на шаге 8"),
        ("IP-лимит для жителей ЖК", "Много честных отсечено → устройство / перс. ссылка"),
        ("Рандом шкалы NPS", "Невалидная метрика → шкалу 0–10 не рандомизировать"),
        ("Таймер на длинном NPS", "Случайные клики → таймер выкл"),
        ("Незавершенные в индексе", "Шум в метрике → только завершенные"),
        ("Лидерборд во время сбора", "Бандвагон → итоги после закрытия"),
        ("Запуск без Голубевой Е. А.", "Нарушение процесса → шаг 7 обязателен"),
    ]
    y = Inches(1.1)
    for title, fix in errs:
        add_shape(s, MARGIN, y, Inches(12.4), Inches(0.55), fill=WHITE)
        add_shape(s, MARGIN, y, Inches(0.08), Inches(0.55), fill=RED)
        tb = textbox(s, MARGIN + Inches(0.2), y + Inches(0.08), Inches(12), Inches(0.45))
        tf = tb.text_frame
        tf.word_wrap = True
        write_paragraph(tf, f"{title}  →  {fix}", size=12, color=INK, space_after=Pt(0))
        y += Inches(0.58)
    finish(s, "Ошибки")

    # --- 13 Wandering cheat ---
    s = new_slide(prs)
    header_bar(s, "Шпаргалка: «блуждающие» ответы", "Три разных явления — три разных действия")
    card(
        s,
        MARGIN,
        Inches(1.2),
        Inches(4.0),
        Inches(5.2),
        "1. Случайный порядок в UI",
        [
            "Это рандомизация в Testograf",
            "Вкл: равноправные списки",
            "Выкл: шкала NPS, логика страниц",
            "См. Настройка вопроса / страницы",
        ],
        accent=VIOLET,
    )
    card(
        s,
        Inches(4.65),
        Inches(1.2),
        Inches(4.0),
        Inches(5.2),
        "2. Быстрые случайные клики",
        [
            "Не настройка сервиса",
            "Профилактика: без таймера, нормальная длина",
            "Лечение: постобработка по времени",
            "и straight-line",
        ],
        accent=ORANGE,
    )
    card(
        s,
        Inches(8.85),
        Inches(1.2),
        Inches(4.0),
        Inches(5.2),
        "3. Бросили анкету на полпути",
        [
            "Незавершенные (PRO+) + напоминания",
            "Короче анкета",
            "В индекс — только завершенные",
            "Drop-off смотреть отдельно",
        ],
        accent=GREEN,
    )
    finish(s, "Блуждающие")

    # --- 14 Close ---
    s = new_slide(prs)
    add_shape(s, 0, 0, SLIDE_W, SLIDE_H, fill=BG)
    add_shape(s, 0, 0, SLIDE_W, Inches(0.12), fill=ORANGE)
    box = textbox(s, MARGIN, Inches(2.2), Inches(12), Inches(3.5))
    tf = box.text_frame
    tf.word_wrap = True
    write_paragraph(tf, "Итог", size=16, bold=True, color=ORANGE)
    write_paragraph(tf, "Сначала процесс, потом настройки", size=28, bold=True, color=INK, space_after=Pt(12))
    write_paragraph(
        tf,
        "Тест без ограничения повторов → согласование Голубевой Е. А. →\n"
        "зачистка тестов → включение защиты → запуск → чистая аналитика.",
        size=16,
        color=GRAY,
        space_after=Pt(18),
    )
    write_paragraph(tf, "Полный текст: docs/testograf-playbook-client-surveys.md", size=13, color=VIOLET_DARK)
    finish(s, "Итог")

    # footers (skip title and final if desired — add to all content slides)
    total = len(prs.slides)
    for i, slide in enumerate(prs.slides, start=1):
        if i == 1 or i == total:
            continue
        footer(slide, i, total)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(OUT))
    print(f"Saved: {OUT}")
    return OUT


if __name__ == "__main__":
    build()
