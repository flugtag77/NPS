#!/usr/bin/env python3
"""Регламент работы с Testograf (клиентские опросы) в формате PPTX.

Факты о сервисе взяты из официальных материалов Testograf
(Функционал, Помощь, Мануал, блог). Полные таблицы со ссылками:
docs/testograf-catalog-types-and-settings.md

В рабочих текстах не используется буква с точками над е.
"""
from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

ORANGE = RGBColor(0xFC, 0x4C, 0x02)
ORANGE_LIGHT = RGBColor(0xFF, 0xEE, 0xE6)
VIOLET = RGBColor(0x96, 0x78, 0xD2)
VIOLET_DARK = RGBColor(0x7C, 0x5E, 0xB8)
VIOLET_LIGHT = RGBColor(0xEF, 0xEA, 0xF8)
INK = RGBColor(0x1A, 0x1A, 0x1A)
GRAY = RGBColor(0x5A, 0x5A, 0x5A)
GRAY_MID = RGBColor(0x9A, 0x9A, 0x9A)
LINE = RGBColor(0xDD, 0xD8, 0xE8)
BG = RGBColor(0xF7, 0xF5, 0xFB)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREEN = RGBColor(0x2F, 0xA8, 0x4F)
GREEN_LIGHT = RGBColor(0xE6, 0xF5, 0xEA)
YELLOW = RGBColor(0xF5, 0xB8, 0x00)
RED = RGBColor(0xE2, 0x4C, 0x4B)
RED_LIGHT = RGBColor(0xFC, 0xEA, 0xEA)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
M = Inches(0.5)
CONTENT_W = SLIDE_W - 2 * M

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "docs" / "Testograf_playbook_client_surveys.pptx"

APPROVER = "Голубева Екатерина Анатольевна"


# ---------------------------------------------------------------- text utils
def _style_run(run, size, bold, color):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = "Arial"
    rpr = run._r.get_or_add_rPr()
    for tag in ("latin", "ea", "cs"):
        el = rpr.find(qn(f"a:{tag}"))
        if el is None:
            el = rpr.makeelement(qn(f"a:{tag}"), {})
            rpr.append(el)
        el.set("typeface", "Arial")


def text(slide, x, y, w, h, paras, size=13, color=INK, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT):
    """paras: list of str | dict(text|runs, size, bold, color, after, align)."""
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.04)
    tf.margin_top = tf.margin_bottom = Inches(0.02)
    first = True
    for item in paras:
        if isinstance(item, str):
            item = {"text": item}
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = item.get("align", align)
        p.space_after = Pt(item.get("after", 6))
        p_size = item.get("size", size)
        p_color = item.get("color", color)
        p_bold = item.get("bold", False)
        runs = item.get("runs") or [(item.get("text", ""), {})]
        for run_text, opts in runs:
            r = p.add_run()
            r.text = run_text
            _style_run(
                r,
                opts.get("size", p_size),
                opts.get("bold", p_bold),
                opts.get("color", p_color),
            )
    return box


def rect(slide, x, y, w, h, fill, shape=MSO_SHAPE.RECTANGLE, line=None):
    sh = slide.shapes.add_shape(shape, x, y, w, h)
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        sh.adjustments[0] = 0.08
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(1)
    sh.shadow.inherit = False
    return sh


def label_line(label, body, label_color=VIOLET_DARK, after=7, size=12.5):
    return {"runs": [(label + " ", {"bold": True, "color": label_color}), (body, {})], "after": after, "size": size}


def bullet(t, after=5, size=12.5, color=INK):
    return {"runs": [("•  ", {"color": ORANGE, "bold": True}), (t, {})], "after": after, "size": size, "color": color}


# ---------------------------------------------------------------- deck
class Deck:
    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width = SLIDE_W
        self.prs.slide_height = SLIDE_H
        self.section = ""
        self.slides: list[tuple[object, str]] = []

    def new(self):
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.slides.append((s, self.section))
        return s

    def header(self, s, title, lead=None):
        rect(s, 0, 0, SLIDE_W, Inches(0.08), ORANGE)
        text(s, M, Inches(0.3), CONTENT_W, Inches(0.6), [{"text": title, "size": 24, "bold": True}])
        if lead:
            text(s, M, Inches(0.92), CONTENT_W, Inches(0.7), [{"text": lead, "size": 14, "color": GRAY}])

    def finish(self):
        total = len(self.slides)
        for i, (s, section) in enumerate(self.slides, start=1):
            if i == 1:
                continue
            rect(s, M, Inches(7.02), CONTENT_W, Pt(0.75), LINE)
            label = "Регламент Testograf" + (f"  ·  {section}" if section else "")
            text(s, M, Inches(7.07), Inches(10), Inches(0.3), [{"text": label, "size": 10, "color": GRAY_MID}])
            text(
                s,
                SLIDE_W - M - Inches(1.5),
                Inches(7.07),
                Inches(1.5),
                Inches(0.3),
                [{"text": f"{i} / {total}", "size": 10, "color": GRAY_MID, "align": PP_ALIGN.RIGHT}],
            )
        OUT.parent.mkdir(parents=True, exist_ok=True)
        self.prs.save(str(OUT))
        return OUT, total


def card(s, x, y, w, h, title, body, accent=VIOLET, title_size=15, fill=WHITE):
    rect(s, x, y, w, h, fill, MSO_SHAPE.ROUNDED_RECTANGLE, line=LINE)
    rect(s, x + Inches(0.2), y + Inches(0.18), Inches(0.5), Pt(3), accent)
    text(s, x + Inches(0.18), y + Inches(0.28), w - Inches(0.36), Inches(0.5), [{"text": title, "size": title_size, "bold": True}])
    text(s, x + Inches(0.18), y + Inches(0.72), w - Inches(0.36), h - Inches(0.85), body)


def callout(s, x, y, w, h, title, body, tone="orange"):
    fill, accent = {
        "orange": (ORANGE_LIGHT, ORANGE),
        "violet": (VIOLET_LIGHT, VIOLET_DARK),
        "green": (GREEN_LIGHT, GREEN),
        "red": (RED_LIGHT, RED),
    }[tone]
    rect(s, x, y, w, h, fill, MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x, y + Inches(0.12), Inches(0.07), h - Inches(0.24), accent)
    paras = []
    if title:
        paras.append({"text": title, "bold": True, "size": 13, "color": accent, "after": 4})
    paras.extend(body if isinstance(body, list) else [{"text": body, "size": 12.5}])
    text(s, x + Inches(0.22), y + Inches(0.1), w - Inches(0.35), h - Inches(0.2), paras, anchor=MSO_ANCHOR.MIDDLE)


def section_slide(deck, number, title, subtitle):
    deck.section = title
    s = deck.new()
    rect(s, 0, 0, SLIDE_W, SLIDE_H, VIOLET_DARK)
    rect(s, 0, 0, SLIDE_W, Inches(0.08), ORANGE)
    text(s, M, Inches(2.3), Inches(3), Inches(1.2), [{"text": f"{number:02d}", "size": 66, "bold": True, "color": ORANGE}])
    text(s, M, Inches(3.55), Inches(11), Inches(0.9), [{"text": title, "size": 34, "bold": True, "color": WHITE}])
    text(s, M, Inches(4.4), Inches(11), Inches(1.2), [{"text": subtitle, "size": 17, "color": VIOLET_LIGHT}])
    return s


# ---------------------------------------------------------------- mockups
def radio(s, x, y, label, checked=False, square=False):
    shape = MSO_SHAPE.ROUNDED_RECTANGLE if square else MSO_SHAPE.OVAL
    sh = rect(s, x, y + Inches(0.03), Inches(0.2), Inches(0.2), WHITE, shape, line=GRAY_MID)
    if checked:
        inner = Inches(0.1)
        rect(s, x + Inches(0.05), y + Inches(0.08), inner, inner, VIOLET_DARK, shape)
    text(s, x + Inches(0.3), y - Inches(0.02), Inches(3.5), Inches(0.3), [{"text": label, "size": 12}])
    return sh


def mock_frame(s, x, y, w, h, question):
    rect(s, x, y, w, h, WHITE, MSO_SHAPE.ROUNDED_RECTANGLE, line=LINE)
    text(s, x + Inches(0.2), y + Inches(0.15), w - Inches(0.4), Inches(0.5), [{"text": question, "size": 12.5, "bold": True}])


def nps_scale(s, x, y, cell=Inches(0.44)):
    for i in range(11):
        color = RED if i <= 6 else (YELLOW if i <= 8 else GREEN)
        rect(s, x + i * (cell + Inches(0.05)), y, cell, cell, color, MSO_SHAPE.ROUNDED_RECTANGLE)
        text(
            s,
            x + i * (cell + Inches(0.05)),
            y,
            cell,
            cell,
            [{"text": str(i), "size": 13, "bold": True, "color": WHITE, "align": PP_ALIGN.CENTER}],
            anchor=MSO_ANCHOR.MIDDLE,
        )


def stars(s, x, y, filled=4, n=5, size=Inches(0.32)):
    for i in range(n):
        rect(s, x + i * (size + Inches(0.06)), y, size, size, YELLOW if i < filled else LINE, MSO_SHAPE.STAR_5_POINT)


def arrow_right(s, x, y, w=Inches(0.35), h=Inches(0.28), color=GRAY_MID):
    rect(s, x, y, w, h, color, MSO_SHAPE.RIGHT_ARROW)


def arrow_down(s, x, y, w=Inches(0.28), h=Inches(0.3), color=GRAY_MID):
    rect(s, x, y, w, h, color, MSO_SHAPE.DOWN_ARROW)


def flow_box(s, x, y, w, h, title, body, fill=WHITE, accent=VIOLET, title_color=INK):
    rect(s, x, y, w, h, fill, MSO_SHAPE.ROUNDED_RECTANGLE, line=LINE)
    rect(s, x, y + Inches(0.15), Inches(0.06), h - Inches(0.3), accent)
    text(
        s,
        x + Inches(0.18),
        y + Inches(0.08),
        w - Inches(0.28),
        h - Inches(0.16),
        [{"text": title, "size": 13, "bold": True, "color": title_color, "after": 3}, {"text": body, "size": 11.5, "color": GRAY}],
        anchor=MSO_ANCHOR.MIDDLE,
    )


def table(s, x, y, col_w, headers, rows, row_h=Inches(0.42), size=11.5, header_fill=VIOLET_DARK, first_bold=True, colorize=None):
    cx = x
    for i, h in enumerate(headers):
        rect(s, cx, y, col_w[i], row_h, header_fill)
        text(
            s,
            cx + Inches(0.06),
            y,
            col_w[i] - Inches(0.12),
            row_h,
            [{"text": h, "size": size, "bold": True, "color": WHITE, "align": PP_ALIGN.LEFT if i == 0 else PP_ALIGN.CENTER}],
            anchor=MSO_ANCHOR.MIDDLE,
        )
        cx += col_w[i]
    cy = y + row_h
    for r_i, row in enumerate(rows):
        cx = x
        fill = WHITE if r_i % 2 == 0 else BG
        for c_i, cell in enumerate(row):
            rect(s, cx, cy, col_w[c_i], row_h, fill)
            color = INK
            bold = first_bold and c_i == 0
            if colorize and c_i > 0:
                color, bold = colorize(cell, color, bold)
            text(
                s,
                cx + Inches(0.06),
                cy,
                col_w[c_i] - Inches(0.12),
                row_h,
                [{"text": cell, "size": size, "bold": bold, "color": color, "align": PP_ALIGN.LEFT if c_i == 0 else PP_ALIGN.CENTER}],
                anchor=MSO_ANCHOR.MIDDLE,
            )
            cx += col_w[c_i]
        cy += row_h
    return cy


def text_table(s, x, y, col_w, headers, rows, row_h, size=11.5):
    """Таблица с длинным текстом в ячейках, выравнивание по левому краю."""
    cx = x
    for i, h in enumerate(headers):
        rect(s, cx, y, col_w[i], Inches(0.42), VIOLET_DARK)
        text(s, cx + Inches(0.08), y, col_w[i] - Inches(0.16), Inches(0.42), [{"text": h, "size": 12, "bold": True, "color": WHITE}], anchor=MSO_ANCHOR.MIDDLE)
        cx += col_w[i]
    cy = y + Inches(0.42)
    for r_i, row in enumerate(rows):
        h = row_h[r_i] if isinstance(row_h, list) else row_h
        cx = x
        fill = WHITE if r_i % 2 == 0 else BG
        for c_i, cell in enumerate(row):
            rect(s, cx, cy, col_w[c_i], h, fill)
            text(
                s,
                cx + Inches(0.08),
                cy + Inches(0.04),
                col_w[c_i] - Inches(0.16),
                h - Inches(0.08),
                [{"text": cell, "size": size, "bold": c_i == 0, "color": INK if c_i == 0 else GRAY}],
                anchor=MSO_ANCHOR.MIDDLE,
            )
            cx += col_w[c_i]
        cy += h
    return cy


# ---------------------------------------------------------------- process
STEPS = [
    ("Бриф", "Заказчик / мы"),
    ("Заводим опрос", "Ответственный"),
    ("Отключаем ограничение повторов", "Ответственный"),
    ("Проверка командой", "Внутренняя команда"),
    ("Проверка заказчиком", "Заказчик"),
    ("Коммуникация", "Параллельно с 4–5"),
    ("Согласование", "Голубева Е. А."),
    ("Зачистка + включение ограничения", "Ответственный"),
    ("Запуск", "Ответственный"),
]


def step_progress(s, current):
    x0 = SLIDE_W - M - Inches(9 * 0.42)
    for i in range(9):
        n = i + 1
        fill = ORANGE if n == current else (VIOLET_DARK if n < current else LINE)
        rect(s, x0 + i * Inches(0.42), Inches(0.42), Inches(0.32), Inches(0.32), fill, MSO_SHAPE.OVAL)
        text(
            s,
            x0 + i * Inches(0.42),
            Inches(0.42),
            Inches(0.32),
            Inches(0.32),
            [{"text": str(n), "size": 10, "bold": True, "color": WHITE if n <= current else GRAY, "align": PP_ALIGN.CENTER}],
            anchor=MSO_ANCHOR.MIDDLE,
        )


def step_slide(deck, n, title, who, blocks, note=None, note_tone="orange", note_title=None):
    s = deck.new()
    rect(s, 0, 0, SLIDE_W, Inches(0.08), ORANGE)
    step_progress(s, n)
    text(s, M, Inches(0.3), Inches(8), Inches(0.5), [{"text": f"Шаг {n} из 9", "size": 13, "bold": True, "color": ORANGE}])
    text(s, M, Inches(0.72), Inches(12), Inches(0.7), [{"text": title, "size": 26, "bold": True}])
    text(s, M, Inches(1.35), Inches(12), Inches(0.4), [{"runs": [("Кто выполняет: ", {"bold": True, "color": VIOLET_DARK}), (who, {})], "size": 14}])

    note_h = Inches(1.05) if note else 0
    area_top = Inches(1.95)
    area_bottom = Inches(6.85) - note_h - (Inches(0.15) if note else 0)
    gap = Inches(0.2)
    cols = len(blocks)
    col_w = (CONTENT_W - gap * (cols - 1)) / cols
    for i, (btitle, body, accent) in enumerate(blocks):
        card(s, M + i * (col_w + gap), area_top, col_w, area_bottom - area_top, btitle, body, accent=accent, title_size=15)
    if note:
        callout(s, M, Inches(6.85) - note_h, CONTENT_W, note_h, note_title, note, tone=note_tone)
    return s


# ---------------------------------------------------------------- build
def build():
    d = Deck()

    # 1. Титул
    s = d.new()
    rect(s, 0, 0, SLIDE_W, SLIDE_H, BG)
    rect(s, 0, 0, Inches(0.35), SLIDE_H, ORANGE)
    text(s, Inches(1.0), Inches(1.4), Inches(11), Inches(0.5), [{"text": "ПИК Комфорт · Департамент маркетинга", "size": 15, "bold": True, "color": ORANGE}])
    text(s, Inches(1.0), Inches(2.0), Inches(11), Inches(1.6), [{"text": "Регламент работы\nс сервисом Testograf", "size": 42, "bold": True}])
    text(
        s,
        Inches(1.0),
        Inches(3.85),
        Inches(11),
        Inches(1.2),
        [
            {"text": "Клиентские опросы: NPS, CSI, CSAT, голосования и опросы после обращения.", "size": 17, "color": GRAY, "after": 6},
            {"text": "Для всех, кто готовит, проверяет или запускает опрос — даже если открывает Testograf впервые.", "size": 17, "color": GRAY},
        ],
    )
    rect(s, Inches(1.0), Inches(5.55), Inches(5.6), Inches(0.9), WHITE, MSO_SHAPE.ROUNDED_RECTANGLE, line=LINE)
    text(
        s,
        Inches(1.2),
        Inches(5.62),
        Inches(5.3),
        Inches(0.8),
        [
            {"runs": [("Тариф: ", {"bold": True}), ("Годовая лицензия PRO+", {})], "size": 13, "after": 3},
            {"runs": [("Согласование запуска: ", {"bold": True}), (APPROVER, {})], "size": 13},
        ],
        anchor=MSO_ANCHOR.MIDDLE,
    )

    # 2. Содержание
    s = d.new()
    d.header(s, "Что в этом регламенте", "Шесть частей — в том порядке, в котором вы будете работать с опросом.")
    parts = [
        ("01", "Сервис и как он устроен", "Что такое Testograf, для каких опросов мы его используем и из каких частей состоит опрос."),
        ("02", "Как мы запускаем опрос", "Обязательный порядок из 9 шагов: что делаем, кто отвечает, где нажать в сервисе."),
        ("03", "Типы вопросов", "Все виды вопросов и ответов в Testograf: что это, что получаем на выходе, для какого опроса."),
        ("04", "Логика", "Как показывать людям разные вопросы в зависимости от их ответов."),
        ("05", "Настройки", "Что включать и выключать, что это меняет и в каких опросах нужно."),
        ("06", "Сбор ответов и результаты", "Каналы рассылки, отчеты и выгрузки, правила подсчета и чеклист перед запуском."),
    ]
    for i, (num, title, desc) in enumerate(parts):
        col, row = i % 2, i // 2
        x = M + col * (CONTENT_W / 2 + Inches(0.1))
        y = Inches(1.75) + row * Inches(1.65)
        w = CONTENT_W / 2 - Inches(0.1)
        rect(s, x, y, w, Inches(1.45), WHITE, MSO_SHAPE.ROUNDED_RECTANGLE, line=LINE)
        text(s, x + Inches(0.25), y + Inches(0.2), Inches(1), Inches(0.9), [{"text": num, "size": 32, "bold": True, "color": ORANGE}])
        text(
            s,
            x + Inches(1.35),
            y + Inches(0.18),
            w - Inches(1.55),
            Inches(1.15),
            [{"text": title, "size": 16, "bold": True, "after": 4}, {"text": desc, "size": 12.5, "color": GRAY}],
        )

    # ================================================================ 01
    section_slide(d, 1, "Сервис и как он устроен", "Что такое Testograf и из каких частей состоит любой опрос")

    s = d.new()
    d.header(s, "Что такое Testograf", "Российский онлайн-конструктор опросов, анкет, тестов и голосований. В нем мы собираем анкету, рассылаем ее и получаем готовую статистику.")
    card(
        s,
        M,
        Inches(1.85),
        Inches(6.4),
        Inches(4.95),
        "Для каких опросов мы его используем",
        [
            label_line("NPS —", "готовность рекомендовать компанию. Вопрос со шкалой от 0 до 10."),
            label_line("CSI —", "удовлетворенность отдельными параметрами: сервисом, услугой, приложением."),
            label_line("CSAT —", "оценка конкретного взаимодействия с клиентом."),
            label_line("Голосования —", "выбор одного или нескольких вариантов из предложенных."),
            label_line("Опросы после обращения —", "обратная связь по конкретной заявке клиента."),
        ],
        accent=ORANGE,
    )
    card(
        s,
        M + Inches(6.6),
        Inches(1.85),
        CONTENT_W - Inches(6.6),
        Inches(4.95),
        "Что дает наш тариф PRO+",
        [
            bullet("API и веб-хуки: ответы приходят в наши системы сразу после каждого ответа"),
            bullet("Продвинутая логика и продвинутая аналитика (кросс-таблицы)"),
            bullet("Сохранение незавершенных ответов"),
            bullet("SMS-рассылка и SMS-подтверждение ответа (через сервис SMS Профи)"),
            bullet("Офлайн-режим и собственный почтовый сервер (SMTP)"),
            bullet("До 50 параметров ссылки, до 15 000 писем в сутки, 15 ГБ хранилища"),
        ],
        accent=VIOLET,
    )

    s = d.new()
    d.header(s, "Из чего состоит опрос", "Опрос собирается в конструкторе из четырех вкладок. Все изменения сохраняются автоматически — отдельной кнопки «Сохранить» нет.")
    tabs = [
        ("1. Создание", "Страницы, вопросы, варианты ответов. Здесь выбирают тип вопроса и делают его обязательным."),
        ("2. Логика", "Правила переходов: какой вопрос человек увидит следующим, кого завершить или отсеять."),
        ("3. Настройка", "Доступ и ограничения, сроки проведения, приветствие, завершение, таймер, тексты кнопок."),
        ("4. Дизайн", "Внешний вид опроса: оформление под фирменный стиль компании."),
    ]
    bw = (CONTENT_W - Inches(0.45) * 3) / 4
    for i, (t, b) in enumerate(tabs):
        x = M + i * (bw + Inches(0.45))
        flow_box(s, x, Inches(1.95), bw, Inches(1.9), t, b, accent=ORANGE)
        if i < 3:
            arrow_right(s, x + bw + Inches(0.05), Inches(2.75))
    text(s, M, Inches(4.15), CONTENT_W, Inches(0.4), [{"text": "Когда опрос собран, работа идет в карточке опроса в личном кабинете:", "size": 14, "bold": True}])
    after = [
        ("Общая информация", "Число ответов и переходов, доля завершивших, среднее время заполнения. Переключатель «Активный опрос», удаление собранных ответов."),
        ("Сбор ответов", "Прямая ссылка, email- и SMS-рассылка, параметры ссылки, виджет, всплывающее окно, встраивание на сайт."),
        ("Результаты", "Сводный отчет с таблицами и диаграммами, отдельные ответы, фильтры, выгрузка в XLSX, CSV, PDF, SPSS."),
    ]
    bw2 = (CONTENT_W - Inches(0.3) * 2) / 3
    for i, (t, b) in enumerate(after):
        flow_box(s, M + i * (bw2 + Inches(0.3)), Inches(4.6), bw2, Inches(1.35), t, b, accent=VIOLET)
    callout(
        s,
        M,
        Inches(6.1),
        CONTENT_W,
        Inches(0.72),
        None,
        [{"runs": [("Предварительный просмотр. ", {"bold": True}), ("Кнопка в конструкторе открывает опрос в новой вкладке и показывает, как он выглядит на разных устройствах. Пользуйтесь ей на каждом шаге проверки.", {})], "size": 12.5}],
        tone="violet",
    )

    # ================================================================ 02
    section_slide(d, 2, "Как мы запускаем опрос", "Обязательный порядок из 9 шагов. Ни один опрос не публикуется для клиентов, пока не пройдены все шаги.")

    s = d.new()
    d.header(s, "9 шагов от брифа до запуска", "Шаги идут строго по порядку. Шаг 6 выполняется одновременно с шагами 4 и 5.")
    bw = Inches(3.65)
    bh = Inches(1.15)
    gx = Inches(0.55)
    xs = [M, M + bw + gx, M + 2 * (bw + gx)]
    ys = [Inches(1.75), Inches(3.3), Inches(4.85)]
    order = [(0, 0), (0, 1), (0, 2), (1, 2), (1, 1), (1, 0), (2, 0), (2, 1), (2, 2)]
    for idx, (r, c) in enumerate(order):
        n = idx + 1
        title, who = STEPS[idx]
        if n in (3, 8):
            fill, accent, tcolor = ORANGE_LIGHT, ORANGE, INK
        elif n == 7:
            fill, accent, tcolor = VIOLET_LIGHT, VIOLET_DARK, INK
        else:
            fill, accent, tcolor = WHITE, VIOLET, INK
        x, y = xs[c], ys[r]
        rect(s, x, y, bw, bh, fill, MSO_SHAPE.ROUNDED_RECTANGLE, line=LINE)
        rect(s, x + Inches(0.18), y + Inches(0.3), Inches(0.55), Inches(0.55), accent, MSO_SHAPE.OVAL)
        text(s, x + Inches(0.18), y + Inches(0.3), Inches(0.55), Inches(0.55), [{"text": str(n), "size": 18, "bold": True, "color": WHITE, "align": PP_ALIGN.CENTER}], anchor=MSO_ANCHOR.MIDDLE)
        text(
            s,
            x + Inches(0.88),
            y + Inches(0.1),
            bw - Inches(1.0),
            bh - Inches(0.2),
            [{"text": title, "size": 14, "bold": True, "color": tcolor, "after": 3}, {"text": who, "size": 12, "color": GRAY}],
            anchor=MSO_ANCHOR.MIDDLE,
        )
    # arrows row 1 → right, row 2 ← left, row 3 → right, vertical connectors
    for c in range(2):
        arrow_right(s, xs[c] + bw + Inches(0.1), ys[0] + Inches(0.44))
        arrow_right(s, xs[c] + bw + Inches(0.1), ys[2] + Inches(0.44))
    for c in range(2):
        rect(s, xs[c] + bw + Inches(0.1), ys[1] + Inches(0.44), Inches(0.35), Inches(0.28), GRAY_MID, MSO_SHAPE.LEFT_ARROW)
    arrow_down(s, xs[2] + bw / 2 - Inches(0.14), ys[0] + bh + Inches(0.07))
    arrow_down(s, xs[0] + bw / 2 - Inches(0.14), ys[1] + bh + Inches(0.07))
    # legend
    ly = Inches(6.25)
    rect(s, M, ly, Inches(0.3), Inches(0.3), ORANGE_LIGHT, MSO_SHAPE.ROUNDED_RECTANGLE, line=ORANGE)
    text(s, M + Inches(0.4), ly - Inches(0.02), Inches(5.5), Inches(0.4), [{"text": "Шаги 3 и 8 — работа с ограничением повторного прохождения", "size": 12, "color": GRAY}])
    rect(s, M + Inches(6.3), ly, Inches(0.3), Inches(0.3), VIOLET_LIGHT, MSO_SHAPE.ROUNDED_RECTANGLE, line=VIOLET_DARK)
    text(s, M + Inches(6.7), ly - Inches(0.02), Inches(5.6), Inches(0.4), [{"text": "Шаг 7 — без согласования запуск невозможен", "size": 12, "color": GRAY}])

    step_slide(
        d,
        1,
        "Получаем бриф опроса",
        "заказчик. Если заказчик — мы сами, бриф готовим сами",
        [
            (
                "Что делаем",
                [
                    {"text": "Заказчик приходит с готовым брифом опроса.", "size": 13, "after": 8},
                    {"text": "Если опрос нужен нам самим — мы выступаем заказчиком и готовим бриф сами.", "size": 13, "after": 8},
                    {"text": "Опрос в Testograf заводим только по брифу.", "size": 13, "bold": True},
                ],
                ORANGE,
            ),
            (
                "Что взять из брифа",
                [
                    bullet("Какой это опрос: NPS, CSI, CSAT, голосование или опрос после обращения → какие типы вопросов брать (часть 03)"),
                    bullet("Кого опрашиваем и знаем ли контакты заранее → как ограничить повторы (часть 05)"),
                    bullet("Через какие каналы рассылаем → как собирать ответы (часть 06)"),
                    bullet("Сроки → период проведения опроса"),
                    bullet("Анонимный опрос или нужно знать, кто ответил"),
                ],
                VIOLET,
            ),
            (
                "Результат шага",
                [
                    {"text": "Есть бриф, из которого понятны цель, аудитория, каналы и сроки опроса.", "size": 13, "after": 10},
                    {"text": "Можно переходить к шагу 2.", "size": 13, "bold": True, "color": GREEN},
                ],
                GREEN,
            ),
        ],
    )

    step_slide(
        d,
        2,
        "Заводим опрос в Testograf",
        "ответственный за опрос",
        [
            (
                "Что делаем",
                [
                    {"text": "Создаем новый опрос в личном кабинете — с нуля или из готового шаблона Testograf (например, есть шаблон NPS).", "size": 13, "after": 8},
                    {"text": "Переносим вопросы из брифа в конструктор.", "size": 13},
                ],
                ORANGE,
            ),
            (
                "Как сделать в сервисе",
                [
                    label_line("Создание —", "добавить страницы и вопросы, выбрать тип каждого вопроса (часть 03)."),
                    label_line("Логика —", "настроить переходы между вопросами (часть 04)."),
                    label_line("Настройка —", "сроки, приветствие, завершение, доступ (часть 05)."),
                    label_line("Дизайн —", "оформление в фирменном стиле."),
                ],
                VIOLET,
            ),
            (
                "Результат шага",
                [
                    {"text": "Опрос полностью собран, открывается в предварительном просмотре и готов к проверке.", "size": 13, "after": 10},
                    {"text": "Переходим к шагу 3.", "size": 13, "bold": True, "color": GREEN},
                ],
                GREEN,
            ),
        ],
    )

    step_slide(
        d,
        3,
        "Отключаем ограничение повторного прохождения",
        "ответственный за опрос",
        [
            (
                "Что делаем",
                [
                    {"text": "Выключаем все ограничения, которые не дают пройти опрос второй раз.", "size": 13, "after": 8},
                    {"text": "Опрос должен проходиться сколько угодно раз с одного устройства.", "size": 13},
                ],
                ORANGE,
            ),
            (
                "Где в сервисе",
                [
                    {"runs": [("Настройка → Основное → блок «Вход»", {"bold": True})], "size": 13, "after": 8},
                    bullet("ограничение повторного заполнения с устройства — выключено"),
                    bullet("ограничение по IP-адресу — выключено"),
                    {"text": "Подробно о каждом способе ограничения — в части 05.", "size": 12, "color": GRAY},
                ],
                VIOLET,
            ),
            (
                "Зачем",
                [
                    {"text": "На шагах 4 и 5 команда и заказчик проходят опрос много раз: проверяют каждую ветку логики и разные варианты ответов.", "size": 13, "after": 8},
                    {"text": "С включенным ограничением второй раз пройти опрос с того же устройства не получится — проверку не сделать.", "size": 13},
                ],
                GREEN,
            ),
        ],
        note=[{"text": "Все ответы, которые появятся в опросе с этого момента и до шага 8, — тестовые. На шаге 8 их обязательно удаляем.", "size": 13}],
        note_title="Важно",
    )

    step_slide(
        d,
        4,
        "Внутренняя команда проверяет опрос",
        "внутренняя команда",
        [
            (
                "Что делаем",
                [
                    {"text": "Передаем ссылку на опрос внутренней команде на проверку логики и ошибок.", "size": 13, "after": 8},
                    {"text": "Каждый проверяющий проходит опрос несколько раз, выбирая разные ответы.", "size": 13},
                ],
                ORANGE,
            ),
            (
                "Что проверяем",
                [
                    bullet("Тексты вопросов и ответов: опечатки, понятность"),
                    bullet("Логику: пройти каждую ветку. Testograf рекомендует тестировать все ветвления перед запуском"),
                    bullet("Обязательные вопросы отмечены там, где нужно"),
                    bullet("Вид на телефоне и компьютере — через предварительный просмотр"),
                    bullet("Страницы приветствия и завершения, экран дисквалификации"),
                ],
                VIOLET,
            ),
            (
                "Результат шага",
                [
                    {"text": "Команда утвердила опрос.", "size": 13, "bold": True, "after": 8},
                    {"text": "Если нашли ошибки — исправляем в конструкторе и проверяем еще раз.", "size": 13, "after": 10},
                    {"text": "Только после утверждения командой — шаг 5.", "size": 13, "bold": True, "color": GREEN},
                ],
                GREEN,
            ),
        ],
    )

    step_slide(
        d,
        5,
        "Заказчик проверяет опрос",
        "заказчик",
        [
            (
                "Что делаем",
                [
                    {"text": "После проверки и утверждения опроса командой передаем его на проверку заказчику.", "size": 13, "after": 8},
                    {"text": "Раньше утверждения командой опрос заказчику не отправляем.", "size": 13, "bold": True},
                ],
                ORANGE,
            ),
            (
                "Что проверяет заказчик",
                [
                    bullet("Опрос решает задачу из брифа"),
                    bullet("Все нужные вопросы на месте, формулировки верные"),
                    bullet("Логика ведет людей туда, куда нужно заказчику"),
                    {"text": "Заказчик тоже может проходить опрос много раз — ограничение повторов по-прежнему выключено.", "size": 12, "color": GRAY},
                ],
                VIOLET,
            ),
            (
                "Результат шага",
                [
                    {"text": "Заказчик утвердил опрос.", "size": 13, "bold": True, "after": 8},
                    {"text": "Если есть правки — вносим их и проверяем исправленный опрос.", "size": 13},
                ],
                GREEN,
            ),
        ],
        note=[{"text": "Одновременно с шагами 4 и 5 идет шаг 6 — подготовка коммуникации об опросе.", "size": 13}],
        note_title=None,
        note_tone="violet",
    )

    step_slide(
        d,
        6,
        "Готовим коммуникацию об опросе",
        "выполняется одновременно с шагами 4 и 5",
        [
            (
                "Что делаем",
                [
                    {"text": "Пока опрос проверяют команда и заказчик, готовим коммуникационное сообщение об этом опросе.", "size": 13, "after": 8},
                    {"text": "Так запуск не ждет, пока напишут текст.", "size": 13},
                ],
                ORANGE,
            ),
            (
                "Как связано с Testograf",
                [
                    {"text": "В сообщение пойдет ссылка на опрос. Способ доставки выбираем по брифу:", "size": 13, "after": 6},
                    bullet("прямая ссылка или QR-код в наших каналах"),
                    bullet("email-рассылка из Testograf (у каждого получателя своя ссылка)"),
                    bullet("SMS-рассылка из Testograf (PRO+, через SMS Профи)"),
                    bullet("виджет или всплывающее окно на сайте"),
                ],
                VIOLET,
            ),
            (
                "Результат шага",
                [
                    {"text": "Готов текст сообщения об опросе для каждого канала.", "size": 13, "after": 10},
                    {"text": "Опрос и сообщение вместе передаем на согласование — шаг 7.", "size": 13, "bold": True, "color": GREEN},
                ],
                GREEN,
            ),
        ],
    )

    step_slide(
        d,
        7,
        "Согласование опроса и коммуникации",
        f"руководитель департамента маркетинга {APPROVER}",
        [
            (
                "Что согласуем",
                [
                    {"text": "Два объекта вместе:", "size": 13, "bold": True, "after": 6},
                    bullet("сам опрос в Testograf", size=13),
                    bullet("коммуникационное сообщение об опросе", size=13),
                ],
                ORANGE,
            ),
            (
                "Кто согласует",
                [
                    {"text": APPROVER, "size": 14, "bold": True, "after": 4},
                    {"text": "руководитель департамента маркетинга", "size": 13, "color": GRAY, "after": 10},
                    {"text": "Опрос передаем после того, как его утвердили команда (шаг 4) и заказчик (шаг 5).", "size": 12.5},
                ],
                VIOLET_DARK,
            ),
            (
                "Результат шага",
                [
                    {"text": "Опрос и сообщение согласованы.", "size": 13, "bold": True, "after": 10},
                    {"text": "Только после этого — шаг 8.", "size": 13, "bold": True, "color": GREEN},
                ],
                GREEN,
            ),
        ],
        note=[{"text": "Без согласования руководителя департамента маркетинга шаги 8 и 9 не выполняются, и опрос клиентам не публикуется.", "size": 13}],
        note_title="Правило",
        note_tone="red",
    )

    step_slide(
        d,
        8,
        "Зачищаем ответы и включаем ограничение",
        "ответственный за опрос",
        [
            (
                "1. Удаляем тестовые ответы",
                [
                    {"text": "Открываем опрос в личном кабинете → «Общая информация» → удалить собранные ответы.", "size": 13, "after": 8},
                    {"text": "Проверяем, что в опросе 0 ответов.", "size": 13, "after": 8},
                    {"text": "Зачем: иначе ответы команды и заказчика попадут в итоговую статистику вместе с ответами клиентов.", "size": 12.5, "color": GRAY},
                ],
                ORANGE,
            ),
            (
                "2. Включаем ограничение",
                [
                    {"runs": [("Настройка → Основное → блок «Вход»", {"bold": True})], "size": 13, "after": 8},
                    {"text": "Включаем ограничение, чтобы опрос нельзя было пройти дважды.", "size": 13, "after": 8},
                    {"text": "Какой способ выбрать (устройство, IP, персональные ссылки, пароли, SMS) — зависит от опроса, см. часть 05.", "size": 12.5, "color": GRAY},
                ],
                VIOLET,
            ),
            (
                "Порядок строго такой",
                [
                    {"text": "Сначала удалить ответы, потом включить ограничение.", "size": 13, "bold": True, "after": 8},
                    {"text": "После этого сами опрос больше не проходим — каждый новый ответ будет считаться ответом клиента.", "size": 13},
                ],
                GREEN,
            ),
        ],
    )

    step_slide(
        d,
        9,
        "Запуск",
        "ответственный за опрос",
        [
            (
                "Перед публикацией",
                [
                    bullet("Задан период проведения опроса: Настройка → Основное", size=13),
                    bullet("Включен переключатель «Активный опрос»: карточка опроса → «Общая информация»", size=13),
                    bullet("В опросе 0 ответов, ограничение повторов включено (шаг 8)", size=13),
                ],
                ORANGE,
            ),
            (
                "Запускаем",
                [
                    {"text": "Публикуем согласованное сообщение со ссылкой на опрос в выбранных каналах или запускаем рассылку из Testograf.", "size": 13, "after": 8},
                    {"text": "В сообщении — ссылка именно на этот опрос.", "size": 13, "bold": True},
                ],
                VIOLET,
            ),
            (
                "После запуска",
                [
                    {"text": "Следим за опросом в «Общей информации»: число ответов и переходов, доля завершивших, среднее время заполнения, устройства и браузеры.", "size": 13, "after": 8},
                    {"text": "По email-рассылке есть своя статистика, неответившим можно отправить напоминание.", "size": 12.5, "color": GRAY},
                ],
                GREEN,
            ),
        ],
    )

    # ================================================================ 03
    section_slide(d, 3, "Типы вопросов", "Какие виды вопросов и ответов есть в Testograf, что каждый дает на выходе и для какого опроса подходит")

    s = d.new()
    d.header(s, "Как выбрать тип вопроса", "Тип вопроса определяет, как человек может ответить и что вы увидите в отчете. Начинайте с вопроса: что нужно получить на выходе?")
    rows = [
        ("Долю людей по вариантам (факт, сегмент)", "Выбор одного ответа, выпадающий список"),
        ("Сколько человек отметили каждую причину", "Выбор нескольких ответов"),
        ("Индекс лояльности к компании", "NPS"),
        ("Удовлетворенность параметрами с учетом их важности", "CSI"),
        ("Оценку по шкале", "Шкала Лайкерта, звезды, смайлы, семантический дифференциал"),
        ("Оценки многих пунктов одной таблицей", "Матрица"),
        ("Объяснение своими словами", "Свободный ответ (короткий или длинный текст)"),
        ("Порядок важности вариантов", "Ранжирование, парное сравнение, MaxDiff"),
        ("Проверку макета, текста или интерфейса", "UX-типы: тест первого клика, 5-секундный тест, A/B тест и др."),
    ]
    text_table(s, M, Inches(1.8), [Inches(6.0), CONTENT_W - Inches(6.0)], ["Что нужно получить", "Какой тип вопроса взять"], rows, Inches(0.53), size=12.5)

    s = d.new()
    d.header(s, "Выбор одного и выбор нескольких ответов", "Самые частые типы. Внешне похожи, но дают разный результат и по-разному работают в логике.")
    for i, (title, q, opts, square, fields) in enumerate(
        [
            (
                "Выбор одного ответа",
                "Пользовались ли вы мобильным приложением?",
                [("Да", True), ("Нет", False), ("Не знаю о приложении", False)],
                False,
                [
                    label_line("Что это:", "можно отметить только один вариант. Можно добавить поле «Свой ответ»."),
                    label_line("На выходе:", "процент людей по каждому варианту. Сумма — 100%."),
                    label_line("В логике:", "условия «выбрал / не выбрал / всегда»."),
                    label_line("Когда:", "варианты исключают друг друга: да/нет, ЖК, возраст, канал."),
                    label_line("Опросы:", "все — для фактов, сегментов и отбора аудитории."),
                ],
            ),
            (
                "Выбор нескольких ответов",
                "Что стоит улучшить в первую очередь?",
                [("Уборка территории", True), ("Работа охраны", False), ("Мобильное приложение", True)],
                True,
                [
                    label_line("Что это:", "можно отметить несколько вариантов. Можно добавить «Свой ответ»."),
                    label_line("На выходе:", "сколько человек отметили каждый вариант. Сумма может быть больше 100%."),
                    label_line("В логике:", "«выбрал / не выбрал / всегда»; сочетания И / ИЛИ — в продвинутой логике (PRO+)."),
                    label_line("Когда:", "причины, проблемы, что важно — человек может назвать несколько."),
                    label_line("Опросы:", "NPS, CSI, CSAT, после обращения — уточнить причины оценки."),
                ],
            ),
        ]
    ):
        x = M + i * (CONTENT_W / 2 + Inches(0.1))
        w = CONTENT_W / 2 - Inches(0.1)
        text(s, x, Inches(1.75), w, Inches(0.4), [{"text": title, "size": 17, "bold": True, "color": ORANGE}])
        mock_frame(s, x, Inches(2.2), w, Inches(1.55), q)
        for j, (o, checked) in enumerate(opts):
            radio(s, x + Inches(0.3), Inches(2.65) + j * Inches(0.33), o, checked, square)
        text(s, x, Inches(3.95), w, Inches(2.95), fields)

    s = d.new()
    d.header(s, "Списки и изображения", "Те же «выбор одного» и «выбор нескольких», но в другом виде — для длинных списков или визуальных вариантов.")
    items = [
        ("Выпадающий список", "Один вариант из раскрывающегося списка.", "Длинные справочники: ЖК, город, регион.", "Все опросы"),
        ("Выпадающий множественный", "Несколько вариантов из раскрывающегося списка.", "Длинный список, где можно выбрать несколько.", "По брифу"),
        ("Вложенные списки", "Выбор по уровням: вариант следующего списка зависит от предыдущего.", "Иерархия: регион → город → объект.", "По брифу"),
        ("Одно изображение", "Вместо текста варианты-картинки, выбирается одна.", "Сравнение визуальных вариантов.", "Голосования"),
        ("Несколько изображений", "Варианты-картинки, можно выбрать несколько.", "Отбор понравившихся вариантов.", "Голосования"),
    ]
    text_table(
        s,
        M,
        Inches(1.8),
        [Inches(2.9), Inches(3.9), Inches(3.6), CONTENT_W - Inches(10.4)],
        ["Тип", "Что это", "Когда использовать", "Опросы"],
        items,
        Inches(0.8),
        size=12.5,
    )
    callout(s, M, Inches(6.25), CONTENT_W, Inches(0.6), None, [{"text": "На выходе — то же, что у выбора одного / нескольких: распределение по вариантам в сводном отчете и выгрузках.", "size": 12.5}], tone="violet")

    s = d.new()
    d.header(s, "NPS — индекс лояльности", "Отдельный тип вопроса. Testograf сам считает индекс в отчете.")
    mock_frame(s, M, Inches(1.75), Inches(6.2), Inches(2.05), "Насколько вероятно, что вы порекомендуете нас друзьям?")
    nps_scale(s, M + Inches(0.25), Inches(2.4))
    text(s, M + Inches(0.25), Inches(3.0), Inches(2.3), Inches(0.6), [{"text": "0–6 критики", "size": 12, "bold": True, "color": RED}])
    text(s, M + Inches(3.05), Inches(3.0), Inches(1.6), Inches(0.6), [{"text": "7–8 нейтралы", "size": 12, "bold": True, "color": YELLOW}])
    text(s, M + Inches(4.5), Inches(3.0), Inches(1.6), Inches(0.6), [{"text": "9–10 промоутеры", "size": 12, "bold": True, "color": GREEN}])
    callout(
        s,
        M,
        Inches(4.05),
        Inches(6.2),
        Inches(1.1),
        "Формула",
        [{"text": "NPS = доля промоутеров (9–10) − доля критиков (0–6)", "size": 15, "bold": True}],
        tone="green",
    )
    text(
        s,
        M,
        Inches(5.35),
        Inches(6.2),
        Inches(1.5),
        [{"text": "В отчете Testograf индекс рассчитывается автоматически, динамику можно отслеживать от волны к волне.", "size": 12.5, "color": GRAY}],
    )
    card(
        s,
        M + Inches(6.5),
        Inches(1.75),
        CONTENT_W - Inches(6.5),
        Inches(5.1),
        "Когда и где использовать",
        [
            label_line("Что измеряет:", "отношение клиента к компании в целом — готовность рекомендовать."),
            label_line("На выходе:", "индекс NPS и распределение по группам: критики, нейтралы, промоутеры."),
            label_line("Опросы:", "основной тип для NPS. В других опросах — только если бриф просит индекс лояльности."),
            label_line("Чем дополнить:", "свободным ответом «Почему такая оценка?» и выбором нескольких причин."),
            label_line("Отличие от CSI:", "NPS — про компанию целиком, CSI — про отдельные параметры."),
        ],
        accent=ORANGE,
    )

    s = d.new()
    d.header(s, "CSI — удовлетворенность по параметрам", "Отдельный тип вопроса. По каждому параметру человек ставит две оценки: насколько это важно и насколько он доволен.")
    rows = [
        ("Удобство личного кабинета", "8", "7", "0,8 × 0,7 = 56%"),
        ("Качество оказания услуг", "5", "10", "50%"),
        ("Стоимость услуг", "10", "7", "70%"),
        ("Вежливость колл-центра", "6", "8", "48%"),
        ("Общее впечатление", "10", "9", "90%"),
    ]
    end_y = table(s, M, Inches(1.8), [Inches(3.0), Inches(1.15), Inches(1.15), Inches(1.6)], ["Параметр", "Важность", "Оценка", "CSI параметра"], rows, row_h=Inches(0.44), size=12)
    text(
        s,
        M,
        end_y + Inches(0.12),
        Inches(6.9),
        Inches(1.4),
        [
            {"runs": [("Итог: ", {"bold": True}), ("CSI = (56 + 50 + 70 + 48 + 90) / 5 ≈ 63%. Условный пример из справки Testograf.", {})], "size": 12.5, "after": 6},
            {"runs": [("Смотрите на разрыв: ", {"bold": True}), ("стоимость максимально важна, но оценена на 7 — первый кандидат на улучшение.", {})], "size": 12.5},
        ],
    )
    card(
        s,
        M + Inches(7.2),
        Inches(1.8),
        CONTENT_W - Inches(7.2),
        Inches(5.05),
        "Как считается и что важно",
        [
            label_line("Шаг 1.", "Средние оценки важности и удовлетворенности по каждому параметру."),
            label_line("Шаг 2.", "Переводим в доли от максимума и перемножаем: важность × оценка × 100%."),
            label_line("Шаг 3.", "Среднее по всем параметрам — итоговый CSI."),
            label_line("Рекомендации Testograf:", "для клиентов (b2c) — 3, максимум 5 параметров; шкала 5 или 10 баллов; четная шкала убирает среднюю оценку."),
            label_line("Опросы:", "основной тип для CSI."),
        ],
        accent=ORANGE,
    )

    s = d.new()
    d.header(s, "Шкалы оценки", "Когда нужно получить оценку числом, но не индекс NPS или CSI.")
    cw = (CONTENT_W - Inches(0.25) * 3) / 4
    scales = [
        (
            "Шкала Лайкерта",
            "Числовая шкала от 0 или 1, от 3 до 10 пунктов: числа, числа в рамках или слайдер. Или своя шкала с подписями. Можно добавить «Не могу оценить».",
            "Согласие, удовлетворенность, частота.",
        ),
        ("Звездный рейтинг", "Оценка звездами.", "Короткая оценка, удобная на телефоне."),
        ("Смайл рейтинг", "Оценка смайлами.", "Быстрая эмоциональная оценка."),
        (
            "Дифференциал",
            "Семантический дифференциал — шкала между двумя противоположностями: «ненадежный — надежный». Часто 5 или 7 делений.",
            "Восприятие бренда, сервиса, продукта.",
        ),
    ]
    for i, (t, what, when) in enumerate(scales):
        x = M + i * (cw + Inches(0.25))
        card(
            s,
            x,
            Inches(1.8),
            cw,
            Inches(3.85),
            t,
            [label_line("Что это:", what, size=12), label_line("Когда:", when, size=12), label_line("На выходе:", "распределение ответов и средняя оценка.", size=12)],
            accent=VIOLET,
            title_size=14,
        )
    stars(s, M + cw + Inches(0.25) + Inches(0.2), Inches(5.0))
    callout(
        s,
        M,
        Inches(5.85),
        CONTENT_W,
        Inches(1.0),
        "Для каких опросов",
        [
            {"text": "CSAT и опросы после обращения — звезды, смайлы или Лайкерт для оценки взаимодействия. В NPS и CSI — дополнительные вопросы. Распределительная шкала тоже есть в сервисе, но в открытой справке не описана — проверяйте ее в конструкторе.", "size": 12.5}
        ],
        tone="violet",
    )

    s = d.new()
    d.header(s, "Матрицы", "Вопрос-таблица: в строках — что оцениваем, в столбцах — варианты ответа. Вместо пяти одинаковых вопросов — одна таблица.")
    mx, my = M, Inches(1.8)
    mock_frame(s, mx, my, Inches(6.2), Inches(2.6), "Оцените, насколько вы довольны:")
    cols_lbl = ["1", "2", "3", "4", "5"]
    for j, c in enumerate(cols_lbl):
        text(s, mx + Inches(2.8) + j * Inches(0.65), my + Inches(0.6), Inches(0.5), Inches(0.3), [{"text": c, "size": 12, "bold": True, "align": PP_ALIGN.CENTER, "color": GRAY}])
    for r, (row_name, pick) in enumerate([("Уборка", 3), ("Охрана", 4), ("Приложение", 1)]):
        yy = my + Inches(1.0) + r * Inches(0.48)
        text(s, mx + Inches(0.25), yy - Inches(0.02), Inches(2.4), Inches(0.35), [{"text": row_name, "size": 12}])
        for j in range(5):
            cx = mx + Inches(2.95) + j * Inches(0.65)
            rect(s, cx, yy + Inches(0.03), Inches(0.2), Inches(0.2), WHITE, MSO_SHAPE.OVAL, line=GRAY_MID)
            if j == pick:
                rect(s, cx + Inches(0.05), yy + Inches(0.08), Inches(0.1), Inches(0.1), VIOLET_DARK, MSO_SHAPE.OVAL)
    text(
        s,
        M,
        Inches(4.6),
        Inches(6.2),
        Inches(2.3),
        [
            label_line("В логике:", "условие задается по конкретной ячейке: «в строке 1 выбран столбец 2»."),
            label_line("Подстановка:", "строки матрицы можно заполнить вариантами, которые человек выбрал раньше (конвейерная подстановка)."),
            label_line("Опросы:", "блоки оценок факторов в CSI, дополнительные блоки в NPS и CSAT."),
        ],
    )
    card(
        s,
        M + Inches(6.5),
        Inches(1.8),
        CONTENT_W - Inches(6.5),
        Inches(5.05),
        "Пять видов матриц в Testograf",
        [
            label_line("Один ответ —", "в каждой строке выбирается один столбец. Классика для оценок."),
            label_line("Несколько ответов —", "в строке можно отметить несколько столбцов."),
            label_line("Звездный рейтинг —", "оценка звездами по каждой строке."),
            label_line("Текстовый ответ —", "в ячейках пишут текст."),
            label_line("Выпадающий список —", "в ячейках выбирают из списка."),
            {"text": "Риск больших матриц: человек может поставить один и тот же ответ во всех строках, не читая. Проверяйте это при анализе.", "size": 12, "color": RED},
        ],
        accent=ORANGE,
    )

    s = d.new()
    d.header(s, "Свободные ответы и контактные данные", "Когда нужны слова человека, число, файл или его контакты.")
    rows = [
        ("Короткий текст", "Строка для короткого ответа.", "Уточнение: «Укажите номер квартиры»."),
        ("Длинный текст", "Поле для развернутого ответа.", "«Почему вы поставили такую оценку?»"),
        ("Группа ответов", "Несколько текстовых полей в одном вопросе.", "Несколько уточнений сразу."),
        ("Число", "Ввод числа.", "Количество, сумма, срок."),
        ("Загрузка файлов", "Респондент прикрепляет файл.", "Фото проблемы, документ."),
        ("Контактные данные", "Имя, email, телефон, место жительства, дата, время.", "Опрос не анонимный, нужна обратная связь."),
        ("Галочка-согласие", "Согласие на обработку персональных данных.", "Обязательно, если собираете контакты."),
    ]
    text_table(s, M, Inches(1.75), [Inches(2.7), Inches(5.0), CONTENT_W - Inches(7.7)], ["Тип", "Что это", "Пример использования"], rows, Inches(0.5), size=12.5)
    callout(
        s,
        M,
        Inches(5.95),
        CONTENT_W,
        Inches(0.9),
        None,
        [
            {"runs": [("На выходе: ", {"bold": True}), ("тексты и значения в отдельных ответах и выгрузках. ", {}), ("В логике: ", {"bold": True}), ("по тексту ответа ветвить нельзя — только переход «всегда». ", {}), ("Опросы: ", {"bold": True}), ("комментарии — во всех; контакты — когда опрос не анонимный.", {})], "size": 12.5}
        ],
        tone="violet",
    )

    s = d.new()
    d.header(s, "Расстановка приоритетов", "Когда нужно понять, что важнее, а не просто получить оценки.")
    cw = (CONTENT_W - Inches(0.25) * 2) / 3
    pr = [
        (
            "Ранжирование",
            [
                label_line("Что это:", "расставить варианты по местам.", size=12),
                label_line("На выходе:", "взвешенный средний балл. Пример при 5 вариантах: 1 место = 5 баллов … 5 место = 1. Выше балл — важнее.", size=12),
                label_line("Опросы:", "приоритеты улучшений в любом опросе.", size=12),
            ],
        ),
        (
            "Парное сравнение",
            [
                label_line("Что это:", "человек много раз выбирает лучший из двух вариантов. От 3 до 12 вариантов, оптимально 4–8.", size=12),
                label_line("На выходе:", "сколько раз каждый вариант победил, диаграмма.", size=12),
                label_line("Опросы:", "голосования, выбор концепции.", size=12),
            ],
        ),
        (
            "MaxDiff (лучший — худший)",
            [
                label_line("Что это:", "в каждом блоке из 3–5 пунктов выбирают самый важный и самый неважный. От 4 до 20 пунктов.", size=12),
                label_line("На выходе:", "индекс важности каждого пункта.", size=12),
                label_line("Опросы:", "продуктовые и маркетинговые исследования по брифу.", size=12),
            ],
        ),
    ]
    for i, (t, body) in enumerate(pr):
        card(s, M + i * (cw + Inches(0.25)), Inches(1.75), cw, Inches(3.3), t, body, accent=VIOLET, title_size=14)
    pr2 = [
        (
            "Conjoint-анализ",
            "Карточки с сочетаниями характеристик (например, цена, бренд, срок). На каждом экране выбирают одну карточку. На выходе — какие характеристики сильнее влияют на выбор.",
        ),
        (
            "Метод Кано",
            "По каждой функции два вопроса: как отнесетесь, если она есть, и если ее нет. На выходе — тип функции: обязательная, важная, привлекательная, безразличная.",
        ),
    ]
    cw2 = (CONTENT_W - Inches(0.25)) / 2
    for i, (t, b) in enumerate(pr2):
        card(s, M + i * (cw2 + Inches(0.25)), Inches(5.2), cw2, Inches(1.65), t, [{"text": b, "size": 12}], accent=GRAY_MID, title_size=14)

    s = d.new()
    d.header(s, "UX-типы: проверка макетов и текстов", "Testograf относит их к исследованиям интерфейсов и продукта. В обычный NPS или CSI их не добавляем — только если бриф ставит такую задачу.")
    ux = [
        ("Тест первого клика", "Человек кликает по изображению интерфейса. До 12 целевых областей, до 10 кликов, можно таймер.", "Тепловая карта, распределение кликов, время до клика."),
        ("5-секундный тест", "Экран показывают от 5 до 60 секунд, затем спрашивают, что запомнилось.", "Что человек понял и запомнил с первого взгляда."),
        ("A/B тест", "Каждый видит только один вариант (картинка, текст, видео) в заданной доле, например 50/50.", "Сравнение реакции групп на разные варианты."),
        ("Тест прототипа Figma", "Прототип из Figma прямо в опросе. Можно записывать экран, голос, камеру.", "Выполнил ли задачу, маршрут по экранам, клики."),
        ("Карточная сортировка", "Человек раскладывает карточки по категориям — заданным или своим.", "Как люди группируют информацию."),
        ("Восприятие текста", "Человек кликает по словам и фразам: непонятным, понравившимся.", "Какие места текста упростить или усилить."),
    ]
    cw = (CONTENT_W - Inches(0.25) * 2) / 3
    for i, (t, what, out) in enumerate(ux):
        col, row = i % 3, i // 3
        card(
            s,
            M + col * (cw + Inches(0.25)),
            Inches(1.85) + row * Inches(2.55),
            cw,
            Inches(2.4),
            t,
            [label_line("Что это:", what, size=12), label_line("На выходе:", out, size=12)],
            accent=GRAY_MID,
            title_size=14,
        )

    s = d.new()
    d.header(s, "Какой тип вопроса в каком опросе", "Основной — на нем строится метрика опроса. Дополнительно — для уточнений и причин. По брифу — только если этого требует задача.")

    def colorize(cell, color, bold):
        if cell == "Основной":
            return GREEN, True
        if cell == "Дополнительно":
            return VIOLET_DARK, False
        if cell == "По брифу":
            return ORANGE, False
        return GRAY_MID, False

    rows = [
        ("Выбор одного ответа", "Дополнительно", "Дополнительно", "Дополнительно", "Основной", "Дополнительно"),
        ("Выбор нескольких ответов", "Дополнительно", "Дополнительно", "Дополнительно", "По брифу", "Дополнительно"),
        ("Выпадающие и вложенные списки", "Дополнительно", "Дополнительно", "Дополнительно", "По брифу", "Дополнительно"),
        ("Выбор изображений", "—", "—", "—", "Основной", "—"),
        ("NPS", "Основной", "По брифу", "По брифу", "—", "По брифу"),
        ("CSI", "По брифу", "Основной", "—", "—", "По брифу"),
        ("Лайкерт, звезды, смайлы", "Дополнительно", "Дополнительно", "Основной", "—", "Основной"),
        ("Семантический дифференциал", "По брифу", "По брифу", "По брифу", "—", "—"),
        ("Матрицы", "Дополнительно", "Основной", "Дополнительно", "—", "Дополнительно"),
        ("Свободный ответ", "Дополнительно", "Дополнительно", "Дополнительно", "По брифу", "Дополнительно"),
        ("Контактные данные + согласие", "По брифу", "По брифу", "По брифу", "По брифу", "По брифу"),
        ("Ранжирование", "По брифу", "По брифу", "—", "По брифу", "—"),
        ("Парное сравнение", "—", "—", "—", "По брифу", "—"),
        ("MaxDiff, Conjoint, Кано, UX-типы", "—", "—", "—", "—", "—"),
    ]
    first_w = Inches(3.55)
    other = (CONTENT_W - first_w) / 5
    table(s, M, Inches(1.75), [first_w] + [other] * 5, ["Тип вопроса", "NPS", "CSI", "CSAT", "Голосование", "После обращения"], rows, row_h=Inches(0.345), size=11, colorize=colorize)

    # ================================================================ 04
    section_slide(d, 4, "Логика", "Как показывать людям разные вопросы в зависимости от их ответов")

    s = d.new()
    d.header(s, "Обычная логика переходов", "Логика — правило, которое решает, какой вопрос человек увидит следующим. Без логики все проходят анкету одинаково, вопрос за вопросом.")
    # diagram
    flow_box(s, M, Inches(2.0), Inches(3.2), Inches(1.0), "Вопрос", "«Вы живете в этом ЖК?»", accent=ORANGE)
    arrow_right(s, M + Inches(3.3), Inches(2.15))
    arrow_right(s, M + Inches(3.3), Inches(2.6))
    flow_box(s, M + Inches(3.8), Inches(1.75), Inches(2.6), Inches(0.75), "Да →", "вопросы о ЖК и NPS", accent=GREEN)
    flow_box(s, M + Inches(3.8), Inches(2.6), Inches(2.6), Inches(0.75), "Нет →", "дисквалификация", accent=RED)
    text(
        s,
        M,
        Inches(3.65),
        Inches(6.4),
        Inches(3.2),
        [
            {"runs": [("Где: ", {"bold": True, "color": VIOLET_DARK}), ("вкладка «Логика» в конструкторе.", {})], "size": 13, "after": 8},
            {"text": "Куда можно направить человека:", "size": 13, "bold": True, "after": 4},
            bullet("на другой вопрос или страницу"),
            bullet("на стандартное завершение или индивидуальное завершение со своим текстом"),
            bullet("на дисквалификацию — экран «опрос не для вас»"),
            {"text": "К одному вопросу можно добавить несколько условий. Если условие удалить, работает прямая логика.", "size": 12, "color": GRAY},
        ],
    )
    card(
        s,
        M + Inches(6.7),
        Inches(1.75),
        CONTENT_W - Inches(6.7),
        Inches(5.1),
        "Какие условия доступны",
        [
            label_line("Выбор одного:", "выбрал вариант / не выбрал / всегда. Можно отметить несколько вариантов в одном условии."),
            label_line("Выбор нескольких:", "выбрал указанные / не выбрал / всегда. Связки И / ИЛИ — только в продвинутой логике."),
            label_line("Матрица:", "условие по ячейке матрицы."),
            label_line("Свободный ответ, ранжирование:", "только переход «всегда» — по содержанию ответа ветвить нельзя."),
            {"text": "Testograf рекомендует перед запуском протестировать все ветвления — это наш шаг 4.", "size": 12.5, "bold": True, "color": ORANGE},
        ],
        accent=VIOLET,
    )

    s = d.new()
    d.header(s, "Продвинутая логика (PRO+)", "Нужна, когда решение зависит сразу от нескольких ответов или от параметра ссылки. Вкладка «Продвинутая логика» → «Добавить логику».")
    flow_box(s, M, Inches(1.85), Inches(3.6), Inches(0.9), "Условие 1", "[Оценка услуги] = «5»", accent=VIOLET)
    text(s, M + Inches(3.7), Inches(2.05), Inches(0.9), Inches(0.5), [{"text": "ИЛИ", "size": 16, "bold": True, "color": ORANGE, "align": PP_ALIGN.CENTER}])
    flow_box(s, M + Inches(4.6), Inches(1.85), Inches(3.6), Inches(0.9), "Условие 2", "[Оценка сервиса] = «Отлично»", accent=VIOLET)
    arrow_right(s, M + Inches(8.3), Inches(2.15))
    flow_box(s, M + Inches(8.75), Inches(1.85), CONTENT_W - Inches(8.75), Inches(0.9), "Действие", "индивидуальное завершение с благодарностью и ссылкой на отзыв", accent=GREEN)
    text(s, M, Inches(2.85), CONTENT_W, Inches(0.35), [{"text": "Пример из справки Testograf", "size": 11, "color": GRAY_MID}])
    card(
        s,
        M,
        Inches(3.35),
        Inches(6.1),
        Inches(3.5),
        "Условие — когда срабатывает",
        [
            bullet("ответы на любые вопросы анкеты", size=12),
            bullet("параметры ссылки (город, услуга, точка продаж): «содержит / не содержит»", size=12),
            bullet("между условиями — И / ИЛИ", size=12),
            {"text": "Пример из справки: регион = «Другое» И «Студент» → дисквалификация.", "size": 12, "color": GRAY},
        ],
        accent=VIOLET,
    )
    card(
        s,
        M + Inches(6.3),
        Inches(3.35),
        CONTENT_W - Inches(6.3),
        Inches(3.5),
        "Действие — что происходит",
        [
            bullet("переход к вопросу или странице", size=12),
            bullet("показать или пропустить список вопросов и страниц", size=12),
            bullet("завершить опрос сразу или после определенного вопроса: стандартное завершение, индивидуальное или дисквалификация", size=12),
            {"text": "Один и тот же вопрос или параметр нельзя использовать дважды в одном наборе условий.", "size": 12, "color": GRAY},
        ],
        accent=GREEN,
    )

    s = d.new()
    d.header(s, "Подстановка ответов", "Два механизма, которые делают анкету персональной: текст и варианты подстраиваются под ответы человека.")
    card(
        s,
        M,
        Inches(1.8),
        Inches(6.1),
        Inches(5.05),
        "Подстановка в тексты (piping)",
        [
            label_line("Что это:", "ответ или параметр ссылки вставляется в текст вопроса, вариантов, приветствия или завершения."),
            label_line("Откуда берется:", "ответы на предыдущие вопросы (например, короткий текст) и параметры ссылки. Полный список — в Помощи Testograf."),
            label_line("Пример:", "«Вы выбрали ЖК [ответ]. Оцените управляющую компанию»."),
            label_line("Ограничение:", "в заголовки строк и столбцов матриц подстановка не работает."),
        ],
        accent=VIOLET,
    )
    card(
        s,
        M + Inches(6.3),
        Inches(1.8),
        CONTENT_W - Inches(6.3),
        Inches(5.05),
        "Конвейерная подстановка вариантов",
        [
            label_line("Что это:", "варианты, выбранные в одном вопросе, становятся вариантами или строками следующего."),
            label_line("Пример из справки:", "человек отметил «кредит», «депозит», «карта» — именно эти пункты появляются строками матрицы для оценки."),
            label_line("Куда можно:", "все типы матриц (в строки), ранжирование, распределительная шкала, группа ответов, CSI и др. Полный список — в Помощи."),
            label_line("Особенность:", "если подставить нечего и своих вариантов нет — вопрос скрывается автоматически."),
        ],
        accent=ORANGE,
    )

    # ================================================================ 05
    section_slide(d, 5, "Настройки", "Что включать и выключать, что это меняет и в каких опросах нужно")

    s = d.new()
    d.header(s, "Где находятся настройки", "Большая часть настроек — во вкладке «Настройка» конструктора. Часть — прямо в каждом вопросе.")
    cw = (CONTENT_W - Inches(0.25) * 3) / 4
    where = [
        (
            "Настройка → Основное",
            [
                bullet("название, заголовок, адрес ссылки", size=12),
                bullet("период проведения", size=12),
                bullet("приветствие и завершение", size=12),
                bullet("блок «Вход»: пароли, ВКонтакте, ограничения по устройству и IP", size=12),
                bullet("незавершенные ответы (PRO+)", size=12),
                bullet("показ результатов, запрет копирования", size=12),
                bullet("«Дополнительно»: SMS-подтверждение", size=12),
            ],
            ORANGE,
        ),
        (
            "Настройка → Параметры",
            [
                bullet("таймер опроса", size=12),
                bullet("индикатор прохождения", size=12),
                bullet("ответ по клику", size=12),
                bullet("нумерация вопросов", size=12),
                bullet("звездочка у обязательных вопросов", size=12),
            ],
            VIOLET,
        ),
        (
            "Настройка → Тексты",
            [
                bullet("переименовать кнопки", size=12),
                bullet("тексты оповещений о правильности заполнения", size=12),
                bullet("перевод на другой язык или два языка сразу", size=12),
            ],
            VIOLET,
        ),
        (
            "В каждом вопросе",
            [
                bullet("обязательный / необязательный", size=12),
                bullet("подсказка: текст, картинка, видео, ссылка", size=12),
                bullet("таймер вопроса", size=12),
                bullet("случайный порядок вариантов", size=12),
                bullet("комментарий к ответу", size=12),
                {"text": "На странице — показ вопросов по одному и случайный порядок вопросов.", "size": 11.5, "color": GRAY},
            ],
            GREEN,
        ),
    ]
    for i, (t, body, acc) in enumerate(where):
        card(s, M + i * (cw + Inches(0.25)), Inches(1.8), cw, Inches(5.05), t, body, accent=acc, title_size=14)

    restrict_rows_1 = [
        (
            "С устройства (cookies)",
            "При первом заполнении в браузере сохраняется отметка. Можно задать, сколько раз и за какой период можно пройти опрос.",
            "Открытые опросы по общей ссылке: ЛК, push, мессенджеры.",
            "Легко обойти: очистить cookies, открыть в режиме инкогнито или с другого устройства.",
        ),
        (
            "По IP-адресу",
            "Повторный ответ с того же сетевого адреса блокируется.",
            "Открытые голосования, где нужно отсечь явные накрутки.",
            "Один IP бывает у многих людей — офис, семья, общий Wi‑Fi: отсекаются честные участники. VPN и мобильные сети снижают защиту.",
        ),
        (
            "Персональные ссылки",
            "У каждого своя ссылка, которая работает один раз. Параметры ссылки → «+ Значения» → «Мульти-загрузка» → «Скачать ссылки» + «Ограничить повторное заполнение по параметру». Или email-рассылка с «Запретить повторное заполнение».",
            "Список участников известен заранее: email, телефоны, ID клиентов.",
            "Нужен список заранее. Если ссылку переслали другому — ответит другой человек.",
        ),
    ]
    s = d.new()
    d.header(s, "Ограничение повторов: 3 основных способа", "Цель — один человек дает один ответ. Все способы в блоке «Вход»: Настройка → Основное.")
    text_table(
        s,
        M,
        Inches(1.75),
        [Inches(2.3), Inches(4.2), Inches(2.9), CONTENT_W - Inches(9.4)],
        ["Способ", "Как работает", "Когда включать", "Ограничения"],
        restrict_rows_1,
        [Inches(1.0), Inches(1.0), Inches(1.6)],
        size=11.5,
    )
    callout(
        s,
        M,
        Inches(5.95),
        CONTENT_W,
        Inches(0.9),
        "Наш регламент",
        [{"text": "На шагах 3–7 все ограничения выключены, чтобы можно было тестировать. На шаге 8 — сначала удаляем тестовые ответы, потом включаем подходящий способ.", "size": 12.5}],
    )

    s = d.new()
    d.header(s, "Ограничение доступа: пароли, ВКонтакте, SMS", "Строже, чем ограничения по устройству и IP, но добавляют человеку лишний шаг перед ответом.")
    restrict_rows_2 = [
        (
            "Общий пароль",
            "Опрос открывается только после ввода пароля, одного для всех.",
            "Закрытый круг участников.",
            "Если пароль утек — пройти может любой.",
        ),
        (
            "Индивидуальные пароли",
            "У каждого свой пароль. Позволяет заполнить опрос один раз, пароль подтягивается к ответу. Доступ по паролю → Индивидуальные пароли → Идентификация ответа → Добавить пароли.",
            "Нужно точно знать, кто ответил.",
            "Нужно заранее раздать пароли участникам.",
        ),
        (
            "Авторизация ВКонтакте",
            "Для прохождения опроса нужно войти через ВКонтакте.",
            "Аудитория пользуется ВКонтакте.",
            "Без аккаунта ВКонтакте опрос не пройти.",
        ),
        (
            "SMS-подтверждение (PRO+)",
            "После ответа человек вводит код из SMS или звонка. Без верного кода ответ не сохраняется. Можно запретить повторный номер. Настройка → Основное → Дополнительно.",
            "Голосования, конкурсы с призами, защита от накруток.",
            "Нужен кабинет SMS Профи, SMS платные. Testograf: в анонимных опросах и опросах лояльности создает барьер.",
        ),
    ]
    text_table(
        s,
        M,
        Inches(1.75),
        [Inches(2.3), Inches(4.6), Inches(2.6), CONTENT_W - Inches(9.5)],
        ["Способ", "Как работает", "Когда включать", "Ограничения"],
        restrict_rows_2,
        [Inches(0.75), Inches(1.2), Inches(0.75), Inches(1.4)],
        size=11.5,
    )

    s = d.new()
    d.header(s, "Сохранение незавершенных ответов (PRO+)", "Человек может закрыть опрос на середине и позже продолжить с того же места. А вы видите даже недозаполненные анкеты.")
    card(
        s,
        M,
        Inches(1.8),
        Inches(6.1),
        Inches(3.2),
        "Как продолжить заполнение",
        [
            label_line("Email- или SMS-рассылка Testograf:", "у каждого своя ссылка — можно продолжить с любого устройства."),
            label_line("Индивидуальные пароли:", "ввел пароль — продолжил с того же места на любом устройстве."),
            label_line("Общая ссылка:", "при первом заходе ссылка становится уникальной для браузера — продолжить можно только на том же устройстве."),
        ],
        accent=VIOLET,
    )
    card(
        s,
        M + Inches(6.3),
        Inches(1.8),
        CONTENT_W - Inches(6.3),
        Inches(3.2),
        "Что видно в личном кабинете",
        [
            bullet("Ответ появляется в результатах, как только пройден хотя бы один вопрос — с пометкой «Незавершен»"),
            bullet("Видно, на каком шаге человек остановился, — можно напомнить ему"),
            bullet("Когда человек завершает анкету, статус меняется автоматически"),
        ],
        accent=ORANGE,
    )
    cw = (CONTENT_W - Inches(0.25) * 2) / 3
    callout(s, M, Inches(5.2), cw, Inches(1.65), "Включать", [{"text": "Длинные опросы и анкеты, которые заполняют в несколько подходов: длинные NPS и CSI.", "size": 12.5}], tone="green")
    callout(s, M + cw + Inches(0.25), Inches(5.2), cw, Inches(1.65), "Не нужно", [{"text": "Короткие опросы на один экран: CSAT, голосования.", "size": 12.5}], tone="violet")
    callout(s, M + 2 * (cw + Inches(0.25)), Inches(5.2), cw, Inches(1.65), "Помните", [{"text": "Незавершенную уникальную ссылку нельзя пересылать — по ней виден черновик. В индексы берем только завершенные анкеты.", "size": 12.5}], tone="red")

    s = d.new()
    d.header(s, "Случайный порядок и «блуждающие ответы»", "Случайный порядок — настройка, чтобы порядок вариантов не влиял на ответы. Отдельной настройки «блуждающие ответы» в Testograf нет.")
    card(
        s,
        M,
        Inches(1.8),
        Inches(6.1),
        Inches(5.05),
        "Случайный порядок (рандомизация)",
        [
            label_line("Где:", "в вопросе — случайный порядок вариантов ответа; на странице — случайный порядок вопросов."),
            label_line("Зачем:", "по Testograf — чтобы последовательность показа не влияла на ответы. Например, первый вариант списка не получал лишних отметок."),
            label_line("Что получаем:", "каждый респондент видит варианты в своем порядке; в отчете ответы сводятся как обычно."),
            label_line("Включать:", "когда варианты равноценны: список причин, вариантов дизайна, услуг."),
            label_line("Не включать:", "когда порядок имеет смысл: шкалы «от плохого к хорошему», вопросы, которые идут логической цепочкой."),
        ],
        accent=VIOLET,
    )
    card(
        s,
        M + Inches(6.3),
        Inches(1.8),
        CONTENT_W - Inches(6.3),
        Inches(5.05),
        "Что обычно называют «блуждающими»",
        [
            label_line("1. Варианты «гуляют» по местам", "у разных людей — это случайный порядок, слева. Включается и выключается в вопросе."),
            label_line("2. Человек отвечает наугад,", "не читая, чтобы быстрее закончить. Настройкой это не выключить. Снижаем риск: короткая анкета, без таймера, проверка при анализе."),
            label_line("3. Человек бросил анкету", "на середине — это незавершенные ответы, предыдущий слайд. В индексы не берем."),
        ],
        accent=ORANGE,
    )

    s = d.new()
    d.header(s, "Остальные настройки вкладки «Основное»", "Что дает каждая настройка и когда она нужна.")
    rows = [
        ("Название, заголовок, адрес", "Название в кабинете, заголовок для респондента, свой адрес ссылки.", "Всегда: понятное название и адрес."),
        ("Период проведения", "Опрос принимает ответы только в заданные даты.", "Всегда задаем при запуске (шаг 9)."),
        ("Страница приветствия", "Текст, картинка, ссылка до начала опроса.", "Когда нужно объяснить цель и время заполнения."),
        ("Страница завершения", "Текст, картинка, ссылка после ответа. Можно выдать промокоды — до 100 000 кодов.", "Всегда — благодарность. Промокоды — если обещан бонус."),
        ("Перенаправление", "После ответа человек попадает на другой опрос или сайт.", "Когда после опроса нужно вести дальше."),
        ("Показ результатов", "После ответа человек видит итоги опроса. Что показать — настраивается.", "Голосования — если решено показывать итоги."),
        ("Запрет копирования", "Текст опроса нельзя скопировать.", "Когда содержание опроса нельзя распространять."),
    ]
    text_table(s, M, Inches(1.65), [Inches(2.9), Inches(5.3), CONTENT_W - Inches(8.2)], ["Настройка", "Что дает", "Когда нужна"], rows, Inches(0.68), size=12)

    s = d.new()
    d.header(s, "Как выглядит опрос для респондента", "Настройки вида и ритма заполнения.")
    rows = [
        ("Показ вопросов", "Полотном (все вопросы на одной странице), по страницам или по одному вопросу.", "Короткий опрос — полотном. Длинный — по страницам."),
        ("Индикатор прохождения", "Полоса прогресса: сколько осталось.", "Длинные опросы."),
        ("Таймер опроса и вопроса", "Ограничивает время на опрос или на один вопрос.", "Тесты и задания на время. В оценочных опросах обычно не нужен: торопит человека."),
        ("Обязательный вопрос", "Без ответа нельзя перейти дальше.", "Ключевые вопросы: NPS, оценка, ЖК. Комментарии — необязательные."),
        ("Подсказка к вопросу", "Пояснение: текст, картинка, видео, ссылка.", "Когда вопрос может быть непонятен."),
        ("Нумерация и звездочка", "Номера вопросов и отметка обязательных.", "Для ясности в длинных анкетах."),
        ("Тексты кнопок", "Свои названия кнопок и оповещений, перевод на другой язык.", "Нестандартный сценарий или другой язык."),
    ]
    text_table(s, M, Inches(1.65), [Inches(2.9), Inches(5.0), CONTENT_W - Inches(7.9)], ["Настройка", "Что дает", "Когда нужна"], rows, Inches(0.68), size=12)

    s = d.new()
    d.header(s, "Настройки по типам опросов", "Как настроить опрос на боевом запуске (после шага 8). Вкл — включить, Выкл — выключить, По брифу — решаем по задаче.")

    def colorize2(cell, color, bold):
        if cell.startswith("Вкл"):
            return GREEN, True
        if cell.startswith("Выкл"):
            return RED, True
        if cell.startswith("По брифу"):
            return ORANGE, False
        return GRAY_MID, False

    rows = [
        ("Ограничение с устройства", "Вкл", "Вкл", "Вкл", "Вкл", "Вкл"),
        ("Ограничение по IP", "Выкл", "Выкл", "По брифу", "По брифу", "Выкл"),
        ("Персональные ссылки", "По брифу", "По брифу", "По брифу", "По брифу", "По брифу"),
        ("SMS-подтверждение", "Выкл", "Выкл", "Выкл", "По брифу", "Выкл"),
        ("Незавершенные ответы", "По брифу", "По брифу", "Выкл", "Выкл", "Выкл"),
        ("Случайный порядок вариантов", "По брифу", "По брифу", "По брифу", "По брифу", "По брифу"),
        ("Таймер опроса", "Выкл", "Выкл", "Выкл", "Выкл", "Выкл"),
        ("Показ результатов респонденту", "Выкл", "Выкл", "Выкл", "По брифу", "Выкл"),
        ("Обязательный ключевой вопрос", "Вкл", "Вкл", "Вкл", "Вкл", "Вкл"),
    ]
    first_w = Inches(3.8)
    other = (CONTENT_W - first_w) / 5
    end_y = table(s, M, Inches(1.75), [first_w] + [other] * 5, ["Настройка", "NPS", "CSI", "CSAT", "Голосование", "После обращения"], rows, row_h=Inches(0.4), size=11.5, colorize=colorize2)
    text(
        s,
        M,
        end_y + Inches(0.15),
        CONTENT_W,
        Inches(1.0),
        [
            {"runs": [("Персональные ссылки ", {"bold": True}), ("включаем, когда есть список участников (email, телефон, ID клиента или обращения). ", {}), ("Незавершенные ", {"bold": True}), ("— для длинных NPS и CSI. ", {}), ("IP ", {"bold": True}), ("— не для жителей ЖК: у соседей и семей общий интернет.", {})], "size": 12, "color": GRAY},
        ],
    )

    # ================================================================ 06
    section_slide(d, 6, "Сбор ответов и результаты", "Как доставить опрос людям, где смотреть ответы и как правильно считать итоги")

    s = d.new()
    d.header(s, "Каналы сбора ответов", "Все — в карточке опроса, раздел «Сбор ответов». Канал выбираем по брифу и коммуникации (шаг 6).")
    ch = [
        ("Прямая ссылка и QR-код", "Одна ссылка для всех: сайт, мессенджеры, push, печатные материалы."),
        ("Email-рассылка", "Своя ссылка каждому, email привязывается к ответу. Статистика рассылки, напоминания неответившим, база контактов. PRO+: до 15 000 писем в сутки."),
        ("SMS-рассылка (PRO+)", "Ссылка на опрос по SMS через сервис SMS Профи. Сообщения оплачиваются по тарифам SMS Профи."),
        ("Параметры ссылки", "Ссылка вида …/?channel=push. Значение сохраняется вместе с ответом: видно канал, сегмент, ID клиента. PRO+: до 50 параметров."),
        ("Виджет, окно, встраивание", "Код для сайта генерируется автоматически. Разместить его помогает разработчик."),
        ("Офлайн", "Распечатка анкеты в DOCX для бумажного сбора. Офлайн-режим — в тарифе PRO+."),
    ]
    cw = (CONTENT_W - Inches(0.25) * 2) / 3
    for i, (t, b) in enumerate(ch):
        col, row = i % 3, i // 3
        card(s, M + col * (cw + Inches(0.25)), Inches(1.75) + row * Inches(2.6), cw, Inches(2.45), t, [{"text": b, "size": 12}], accent=ORANGE if i in (1, 3) else VIOLET, title_size=14)

    s = d.new()
    d.header(s, "Где смотреть результаты", "Результаты видны в реальном времени, с первого ответа.")
    res = [
        ("Общая информация", "Ответы, переходы, доля завершивших, среднее время заполнения, устройства, браузеры, системы."),
        ("Сводный отчет", "Таблицы и диаграммы по каждому вопросу, фильтры, кросс-таблицы (PRO+). Выгрузка: PDF, XLSX, CSV, SPSS."),
        ("Отдельные ответы", "Каждая анкета целиком. Выгрузка в DOC и ZIP."),
        ("Email-уведомления", "Письма на почту о новых ответах — чтобы не заходить в кабинет вручную."),
        ("Публичная ссылка", "Ссылка на результаты, чтобы поделиться отчетом с коллегами."),
        ("API и веб-хуки (PRO+)", "Ответы автоматически уходят в наши системы. Лимит API: 120 запросов в минуту и 500 в сутки на пользователя лицензии."),
    ]
    for i, (t, b) in enumerate(res):
        col, row = i % 3, i // 3
        card(s, M + col * (cw + Inches(0.25)), Inches(1.75) + row * Inches(2.6), cw, Inches(2.45), t, [{"text": b, "size": 12}], accent=VIOLET if i != 5 else ORANGE, title_size=14)

    s = d.new()
    d.header(s, "Как считаем итоговые цифры", "Правила, по которым ответы попадают в индексы и отчеты.")
    rules = [
        ("1", "Тестовых ответов в данных нет", "Их удаляют на шаге 8, до запуска. Если что-то осталось — исключаем из расчета."),
        ("2", "В индексы берем только завершенные анкеты", "Незавершенные (пометка «Незавершен») смотрим отдельно — чтобы понять, где люди бросают опрос."),
        ("3", "Малые группы — с пометкой", "Если в сегменте (ЖК, регион) меньше 30 ответов, результат показываем как гипотезу, без жестких выводов."),
        ("4", "Время — по Москве", "Даты и часы ответов считаем по московскому времени (UTC+3)."),
        ("5", "Проверяем качество ответов (рекомендация)", "Смотрим анкеты, где во всей матрице один и тот же ответ, и слишком быстрые прохождения — их обсуждаем с владельцем исследования."),
    ]
    for i, (n, t, b) in enumerate(rules):
        y = Inches(1.7) + i * Inches(1.03)
        rect(s, M, y, CONTENT_W, Inches(0.9), WHITE, MSO_SHAPE.ROUNDED_RECTANGLE, line=LINE)
        rect(s, M + Inches(0.2), y + Inches(0.18), Inches(0.55), Inches(0.55), ORANGE, MSO_SHAPE.OVAL)
        text(s, M + Inches(0.2), y + Inches(0.18), Inches(0.55), Inches(0.55), [{"text": n, "size": 16, "bold": True, "color": WHITE, "align": PP_ALIGN.CENTER}], anchor=MSO_ANCHOR.MIDDLE)
        text(s, M + Inches(0.95), y + Inches(0.06), CONTENT_W - Inches(1.1), Inches(0.8), [{"text": t, "size": 14, "bold": True, "after": 2}, {"text": b, "size": 12.5, "color": GRAY}], anchor=MSO_ANCHOR.MIDDLE)

    s = d.new()
    d.header(s, "Чеклист перед запуском", "Пройдитесь по списку перед публикацией. Если хоть один пункт не выполнен — опрос не запускаем.")
    groups = [
        (
            "Подготовка",
            [
                "☐ Есть бриф: цель, аудитория, каналы, сроки",
                "☐ Типы вопросов выбраны под цель опроса",
                "☐ Логика настроена, обязательные вопросы отмечены",
                "☐ Ограничение повторов было выключено на время проверки",
            ],
            VIOLET,
        ),
        (
            "Проверка и согласование",
            [
                "☐ Команда прошла все ветки и утвердила опрос",
                "☐ Заказчик утвердил опрос",
                "☐ Готова коммуникация об опросе",
                f"☐ {APPROVER} согласовала опрос и сообщение",
            ],
            VIOLET_DARK,
        ),
        (
            "Запуск",
            [
                "☐ Тестовые ответы удалены — в опросе 0 ответов",
                "☐ Ограничение повторов включено",
                "☐ Задан период проведения",
                "☐ «Активный опрос» включен",
                "☐ В сообщении ссылка именно на этот опрос",
            ],
            ORANGE,
        ),
    ]
    cw = (CONTENT_W - Inches(0.25) * 2) / 3
    for i, (t, items, acc) in enumerate(groups):
        card(s, M + i * (cw + Inches(0.25)), Inches(1.75), cw, Inches(5.1), t, [{"text": it, "size": 13, "after": 10} for it in items], accent=acc)

    s = d.new()
    d.header(s, "Где искать помощь", "Если что-то непонятно в сервисе или в процессе.")
    card(
        s,
        M,
        Inches(1.8),
        Inches(6.1),
        Inches(2.8),
        "По сервису Testograf",
        [
            label_line("Справка с видео:", "testograf.ru/ru/help"),
            label_line("Мануал:", "testograf.ru/ru/blog/manual"),
            label_line("Функции по тарифам:", "testograf.ru/ru/functions"),
            label_line("Поддержка:", "info@testograf.ru, +7 (495) 120-65-19"),
        ],
        accent=VIOLET,
    )
    card(
        s,
        M + Inches(6.3),
        Inches(1.8),
        CONTENT_W - Inches(6.3),
        Inches(2.8),
        "По нашему процессу",
        [
            label_line("Согласование опроса и коммуникации:", f"{APPROVER}, руководитель департамента маркетинга."),
            label_line("Полный справочник типов вопросов и настроек:", "docs/testograf-catalog-types-and-settings.md"),
        ],
        accent=ORANGE,
    )

    return d.finish()


if __name__ == "__main__":
    out, total = build()
    print(f"Saved {out} ({total} slides)")
