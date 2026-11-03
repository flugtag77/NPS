#!/usr/bin/env python3
"""Обучающая PPTX по Testograf для клиентских опросов.

Не шпаргалка-обрывки, а объяснения: что это, зачем, пример, когда включать.
В рабочих текстах не используется буква с точками над е.
"""
from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

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
M = Inches(0.45)
OUT = Path(__file__).resolve().parent / "build" / "Testograf_playbook_client_surveys.pptx"


def _font(run, size=14, bold=False, color=INK):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = "Arial"
    rPr = run._r.get_or_add_rPr()
    for tag in ("latin", "ea", "cs"):
        el = rPr.find(qn(f"a:{tag}"))
        if el is None:
            el = rPr.makeelement(qn(f"a:{tag}"), {})
            rPr.append(el)
        el.set("typeface", "Arial")


def P(tf, text, *, size=14, bold=False, color=INK, align=PP_ALIGN.LEFT, after=6):
    p = tf.paragraphs[0] if tf.paragraphs[0].text == "" else tf.add_paragraph()
    p.text = text
    p.alignment = align
    p.space_after = Pt(after)
    for r in p.runs:
        _font(r, size=size, bold=bold, color=color)


def rect(slide, l, t, w, h, fill):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    sh.line.fill.background()
    return sh


def round_rect(slide, l, t, w, h, fill):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    sh.adjustments[0] = 0.06
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    sh.line.fill.background()
    return sh


def tb(slide, l, t, w, h):
    return slide.shapes.add_textbox(l, t, w, h)


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def header(slide, title, subtitle=""):
    rect(slide, 0, 0, SLIDE_W, Inches(0.07), ORANGE)
    rect(slide, 0, Inches(0.07), SLIDE_W, Inches(0.78), BG)
    b = tb(slide, M, Inches(0.18), Inches(2.8), Inches(0.5))
    P(b.text_frame, "ПИК Комфорт", size=15, bold=True, color=ORANGE, after=0)
    b = tb(slide, Inches(3.4), Inches(0.14), Inches(9.4), Inches(0.6))
    tf = b.text_frame
    tf.word_wrap = True
    P(tf, title, size=18, bold=True, color=INK, after=1)
    if subtitle:
        P(tf, subtitle, size=12, color=GRAY, after=0)


def footer(slide, n, total):
    b = tb(slide, M, Inches(7.15), Inches(10), Inches(0.28))
    P(b.text_frame, "Обучение Testograf · клиентские опросы", size=10, color=GRAY_LIGHT, after=0)
    b = tb(slide, Inches(11.4), Inches(7.15), Inches(1.4), Inches(0.28))
    P(b.text_frame, f"{n}/{total}", size=10, color=GRAY_LIGHT, align=PP_ALIGN.RIGHT, after=0)


def body_box(slide, l, t, w, h, lines, size=13):
    box = tb(slide, l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(lines):
        if isinstance(item, tuple):
            text, kwargs = item
            P(tf, text, **kwargs)
        else:
            P(tf, item, size=size, color=INK, after=8)
    return box


def panel(slide, l, t, w, h, title, lines, accent=VIOLET):
    round_rect(slide, l, t, w, h, WHITE)
    rect(slide, l, t, Inches(0.09), h, accent)
    box = tb(slide, l + Inches(0.22), t + Inches(0.14), w - Inches(0.35), h - Inches(0.25))
    tf = box.text_frame
    tf.word_wrap = True
    P(tf, title, size=14, bold=True, color=INK, after=8)
    for line in lines:
        P(tf, line, size=12, color=GRAY, after=5)


def build():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    slides = []

    def add():
        s = blank(prs)
        slides.append(s)
        return s

    # 1 Title
    s = add()
    rect(s, 0, 0, SLIDE_W, SLIDE_H, BG)
    rect(s, 0, 0, SLIDE_W, Inches(0.1), ORANGE)
    rect(s, 0, Inches(6.85), SLIDE_W, Inches(0.65), VIOLET_DARK)
    b = tb(s, M, Inches(1.8), Inches(12), Inches(4))
    tf = b.text_frame
    tf.word_wrap = True
    P(tf, "ПИК Комфорт · Департамент маркетинга", size=14, bold=True, color=ORANGE, after=10)
    P(tf, "Как работать в Testograf", size=34, bold=True, color=INK, after=12)
    P(
        tf,
        "Обучающая презентация: что умеет сервис, какие бывают вопросы,\n"
        "что такое логика и «блуждающие» ответы, и как мы запускаем опросы.",
        size=16,
        color=GRAY,
        after=16,
    )
    P(tf, "Для NPS, CSI, CSAT, голосований и послесервисных опросов · тариф PRO+", size=13, color=VIOLET_DARK)
    b = tb(s, M, Inches(7.0), Inches(12), Inches(0.35))
    P(b.text_frame, "Согласование запуска: Голубева Екатерина Анатольевна", size=12, color=WHITE, after=0)

    # 2 Goals of deck
    s = add()
    header(s, "Зачем эта презентация", "После нее должно быть понятно без догадок")
    panel(
        s,
        M,
        Inches(1.2),
        Inches(12.4),
        Inches(5.5),
        "Вы сможете ответить на вопросы:",
        [
            "1. Что такое Testograf и из каких частей состоит опрос.",
            "2. Чем «выбор одного» отличается от «выбора нескольких» и зачем матрицы.",
            "3. Что такое обычная логика и чем продвинутая логика от нее отличается.",
            "4. Что люди имеют в виду под «блуждающими ответами» — и как с этим работать.",
            "5. Как у нас устроен процесс от брифа до запуска (9 шагов).",
            "6. Какие настройки включать на тесте, а какие — только перед боевым стартом.",
        ],
        accent=ORANGE,
    )

    # 3 What is Testograf
    s = add()
    header(s, "Что такое Testograf", "Коротко о сервисе")
    body_box(
        s,
        M,
        Inches(1.15),
        Inches(12.4),
        Inches(5.8),
        [
            ("Testograf — российский онлайн-конструктор опросов, анкет, тестов и голосований.", {"size": 15, "bold": True, "color": INK, "after": 12}),
            ("Мы используем его, чтобы собирать мнение жителей и клиентов: NPS, удовлетворенность (CSI/CSAT), голосования и опросы после обращения в сервис.", {"size": 14, "color": GRAY, "after": 12}),
            ("Опрос в Testograf — это не «одна ссылка». Это:", {"size": 14, "bold": True, "color": INK, "after": 8}),
            ("• структура вопросов (типы, тексты, обязательность);", {"size": 14, "color": GRAY, "after": 6}),
            ("• логика (кому какой следующий вопрос показать);", {"size": 14, "color": GRAY, "after": 6}),
            ("• настройки доступа и защиты (повторы, пароли, SMS);", {"size": 14, "color": GRAY, "after": 6}),
            ("• дизайн и каналы раздачи (ссылка, email, SMS, виджет, QR);", {"size": 14, "color": GRAY, "after": 6}),
            ("• результаты и выгрузки (таблицы, фильтры, API на тарифе PRO+).", {"size": 14, "color": GRAY, "after": 6}),
        ],
    )

    # 4 Constructor steps
    s = add()
    header(s, "Как устроен конструктор", "Четыре вкладки — запомните порядок")
    steps = [
        ("1. Создание", "Страницы и вопросы. Здесь выбираете тип вопроса, тексты, обязательность, случайный порядок вариантов."),
        ("2. Логика", "Ветвление: куда идти после ответа. Дисквалификация, пропуск блоков, разные финалы."),
        ("3. Настройка", "Доступ, повторы, незавершенные, таймер, SMS, период проведения, страница завершения."),
        ("4. Дизайн", "Логотип, цвета, шрифты, фон — под бренд."),
    ]
    y = Inches(1.15)
    for title, text in steps:
        round_rect(s, M, y, Inches(12.4), Inches(1.2), WHITE)
        rect(s, M, y, Inches(0.1), Inches(1.2), ORANGE)
        b = tb(s, M + Inches(0.3), y + Inches(0.2), Inches(11.8), Inches(0.9))
        tf = b.text_frame
        tf.word_wrap = True
        P(tf, title, size=16, bold=True, color=INK, after=4)
        P(tf, text, size=13, color=GRAY, after=0)
        y += Inches(1.35)

    # 5 Question types map
    s = add()
    header(s, "Какие типы вопросов бывают", "Группы, которыми удобно думать")
    cols = [
        (
            "Выбор из списка",
            [
                "Выбор одного",
                "Выбор нескольких",
                "Выпадающий список",
                "Картинки (одно/несколько)",
                "Зачем: простые факты и предпочтения",
            ],
        ),
        (
            "Шкалы и индексы",
            [
                "NPS (0–10)",
                "CSI / Лайкерт",
                "Звезды / смайлы",
                "Семантический дифференциал",
                "Зачем: измерить оценку числом",
            ],
        ),
        (
            "Матрицы",
            [
                "Один ответ в строке",
                "Несколько в строке",
                "Звезды / текст / список",
                "Зачем: оценить много пунктов",
                "одной таблицей",
            ],
        ),
        (
            "Прочее",
            [
                "Свободный текст",
                "Контакты / файлы",
                "Ранжирование",
                "UX: first click, Figma…",
                "Зачем: детали и спецзадачи",
            ],
        ),
    ]
    x = M
    for title, lines in cols:
        panel(s, x, Inches(1.15), Inches(3.05), Inches(5.5), title, lines, accent=VIOLET)
        x += Inches(3.2)

    # 6 Single vs multi
    s = add()
    header(s, "«Выбор одного» и «выбор нескольких»", "Самые частые типы — разница критична")
    panel(
        s,
        M,
        Inches(1.15),
        Inches(6.1),
        Inches(5.5),
        "Выбор одного (single choice)",
        [
            "Респондент отмечает РОВНО один вариант.",
            "",
            "Пример: «В каком ЖК вы живете?»",
            "Пример: «Пользовались ли услугой? Да / Нет»",
            "",
            "Зачем: взаимоисключающие ответы,",
            "скринер ЦА, база для логики ветвления.",
            "",
            "Важно: от ответа часто зависит,",
            "какой следующий вопрос показывать.",
        ],
        accent=ORANGE,
    )
    panel(
        s,
        Inches(6.8),
        Inches(1.15),
        Inches(6.0),
        Inches(5.5),
        "Выбор нескольких (multi choice)",
        [
            "Респондент отмечает НЕСКОЛЬКО вариантов.",
            "",
            "Пример: «Что улучшить?» — двор, паркинг,",
            "управляющая компания, приложение…",
            "",
            "Зачем: список проблем/причин без нужды",
            "выбирать «самое главное».",
            "",
            "Риск: люди отмечают все подряд.",
            "Иногда ставят лимит «не больше N».",
            "Для логики условия сложнее (И / ИЛИ).",
        ],
        accent=VIOLET,
    )

    # 7 Scales
    s = add()
    header(s, "Шкалы: NPS, CSI, звезды", "Числовая оценка — основа наших метрик")
    body_box(
        s,
        M,
        Inches(1.15),
        Inches(12.4),
        Inches(5.8),
        [
            ("NPS — вопрос «порекомендуете ли…» по шкале 0–10. Ответы делят на критиков (0–6), нейтралов (7–8), промоутеров (9–10). Индекс = % промоутеров − % критиков.", {"size": 14, "color": INK, "after": 12}),
            ("CSI / Лайкерт — оценка согласия или удовлетворенности по пунктам (например 1–5: «совсем не согласен» → «полностью согласен»).", {"size": 14, "color": INK, "after": 12}),
            ("Звезды / смайлы — та же идея «оценка», но визуально проще для коротких CSAT.", {"size": 14, "color": INK, "after": 12}),
            ("Правило: шкалу NPS 0–10 НЕ перемешиваем случайным порядком. Цифры всегда идут по порядку, иначе метрика ломается.", {"size": 14, "bold": True, "color": ORANGE, "after": 8}),
        ],
    )

    # 8 Matrices
    s = add()
    header(s, "Что такое матрица", "Таблица: строки = объекты, столбцы = оценки")
    body_box(
        s,
        M,
        Inches(1.1),
        Inches(12.4),
        Inches(2.2),
        [
            ("Матрица нужна, когда много однотипных оценок. Вместо десяти отдельных вопросов — одна таблица.", {"size": 14, "color": INK, "after": 10}),
            ("Пример CSI: строки — «чистота», «охрана», «приложение»; столбцы — оценки от 1 до 5.", {"size": 14, "color": GRAY, "after": 8}),
        ],
    )
    mats = [
        ("Матрица: один ответ", "В каждой строке выбирается один столбец. Классика для CSI/Лайкерт."),
        ("Матрица: несколько", "В строке можно отметить несколько столбцов. Реже; осторожно с перегрузом."),
        ("Матрица: звезды", "Оценка звездами по каждой строке. Удобно визуально."),
        ("Матрица: текст / список", "Свободный текст или выпадающий список в ячейках."),
    ]
    y = Inches(3.3)
    for title, text in mats:
        round_rect(s, M, y, Inches(12.4), Inches(0.75), WHITE)
        b = tb(s, M + Inches(0.2), y + Inches(0.12), Inches(12), Inches(0.55))
        tf = b.text_frame
        tf.word_wrap = True
        P(tf, f"{title} — {text}", size=13, color=INK, after=0)
        y += Inches(0.85)
    note = tb(s, M, Inches(6.55), Inches(12.4), Inches(0.4))
    P(note.text_frame, "Риск матриц: «straight-line» — человек ставит одну колонку во все строки не читая. Ловим это на постобработке.", size=12, color=ORANGE, after=0)

    # 9 Logic basic
    s = add()
    header(s, "Что такое логика", "Обычная логика переходов (вкладка «Логика»)")
    body_box(
        s,
        M,
        Inches(1.15),
        Inches(12.4),
        Inches(5.8),
        [
            ("Логика — правило: «если человек ответил так, покажи ему вот это (или заверши опрос)».", {"size": 15, "bold": True, "color": INK, "after": 12}),
            ("Без логики все идут по одним и тем же вопросам подряд. С логикой маршрут персональный.", {"size": 14, "color": GRAY, "after": 12}),
            ("Простой пример:", {"size": 14, "bold": True, "color": INK, "after": 8}),
            ("Вопрос: «Живете ли вы в нашем ЖК?» → Да / Нет", {"size": 14, "color": GRAY, "after": 6}),
            ("Если «Нет» → дисквалификация (вежливый экран «опрос не для вас»). NPS не спрашиваем.", {"size": 14, "color": GRAY, "after": 6}),
            ("Если «Да» → дальше вопросы про двор, УК, NPS.", {"size": 14, "color": GRAY, "after": 12}),
            ("Обычная логика обычно смотрит на ОДИН предыдущий вопрос (выбрал / не выбрал вариант).", {"size": 14, "color": INK, "after": 8}),
            ("Куда можно отправить: на вопрос, на страницу, на стандартное/индивидуальное завершение, на дисквалификацию.", {"size": 14, "color": GRAY, "after": 8}),
        ],
    )

    # 10 Advanced logic
    s = add()
    header(s, "Что такое продвинутая логика", "Только PRO+ · вкладка «Продвинутая логика»")
    body_box(
        s,
        M,
        Inches(1.1),
        Inches(12.4),
        Inches(5.9),
        [
            ("Продвинутая логика нужна, когда одного условия мало. Она умеет комбинировать ответы на РАЗНЫЕ вопросы и параметры ссылки через И / ИЛИ.", {"size": 14, "bold": True, "color": INK, "after": 12}),
            ("Пример 1: показать блок «паркинг», только если человек живет в ЖК И отметил, что у него есть машина.", {"size": 14, "color": GRAY, "after": 10}),
            ("Пример 2: если в параметре ссылки region=msk ИЛИ человек выбрал Москву в анкете — показать московские формулировки.", {"size": 14, "color": GRAY, "after": 10}),
            ("Пример 3: в матрице — условие по конкретной ячейке («в строке 2 выбран столбец 3»).", {"size": 14, "color": GRAY, "after": 12}),
            ("Чем отличается от обычной:", {"size": 14, "bold": True, "color": INK, "after": 8}),
            ("• обычная — проще, чаще от одного вопроса;", {"size": 14, "color": GRAY, "after": 6}),
            ("• продвинутая — несколько условий сразу, параметры ссылки, гибкий показ/пропуск блоков.", {"size": 14, "color": GRAY, "after": 10}),
            ("Обязательно: перед запуском прогнать ВСЕ ветки в предпросмотре. Ошибка в логике = все ушли не туда.", {"size": 14, "bold": True, "color": ORANGE, "after": 6}),
        ],
    )

    # 11 Wandering answers - THE KEY SLIDE
    s = add()
    header(s, "Что такое «блуждающие» ответы", "Это не одна кнопка в Testograf. Обычно имеют в виду одно из трех")
    panel(
        s,
        M,
        Inches(1.15),
        Inches(4.0),
        Inches(5.5),
        "Смысл A. Случайный порядок",
        [
            "В сервисе это называется",
            "рандомизация / случайный порядок",
            "вопросов или вариантов ответа.",
            "",
            "Зачем: чтобы первый пункт списка",
            "не получал лишних голосов.",
            "",
            "Вкл: списки равноправных опций.",
            "Выкл: шкала NPS 0–10 и логичные",
            "последовательности страниц.",
        ],
        accent=VIOLET,
    )
    panel(
        s,
        Inches(4.65),
        Inches(1.15),
        Inches(4.0),
        Inches(5.5),
        "Смысл B. Случайные клики",
        [
            "Респондент «тыкает наугад»,",
            "чтобы быстрее закончить.",
            "",
            "Признаки: очень короткое время,",
            "одна колонка во всей матрице,",
            "противоречия в ответах.",
            "",
            "В UI это НЕ включается.",
            "Профилактика: без таймера,",
            "не раздувать анкету.",
            "Лечение: чистка после сбора.",
        ],
        accent=ORANGE,
    )
    panel(
        s,
        Inches(8.85),
        Inches(1.15),
        Inches(4.0),
        Inches(5.5),
        "Смысл C. Бросили на полпути",
        [
            "Человек начал и ушел.",
            "В Testograf это незавершенные",
            "ответы (PRO+).",
            "",
            "Можно сохранить черновик",
            "и дать вернуться.",
            "",
            "В индекс NPS/CSI берем",
            "только завершенные анкеты.",
            "Незавершенные — для анализа,",
            "где люди отваливаются.",
        ],
        accent=GREEN,
    )

    # 12 Randomization detail
    s = add()
    header(s, "Рандомизация подробно", "Смысл A «блуждающих» — как настроить правильно")
    body_box(
        s,
        M,
        Inches(1.15),
        Inches(12.4),
        Inches(5.8),
        [
            ("Где включается: в настройках страницы — случайный порядок вопросов; в настройках вопроса — случайный порядок вариантов ответа.", {"size": 14, "color": INK, "after": 12}),
            ("Когда ВКЛЮЧАТЬ: список факторов «что улучшить», варианты дизайна, бренды/опции без «правильного» порядка.", {"size": 14, "color": GRAY, "after": 10}),
            ("Когда ВЫКЛЮЧАТЬ: NPS 0–10; шкалы «от плохо к хорошо»; страницы, которые должны идти смысловым сюжетом (профиль → оценка → причины).", {"size": 14, "color": GRAY, "after": 10}),
            ("Связь с логикой: если ветвление завязано на конкретный вариант, рандомизация вариантов обычно не мешает (условие по ID ответа, не по месту на экране). Но тестировать все равно нужно.", {"size": 14, "color": GRAY, "after": 10}),
            ("Итог: «блуждающий порядок» — инструмент против смещения. «Блуждающие клики» — проблема качества данных. Не путайте.", {"size": 14, "bold": True, "color": ORANGE, "after": 6}),
        ],
    )

    # 13 Repeats
    s = add()
    header(s, "Ограничение повторного прохождения", "Чтобы один человек не накрутил результаты")
    body_box(
        s,
        M,
        Inches(1.15),
        Inches(12.4),
        Inches(5.8),
        [
            ("Зачем: в аналитике должен быть принцип «один респондент — один учтенный ответ» (если бриф не говорит иное).", {"size": 14, "color": INK, "after": 10}),
            ("Устройство / cookies — базовая защита. Где: Настройка → Основное → блок «Вход». Легко обойти инкогнито или другим телефоном.", {"size": 14, "color": GRAY, "after": 8}),
            ("IP — строже, но опасно для ЖК: у многих жителей один Wi‑Fi / один NAT. Для наших опросов жителей обычно НЕ включаем.", {"size": 14, "color": GRAY, "after": 8}),
            ("Персональная ссылка / параметр ID — самый надежный способ, когда знаем участника (CSAT после заявки, рассылка).", {"size": 14, "color": GRAY, "after": 8}),
            ("SMS-подтверждение (PRO+) — для голосований с риском накрутки: один номер = один голос.", {"size": 14, "color": GRAY, "after": 12}),
            ("Наш процесс: на тесте ограничение ВЫКЛЮЧЕНО (шаг 3). Перед запуском ВКЛЮЧЕНО (шаг 8), после зачистки тестовых ответов.", {"size": 14, "bold": True, "color": ORANGE, "after": 6}),
        ],
    )

    # 14 Incomplete
    s = add()
    header(s, "Незавершенные ответы", "Смысл C «блуждающих» · функция PRO+")
    body_box(
        s,
        M,
        Inches(1.15),
        Inches(12.4),
        Inches(5.8),
        [
            ("Если включить сохранение незавершенных: человек может закрыть анкету и позже продолжить с того же места (по именной ссылке / паролю / токену в браузере).", {"size": 14, "color": INK, "after": 12}),
            ("В личном кабинете такие ответы видны с пометкой «Незавершен» — видно, на каком вопросе остановились.", {"size": 14, "color": GRAY, "after": 12}),
            ("Когда включать: длинные NPS/CSI.", {"size": 14, "color": GRAY, "after": 8}),
            ("Когда выключать: короткие CSAT и голосования «в один экран».", {"size": 14, "color": GRAY, "after": 12}),
            ("Жесткое правило аналитики: в расчет индекса идут только завершенные. Иначе сравниваем «яблоки с недоеденными яблоками».", {"size": 14, "bold": True, "color": ORANGE, "after": 8}),
        ],
    )

    # 15 Process
    s = add()
    header(s, "Как мы запускаем опрос", "Обязательные 9 шагов")
    steps = [
        "1. Бриф от заказчика (или готовим сами, если заказчик — мы).",
        "2. Заводим опрос в Testograf (создание → логика → настройка → дизайн).",
        "3. Отключаем ограничение повторного прохождения — иначе нельзя нормально тестировать.",
        "4. Внутренняя команда проверяет логику, тексты, ошибки, мобильный вид.",
        "5. После внутренней проверки — отдаем заказчику.",
        "6. Параллельно готовим коммуникацию (push / ТГ / email / ЛК).",
        "7. Опрос И текст коммуникации согласовывает Голубева Екатерина Анатольевна.",
        "8. Зачищаем тестовые ответы и ВКЛЮЧАЕМ ограничение «нельзя пройти дважды».",
        "9. Запуск: период проведения, каналы, фиксация даты старта.",
    ]
    b = tb(s, M, Inches(1.1), Inches(12.4), Inches(5.8))
    tf = b.text_frame
    tf.word_wrap = True
    for line in steps:
        P(tf, line, size=13, color=INK, after=7)

    # 16 Why step 3 and 8
    s = add()
    header(s, "Почему так с ограничением повторов", "Самая частая путаница в процессе")
    panel(
        s,
        M,
        Inches(1.2),
        Inches(6.1),
        Inches(5.3),
        "На тесте (шаги 3–7) — ВЫКЛ",
        [
            "Команда и заказчик должны пройти",
            "опрос много раз: разные ветки логики,",
            "мобильный Safari, правки текстов.",
            "",
            "Если ограничение включено слишком рано,",
            "проверка ломается: «мне пишет,",
            "что я уже проходил».",
        ],
        accent=ORANGE,
    )
    panel(
        s,
        Inches(6.8),
        Inches(1.2),
        Inches(6.0),
        Inches(5.3),
        "Перед запуском (шаг 8) — ВКЛ",
        [
            "Сначала удаляем все тестовые ответы.",
            "Потом включаем защиту от повторов.",
            "Только после этого — боевая ссылка",
            "в коммуникации.",
            "",
            "Если нашли баг после старта:",
            "выкл → фикс → зачистка тестов →",
            "снова вкл → продолжаем сбор.",
        ],
        accent=GREEN,
    )

    # 17 Post processing
    s = add()
    header(s, "После сбора: что выкидываем из аналитики", "Даже при идеальных настройках нужна чистка")
    body_box(
        s,
        M,
        Inches(1.15),
        Inches(12.4),
        Inches(5.8),
        [
            ("1. Тестовые прогоны команды и заказчика — удалить до запуска; если остались, выкинуть из базы.", {"size": 14, "color": INK, "after": 10}),
            ("2. Незавершенные — не считать в NPS/CSI/CSAT.", {"size": 14, "color": INK, "after": 10}),
            ("3. Слишком быстрые / «тыканье» / straight-line по матрице — помечать как низкое качество и не включать в индекс.", {"size": 14, "color": INK, "after": 10}),
            ("4. Дубли одного человека / одного обращения — оставить один ответ по правилу брифа.", {"size": 14, "color": INK, "after": 10}),
            ("5. Сегменты с n < 30 — показывать только как гипотезу, без жестких выводов.", {"size": 14, "color": INK, "after": 10}),
            ("6. Часовой пояс учета событий — Москва (UTC+3).", {"size": 14, "color": INK, "after": 8}),
        ],
    )

    # 18 Checklist
    s = add()
    header(s, "Чеклист перед запуском", "Если чего-то нет — не стартуем")
    items = [
        "□ Бриф понятен: цель, ЦА, каналы, нужен ли 1 ответ на человека",
        "□ Типы вопросов выбраны осознанно (один / несколько / шкала / матрица)",
        "□ Логика (и продвинутая, если есть) прогнана по всем веткам",
        "□ NPS-шкала без рандомизации вариантов",
        "□ Согласование Голубевой Е. А. по опросу и коммуникации получено",
        "□ Тестовые ответы удалены",
        "□ Ограничение повторного прохождения включено",
        "□ Показ промежуточных итогов респондентам выключен",
        "□ В аналитике заранее договорено: только завершенные",
    ]
    b = tb(s, M, Inches(1.15), Inches(12.4), Inches(5.8))
    tf = b.text_frame
    tf.word_wrap = True
    for line in items:
        P(tf, line, size=15, color=INK, after=10)

    # 19 Mistakes
    s = add()
    header(s, "Типичные ошибки простым языком", "Что идет не так и почему это плохо")
    errs = [
        ("Путают логику и рандомизацию", "Логика решает КУДА идти. Рандомизация решает В КАКОМ ПОРЯДКЕ показать варианты."),
        ("Рандомят шкалу NPS", "Цифры 0–10 должны стоять по порядку. Иначе индекс нельзя интерпретировать."),
        ("Считают незавершенные в NPS", "Сравнивают полные и оборванные анкеты — метрика плывет."),
        ("Включают IP-лимит для ЖК", "Соседи на одном Wi‑Fi: ответит только первый."),
        ("Запускают без шага 7", "Нет согласования Голубевой Е. А. — запуск запрещен процессом."),
        ("Оставляют тестовые ответы", "Свои «проверки» портят индекс наравне с живыми ответами."),
    ]
    y = Inches(1.1)
    for title, text in errs:
        round_rect(s, M, y, Inches(12.4), Inches(0.85), WHITE)
        rect(s, M, y, Inches(0.09), Inches(0.85), RED)
        b = tb(s, M + Inches(0.25), y + Inches(0.12), Inches(11.9), Inches(0.65))
        tf = b.text_frame
        tf.word_wrap = True
        P(tf, title, size=13, bold=True, color=INK, after=2)
        P(tf, text, size=12, color=GRAY, after=0)
        y += Inches(0.92)

    # 20 Close
    s = add()
    rect(s, 0, 0, SLIDE_W, SLIDE_H, BG)
    rect(s, 0, 0, SLIDE_W, Inches(0.1), ORANGE)
    b = tb(s, M, Inches(2.0), Inches(12), Inches(4))
    tf = b.text_frame
    tf.word_wrap = True
    P(tf, "Короткий итог", size=14, bold=True, color=ORANGE, after=10)
    P(tf, "Сначала понимаем инструмент,\nпотом нажимаем настройки", size=28, bold=True, color=INK, after=14)
    P(
        tf,
        "Тип вопроса задает, ЧТО можно ответить.\n"
        "Логика задает, КОМУ ЧТО показать.\n"
        "«Блуждающие» — это либо рандом порядка, либо плохие клики, либо незавершенные.\n"
        "Запуск у нас — только через 9 шагов и согласование Голубевой Е. А.",
        size=15,
        color=GRAY,
        after=16,
    )
    P(tf, "Полный текст: docs/testograf-playbook-client-surveys.md", size=13, color=VIOLET_DARK)

    total = len(slides)
    for i, slide in enumerate(slides, start=1):
        if i in (1, total):
            continue
        footer(slide, i, total)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(OUT))
    print(f"Saved {OUT} ({total} slides)")
    return OUT


if __name__ == "__main__":
    build()
