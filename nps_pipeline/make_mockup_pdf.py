"""NPS-2026 · PDF-макет продукта на 10 листов в фирменном стиле ПИК-Комфорт.
Формат A4 landscape. Шрифт Inter (аналог Graphik LCG).
Палитра из брендбука 2026.
"""
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors
from reportlab.lib.units import mm

# --- Fonts ---
pdfmetrics.registerFont(TTFont("Graphik",     "/home/user/workspace/fonts/inter/Inter-Regular.ttf"))
pdfmetrics.registerFont(TTFont("Graphik-Md",  "/home/user/workspace/fonts/inter/Inter-Medium.ttf"))
pdfmetrics.registerFont(TTFont("Graphik-Sb",  "/home/user/workspace/fonts/inter/Inter-SemiBold.ttf"))
pdfmetrics.registerFont(TTFont("Graphik-Bd",  "/home/user/workspace/fonts/inter/Inter-Bold.ttf"))

# --- Brand palette ---
VIOLET       = colors.HexColor("#9678D2")
VIOLET_DARK  = colors.HexColor("#7C5EB8")
VIOLET_MED   = colors.HexColor("#A68ADE")
VIOLET_LIGHT = colors.HexColor("#B399E8")
ORANGE       = colors.HexColor("#FC4C02")
PEACH        = colors.HexColor("#FFB9A2")
GRAY_DARK    = colors.HexColor("#686868")
GRAY_LIGHT   = colors.HexColor("#B2B2B2")
INK          = colors.HexColor("#1A1A1A")
BG_LIGHT     = colors.HexColor("#F7F5FB")

# Сегментные (стандарт NPS)
GREEN        = colors.HexColor("#2FA84F")
YELLOW       = colors.HexColor("#F5B800")
RED          = colors.HexColor("#E24C4B")

PAGE_W, PAGE_H = landscape(A4)  # 842 x 595 pt
MARGIN = 32
CONTENT_W = PAGE_W - 2*MARGIN
CONTENT_H = PAGE_H - 2*MARGIN

OUT = "/home/user/workspace/nps_pipeline/build/NPS-2026_mockup.pdf"

# --- Полный периметр 96 ЖК: подключаем на этапе импорта ---
import csv as _csv
ZHK_FULL = []
with open("/home/user/workspace/nps_pipeline/db/zhk_full.tsv", encoding="utf-8") as _f:
    _r = _csv.DictReader(_f, delimiter="\t")
    for _row in _r:
        _row["num"] = int(_row["num"])
        _row["n26_int"] = int(_row["n26"]) if _row["n26"].isdigit() else 0
        ZHK_FULL.append(_row)

def _seg_color(seg):
    return {"Звезды":GREEN,"Восходящие":GREEN,"Спящие":YELLOW,"Кризис":RED,"Недобор":GRAY_DARK}.get(seg, GRAY_DARK)


# ==============================================================
def page_header(c, page_no, title, subtitle=""):
    """Фирменная шапка + оранжевая полоса + номер страницы + логотип-плейсхолдер."""
    # Верхняя оранжевая тонкая полоса
    c.setFillColor(ORANGE)
    c.rect(0, PAGE_H-6, PAGE_W, 6, stroke=0, fill=1)

    # Логотип ПИК-Комфорт (текстовый плейсхолдер)
    c.setFillColor(ORANGE)
    c.setFont("Graphik-Bd", 20)
    c.drawString(MARGIN, PAGE_H-38, "ПИК")
    pik_w = c.stringWidth("ПИК", "Graphik-Bd", 20)
    c.setFillColor(INK)
    c.setFont("Graphik-Md", 14)
    c.drawString(MARGIN+pik_w+4, PAGE_H-38, "Комфорт")

    # Раздел / заголовок листа
    c.setFillColor(INK)
    c.setFont("Graphik-Sb", 16)
    c.drawString(MARGIN+200, PAGE_H-32, title)
    if subtitle:
        c.setFillColor(GRAY_DARK)
        c.setFont("Graphik", 10)
        c.drawString(MARGIN+200, PAGE_H-46, subtitle)

    # Номер страницы (тонкий кружок фиолетового)
    c.setFillColor(VIOLET)
    c.circle(PAGE_W-MARGIN-14, PAGE_H-30, 12, stroke=0, fill=1)
    c.setFillColor(colors.white)
    c.setFont("Graphik-Sb", 11)
    c.drawCentredString(PAGE_W-MARGIN-14, PAGE_H-33, f"{page_no:02d}")

    # Нижний футер
    c.setFillColor(GRAY_LIGHT)
    c.setFont("Graphik", 8)
    c.drawString(MARGIN, 18, "NPS-2026 · макет продукта · внутренний документ")
    c.drawRightString(PAGE_W-MARGIN, 18, "10.09.2026")

def box(c, x, y, w, h, fill=None, stroke=None, radius=8):
    if fill:
        c.setFillColor(fill); c.setStrokeColor(fill)
    if stroke:
        c.setStrokeColor(stroke)
    c.roundRect(x, y, w, h, radius, stroke=1 if stroke else 0, fill=1 if fill else 0)

def kpi_tile(c, x, y, w, h, label, value, sub, tile_color=VIOLET, value_color=None, value_size=34, label_size=9):
    box(c, x, y, w, h, fill=tile_color, radius=10)
    c.setFillColor(colors.white)
    c.setFont("Graphik-Md", label_size)
    c.drawString(x+14, y+h-18, label.upper())
    c.setFillColor(value_color or colors.white)
    c.setFont("Graphik-Bd", value_size)
    # Значение по центру между лейблом и подписью
    c.drawString(x+14, y+h-18-value_size-4, value)
    if sub:
        c.setFillColor(colors.white)
        c.setFont("Graphik", 9)
        lines = sub.split("\n")
        # снизу вверх, чтобы не наехать на значение
        for i, line in enumerate(lines):
            c.drawString(x+14, y+12+(len(lines)-1-i)*11, line)

def h2(c, x, y, text, color=INK):
    """CAPS-заголовок Semibold с трекингом — стиль референса NPS-2025."""
    text_up = text.upper().replace("Ё","Е")
    to = c.beginText(x, y)
    to.setFillColor(color)
    to.setFont("Graphik-Sb", 11)
    to.setCharSpace(0.7)
    to.textOut(text_up)
    c.drawText(to)
    # Тонкая подчерка
    tw = c.stringWidth(text_up, "Graphik-Sb", 11) + (len(text_up)-1)*0.7
    c.setStrokeColor(VIOLET); c.setLineWidth(1.2)
    c.line(x, y-4, x+min(tw, 60), y-4)

def caption(c, x, y, text, color=GRAY_DARK, font_size=8):
    c.setFillColor(color); c.setFont("Graphik", font_size)
    c.drawString(x, y, text)

def hline(c, x1, y, x2, color=GRAY_LIGHT):
    c.setStrokeColor(color); c.setLineWidth(0.5); c.line(x1,y,x2,y)

def table(c, x, y, col_widths, rows, header_fill=VIOLET, header_text=colors.white,
          row_h=18, header_font="Graphik-Sb", body_font="Graphik", stripe=BG_LIGHT,
          highlight_col=None, highlight_color=ORANGE):
    """Простая таблица с шапкой, зебра-строками."""
    total_w = sum(col_widths)
    # Header
    c.setFillColor(header_fill)
    c.rect(x, y-row_h, total_w, row_h, stroke=0, fill=1)
    c.setFillColor(header_text)
    c.setFont(header_font, 9)
    cx = x
    for i,(cw, cell) in enumerate(zip(col_widths, rows[0])):
        c.drawString(cx+6, y-row_h+5, str(cell))
        cx += cw
    # Body
    cy = y - row_h
    for ri, row in enumerate(rows[1:]):
        cy -= row_h
        if ri % 2 == 0:
            c.setFillColor(stripe)
            c.rect(x, cy, total_w, row_h, stroke=0, fill=1)
        cx = x
        for i,(cw, cell) in enumerate(zip(col_widths, row)):
            is_hl = (highlight_col is not None and i == highlight_col)
            c.setFillColor(highlight_color if is_hl else INK)
            c.setFont("Graphik-Sb" if is_hl else body_font, 9)
            c.drawString(cx+6, cy+5, str(cell))
            cx += cw
    # Внешняя рамка
    c.setStrokeColor(GRAY_LIGHT); c.setLineWidth(0.4)
    c.rect(x, cy, total_w, y-cy, stroke=1, fill=0)
    return cy


def table_compact(c, x, y, col_widths, header, body_rows, header_fill=VIOLET_DARK,
                  header_text=colors.white, row_h=10, font_size=6.5,
                  seg_col=None, nps_col=None, contrib_col=None):
    """Компактная таблица для 96+ строк на одной странице.
       row_h=10, шрифт=6.5pt. Заливка сегмента и цвет NPS по знаку."""
    total_w = sum(col_widths)
    # header
    c.setFillColor(header_fill)
    c.rect(x, y-row_h-3, total_w, row_h+3, stroke=0, fill=1)
    c.setFillColor(header_text); c.setFont("Graphik-Sb", font_size+0.5)
    cx = x
    for cw, cell in zip(col_widths, header):
        c.drawString(cx+3, y-row_h+1, str(cell))
        cx += cw
    # body
    cy = y - row_h - 3
    for ri, row in enumerate(body_rows):
        cy -= row_h
        if ri % 2 == 0:
            c.setFillColor(BG_LIGHT); c.rect(x, cy, total_w, row_h, stroke=0, fill=1)
        cx = x
        for i, (cw, cell) in enumerate(zip(col_widths, row)):
            txt = str(cell)
            # цвет
            col = INK; font = "Graphik"
            if seg_col is not None and i == seg_col:
                col = _seg_color(row[seg_col]); font = "Graphik-Sb"
            elif nps_col is not None and i == nps_col:
                # знак NPS красит текст
                if txt.startswith("+"): col = GREEN; font = "Graphik-Sb"
                elif txt.startswith("−") or txt.startswith("-"): col = RED; font = "Graphik-Sb"
            elif contrib_col is not None and i == contrib_col:
                if txt.startswith("+"): col = GREEN
                elif txt.startswith("−") or txt.startswith("-"): col = ORANGE
                font = "Graphik-Sb"
            c.setFillColor(col); c.setFont(font, font_size)
            # обрезка длинного имени
            max_chars = int(cw / (font_size*0.42))
            if len(txt) > max_chars:
                txt = txt[:max_chars-1] + "…"
            c.drawString(cx+3, cy+2.5, txt)
            cx += cw
    c.setStrokeColor(GRAY_LIGHT); c.setLineWidth(0.3)
    c.rect(x, cy, total_w, y-cy, stroke=1, fill=0)
    return cy

# ==============================================================
# ================== СТРАНИЦЫ ==================================
# ==============================================================

def page_00_cover(c):
    """Обложка в стиле брендбука."""
    # фиолетовый фон
    c.setFillColor(VIOLET); c.rect(0,0,PAGE_W,PAGE_H,stroke=0,fill=1)
    # Логотип белый
    c.setFillColor(colors.white); c.setFont("Graphik-Bd", 30)
    c.drawString(MARGIN+8, PAGE_H-70, "ПИК")
    c.setFillColor(colors.HexColor("#D6C6EF"))
    c.setFont("Graphik-Md", 22)
    c.drawString(MARGIN+72, PAGE_H-70, "Комфорт")

    # Заголовок
    c.setFillColor(colors.white); c.setFont("Graphik-Bd", 56)
    c.drawString(MARGIN+8, PAGE_H/2-10, "NPS-2026")
    c.setFillColor(colors.white); c.setFont("Graphik-Md", 22)
    c.drawString(MARGIN+8, PAGE_H/2-46, "Оперативный мониторинг")

    # Оранжевая полоса + год
    c.setFillColor(ORANGE); c.rect(0, 42, PAGE_W, 46, stroke=0, fill=1)
    c.setFillColor(colors.white); c.setFont("Graphik-Bd", 22)
    c.drawRightString(PAGE_W-MARGIN, 56, "2026")

    # Служебная подпись
    c.setFillColor(colors.white); c.setFont("Graphik", 10)
    c.drawString(MARGIN+8, 108, "Макет продукта · 10 листов · для согласования")

    c.showPage()

def page_content(c, num, title, subtitle, draw_body):
    """Обертка: контентная страница с шапкой, телом и служебным колонтитулом."""
    c.setFillColor(colors.white); c.rect(0,0,PAGE_W,PAGE_H,stroke=0,fill=1)
    page_header(c, num, title, subtitle)
    draw_body(c)
    # Служебный колонтитул (стиль референса NPS-2025)
    c.setStrokeColor(GRAY_LIGHT); c.setLineWidth(0.4)
    c.line(MARGIN, 24, PAGE_W-MARGIN, 24)
    c.setFillColor(GRAY_DARK); c.setFont("Graphik", 8)
    c.drawString(MARGIN, 12, "NPS-2026 · оперативный мониторинг · макет продукта")
    c.setFillColor(GRAY_DARK); c.setFont("Graphik-Md", 8)
    c.drawRightString(PAGE_W-MARGIN, 12, f"стр. {num} / 10")
    c.showPage()

# ---------- Лист 1 «Лист» ----------
def body_p1(c):
    y_top = PAGE_H - 80
    # Ряд 1: 2 крупные KPI + 4 малые опорные плитки справа
    big_h = 150
    kpi_tile(c, MARGIN, y_top-big_h, 235, big_h,
             "NPS текущий · n=1797", "−53,5",
             "промоутеры  16,7%\nнейтралы    13,1%\nкритики     70,2%",
             tile_color=VIOLET_DARK, value_size=42)

    kpi_tile(c, MARGIN+245, y_top-big_h, 235, big_h,
             "Прогноз NPS к 30.09", "−41,2",
             "динамика 2026 + профиль 2025\nΔ к KPI (−35):  −6,2 пп\nΔ к цели (−30): −11,2 пп",
             tile_color=ORANGE, value_size=42)

    # Малые опорные плитки справа — 2×2 сетка, суммарная высота = big_h
    x0 = MARGIN + 490
    small_w = 145
    small_h = (big_h - 8) / 2
    kpi_tile(c, x0,               y_top-small_h,       small_w, small_h, "2025 (факт)",  "−32",  "", tile_color=GRAY_DARK,    value_size=24)
    kpi_tile(c, x0+small_w+8,     y_top-small_h,       small_w, small_h, "KPI 2026",     "−35",  "", tile_color=VIOLET_LIGHT, value_size=24, value_color=INK)
    kpi_tile(c, x0,               y_top-big_h,         small_w, small_h, "Цель 2026",    "−30",  "", tile_color=ORANGE,       value_size=24)
    kpi_tile(c, x0+small_w+8,     y_top-big_h,         small_w, small_h, "Покрытие ЖК",  "76/96","", tile_color=VIOLET,       value_size=24)

    caption(c, x0, y_top-big_h-14, "12 ЖК в зоне риска недобора (n < 50)")

    # Стек сегментов
    y = y_top - big_h - 40
    h2(c, MARGIN, y, "Состав ответов · промоутеры / нейтралы / критики")
    y -= 18; bar_w = CONTENT_W; bar_h = 22
    # Пропорции
    p, n, d = 16.7, 13.1, 70.2
    total = p+n+d
    wp, wn, wd = bar_w*p/total, bar_w*n/total, bar_w*d/total
    c.setFillColor(GREEN);  c.rect(MARGIN, y-bar_h, wp, bar_h, stroke=0, fill=1)
    c.setFillColor(YELLOW); c.rect(MARGIN+wp, y-bar_h, wn, bar_h, stroke=0, fill=1)
    c.setFillColor(RED);    c.rect(MARGIN+wp+wn, y-bar_h, wd, bar_h, stroke=0, fill=1)
    c.setFillColor(colors.white); c.setFont("Graphik-Sb", 10)
    c.drawString(MARGIN+8, y-bar_h+7, f"{p}%")
    c.drawString(MARGIN+wp+8, y-bar_h+7, f"{n}%")
    c.drawString(MARGIN+wp+wn+8, y-bar_h+7, f"{d}%")

    # CSI плитки
    y -= bar_h + 30
    h2(c, MARGIN, y, "CSI взвешенный 64,5%  ·  CSAT УК (4+5) 29,1%  ·  топ-3 провала")
    y -= 20
    tiles = [
      ("Оператив.",     "57,8%", RED),
      ("Лифты",         "57,9%", RED),
      ("Уборка МОП",    "59,6%", RED),
      ("Информир.",     "62,4%", ORANGE),
      ("Инженерия",     "63,1%", ORANGE),
      ("Домофон",       "66,2%", ORANGE),
      ("Домоф. 2",      "66,8%", ORANGE),
      ("Пожарная",      "68,0%", ORANGE),
      ("ЛК/сайт",       "70,0%", VIOLET),
      ("Территор.",     "70,5%", VIOLET),
      ("Мебель",        "72,4%", VIOLET),
      ("Озеленен.",     "73,8%", VIOLET),
      ("Соседи",        "74,2%", VIOLET),
    ]
    gap = 5
    tw = (CONTENT_W - gap*12) / 13
    for i,(name, val, col) in enumerate(tiles):
        x = MARGIN + i*(tw+gap)
        box(c, x, y-64, tw, 64, fill=col, radius=6)
        c.setFillColor(colors.white); c.setFont("Graphik-Sb", 11)
        c.drawString(x+6, y-24, val)
        c.setFillColor(colors.white); c.setFont("Graphik", 7)
        c.drawString(x+6, y-42, name)

    # Мини-буллет в правом нижнем углу: прогноз при текущей динамике
    box(c, PAGE_W-MARGIN-260, 40, 260, 40, fill=BG_LIGHT, stroke=VIOLET_LIGHT, radius=6)
    c.setFillColor(VIOLET_DARK); c.setFont("Graphik-Sb", 9)
    c.drawString(PAGE_W-MARGIN-250, 66, "ПРОГНОЗ NPS ПРИ ТЕКУЩЕЙ ДИНАМИКЕ")
    c.setFillColor(ORANGE); c.setFont("Graphik-Bd", 16)
    c.drawString(PAGE_W-MARGIN-250, 48, "−41,2")
    c.setFillColor(GRAY_DARK); c.setFont("Graphik", 8)
    c.drawString(PAGE_W-MARGIN-200, 51, "к 30.09 · 70/30 факт/история")

# ---------- Лист 2 Прогноз ----------
def body_p2(c):
    y_head = PAGE_H - 80
    # Композиция референса: большой KPI слева + мини-график сверху + все 96 ЖК внизу
    # Левая крупная плитка-прогноз (стиль референса — крупные цифры)
    kpi_tile(c, MARGIN, y_head-120, 260, 120,
             "Прогноз NPS к 30.09",
             "−41,2",
             "динамика 2026 + профиль 2025\nΔ к KPI (−35):  −6,2 пп\nΔ к цели (−30): −11,2 пп",
             tile_color=ORANGE, value_size=52)

    # Мини-график справа
    gx1, gy1 = MARGIN+290, y_head-110
    gx2, gy2 = MARGIN+CONTENT_W-10, y_head-10
    box(c, gx1-8, gy1-14, gx2-gx1+16, gy2-gy1+22, fill=BG_LIGHT, radius=8)
    caption(c, gx1, gy2+4, "Динамика NPS по дням: факт vs. прогноз", color=INK, font_size=9)
    c.setStrokeColor(GRAY_DARK); c.setLineWidth(0.5)
    c.line(gx1, gy1, gx1, gy2); c.line(gx1, gy1, gx2, gy1)
    def yv(v):
        vmin, vmax = -80, -20
        return gy1 + (v-vmin)/(vmax-vmin)*(gy2-gy1)
    for lvl,lbl,col in [(-30,"цель −30",ORANGE),(-35,"KPI −35",VIOLET_DARK),(-32,"2025 −32",GRAY_DARK)]:
        yy = yv(lvl); c.setStrokeColor(col); c.setDash(2,2); c.line(gx1, yy, gx2-70, yy); c.setDash()
        c.setFillColor(col); c.setFont("Graphik-Md", 7); c.drawString(gx2-64, yy-2, lbl)
    pts  = [(-62,"08"),(-58,"09"),(-53,"10"),(-53,"11"),(-51,"12"),(-49,"13"),(-47,"14")]
    prog = [(-45,"15"),(-43,"18"),(-42,"22"),(-41,"26"),(-41,"30")]
    xs = [gx1 + i*(gx2-gx1)/16 for i in range(1,8)]
    c.setStrokeColor(VIOLET_DARK); c.setLineWidth(1.8)
    for i in range(len(pts)-1):
        c.line(xs[i], yv(pts[i][0]), xs[i+1], yv(pts[i+1][0]))
    for i,(v,d) in enumerate(pts):
        c.setFillColor(VIOLET_DARK); c.circle(xs[i], yv(v), 2, stroke=0, fill=1)
        c.setFont("Graphik", 6); c.setFillColor(GRAY_DARK); c.drawCentredString(xs[i], gy1-10, d)
    xs2 = [gx1 + i*(gx2-gx1)/16 for i in range(7,13)]
    c.setStrokeColor(ORANGE); c.setDash(3,2); c.setLineWidth(1.6)
    prev = xs[-1], yv(pts[-1][0])
    for (v,d),xx in zip(prog, xs2):
        c.line(prev[0], prev[1], xx, yv(v))
        prev = (xx, yv(v))
        c.setFillColor(GRAY_DARK); c.setFont("Graphik", 6); c.drawCentredString(xx, gy1-10, d)
    c.setDash()

    # ---- Полный список 96 ЖК по потенциалу вклада ----
    ty = y_head - 130
    h2(c, MARGIN, ty, "ВСЕ 96 ЖК ПО ПОТЕНЦИАЛУ ВКЛАДА · ранжирование по вкладу в общий NPS (пп)")
    caption(c, MARGIN, ty-10, "«−» вклад означает вклад в снижение общего NPS. n<50 — маркер «недобор»: NPS ЖК ещё не стабилизирован.", color=GRAY_DARK, font_size=7)

    # ранжируем: сначала худшие вкладчики (наиболее отрицательные), потом лучшие
    def _key(z):
        v = z["vklad"].replace("+","").replace("−","-").replace(",",".")
        try: return float(v)
        except: return 0.0
    ranked = sorted(ZHK_FULL, key=_key)  # худшие вверху

    # 96 строк в 2 колонки по 48 строк
    n_per_col = 48
    col_w = [16, 118, 34, 24, 34, 34]  # № · ЖК · NPS26 · n26 · вклад · сегмент = 260
    hdr = ["№","ЖК","NPS26","n26","вклад пп","сегмент"]

    # первая колонка
    rows1 = []
    for i,z in enumerate(ranked[:n_per_col], 1):
        rows1.append([str(i), z["zhk"], z["nps26"], z["n26"], z["vklad"], z["seg"]])
    table_compact(c, MARGIN, ty-14, col_w, hdr, rows1,
                  header_fill=VIOLET_DARK, row_h=7.2, font_size=5.6,
                  nps_col=2, contrib_col=4, seg_col=5)

    # вторая колонка
    x2 = MARGIN + sum(col_w) + 16
    rows2 = []
    for i,z in enumerate(ranked[n_per_col:], n_per_col+1):
        rows2.append([str(i), z["zhk"], z["nps26"], z["n26"], z["vklad"], z["seg"]])
    table_compact(c, x2, ty-14, col_w, hdr, rows2,
                  header_fill=VIOLET_DARK, row_h=7.2, font_size=5.6,
                  nps_col=2, contrib_col=4, seg_col=5)

    # Правая колонка — сценарии, чтобы дойти до цели
    x3 = x2 + sum(col_w) + 20
    caption(c, x3, ty-4, "ЧТОБЫ ДОЙТИ ДО ЦЕЛИ", color=VIOLET_DARK, font_size=9)
    cw_scen = [56, 62, 44, 60]
    table_compact(c, x3, ty-14, cw_scen,
        ["Сценарий","Ответов ещё","NPS новых","Реальность"],
        [
          ["KPI  −35","+2 100","≥ −20","высокая"],
          ["Цель −30","+3 800","≥  −5","средняя"],
        ],
        header_fill=ORANGE, row_h=13, font_size=6.5)

    # Блок «Правила прогноза» под сценарием
    box(c, x3, 90, sum(cw_scen), 120, fill=BG_LIGHT, radius=6)
    caption(c, x3+10, 200, "ПРАВИЛА ПРОГНОЗА", color=VIOLET_DARK, font_size=8)
    lines = [
        "N ≥ 30 → 70/30 (факт/история)",
        "5 ≤ N < 30 → 50/50",
        "N < 5 → только история 2024–25",
        "нет данных → общий текущий NPS",
        "итог: взвешено по плановым",
        "ответам каждого ЖК",
    ]
    for i,ln in enumerate(lines):
        caption(c, x3+10, 186-i*14, ln, color=INK, font_size=7)

# ---------- Лист 3 Пульс дня ----------
def body_p3(c):
    y = PAGE_H-80
    h2(c, MARGIN, y, "Одна строка = один день · слоты по MSK")
    y -= 6
    cols = [46, 44, 52, 40, 40, 40, 44, 66, 66, 66, 54, 160]
    rows = [
      ["Дата","n день","NPS кон.","%П","%Н","%К","Δ пп","00–08 n/%","08–16 n/%","16–00 n/%","Топ ч","Коммуникация"],
      ["08.09","380","−58,2","12,1","11,4","76,5","—","22 / 6%","180 / 47%","178 / 47%","20, 12","старт опроса, пуш в ЛКК"],
      ["09.09","512","−55,8","14,2","12,8","73,0","+2,4","30 / 6%","240 / 47%","242 / 47%","19, 21","пост в МАХ · 20 ЖК"],
      ["10.09","435","−53,5","16,7","13,1","70,2","+2,3","25 / 6%","210 / 48%","200 / 46%","14, 20","CRM по лояльникам · 12 ЖК"],
      ["11.09","—","—","—","—","—","—","—","—","—","—","(ожидается)"],
      ["12.09","—","—","—","—","—","—","—","—","—","—","(ожидается)"],
    ]
    y = table(c, MARGIN, y, cols, rows, header_fill=VIOLET_DARK, row_h=20, highlight_col=2)
    # Пояснение
    y -= 20
    box(c, MARGIN, y-70, CONTENT_W, 70, fill=BG_LIGHT, radius=8)
    caption(c, MARGIN+12, y-16, "Столбец «Коммуникация»", color=INK, font_size=10)
    caption(c, MARGIN+12, y-32, "Заполняется по команде из чата: «добавь в колонку коммуникация — ...».")
    caption(c, MARGIN+12, y-46, "Без даты — запись падает в сегодняшнюю строку. С датой — в указанную.")
    caption(c, MARGIN+12, y-60, "«Топ ч» — самые активные часы дня по числу ответов (топ-2).")

# ---------- Лист 4 Карта ЖК ----------
def body_p4(c):
    y = PAGE_H-80
    # Фильтры
    def chip(x,y,label,w=90,fill=VIOLET_LIGHT):
        box(c,x,y,w,18,fill=fill,radius=9)
        c.setFillColor(INK); c.setFont("Graphik-Md",8); c.drawString(x+10,y+5,label)
    chip(MARGIN, y-14, "Регион ▼")
    chip(MARGIN+100, y-14, "Мероприятие ▼", w=110)
    chip(MARGIN+220, y-14, "Оси: NPS × ΔNPS ▼", w=140)

    # ---- BCG-матрица, слева ----
    mx1, my1 = MARGIN, 130
    mx2, my2 = MARGIN + 320, y - 40
    box(c, mx1, my1, mx2-mx1, my2-my1, fill=BG_LIGHT, radius=8)
    ax = mx1 + (mx2-mx1)/2; ay = my1 + (my2-my1)/2
    c.setStrokeColor(GRAY_DARK); c.setDash(2,2)
    c.line(mx1+20, ay, mx2-10, ay); c.line(ax, my1+20, ax, my2-10)
    c.setDash()
    c.setFillColor(GRAY_DARK); c.setFont("Graphik-Sb", 9)
    c.drawString(mx1+12,  my2-22, "Восходящие"); c.drawRightString(mx2-12, my2-22, "Звезды")
    c.drawString(mx1+12,  my1+12, "Кризис");     c.drawRightString(mx2-12, my1+12, "Спящие")
    caption(c, ax+4, my2-14, "NPS →"); caption(c, mx2-42, ay+6, "Δ YoY ↑")

    # Все ЖК с n>=10 — точками, чтобы карта была честной
    def _to_float(s):
        s = s.replace("+","").replace("−","-").replace(",",".")
        try: return float(s)
        except: return None
    NPS_MIN, NPS_MAX = -100, 20
    DY_MIN, DY_MAX = -50, 50
    for z in ZHK_FULL:
        nps26 = _to_float(z["nps26"]); nps25 = _to_float(z["nps25"])
        if nps26 is None: continue
        dy = None if nps25 is None else nps26 - nps25
        if dy is None: continue
        xr = (nps26 - NPS_MIN)/(NPS_MAX-NPS_MIN); xr = max(0, min(1, xr))
        yr = (dy - DY_MIN)/(DY_MAX-DY_MIN);   yr = max(0, min(1, yr))
        px = mx1 + 24 + xr*(mx2-mx1-48); py = my1 + 24 + yr*(my2-my1-48)
        col = _seg_color(z["seg"])
        r = 2 + min(4, z["n26_int"]/8)
        c.setFillColor(col); c.setStrokeColor(colors.white); c.setLineWidth(0.3)
        c.circle(px, py, r, stroke=1, fill=1)

    # Точки с n<50 — прозрачнее (сделаем более тусклым цветом уже через сегмент "Недобор" GRAY_DARK)
    caption(c, mx1+12, my1-12, "Размер точки = n2026 ; серые точки — n<50 «недобор»", color=GRAY_DARK, font_size=7)

    # ---- Все 96 ЖК по сегментам справа ----
    x3 = MARGIN + 340
    caption(c, x3, y-14, "ВСЕ 96 ЖК ПО СЕГМЕНТАМ", color=VIOLET_DARK, font_size=9)
    caption(c, x3, y-24, "по данным 2026 · сортировка внутри сегмента по NPS", color=GRAY_DARK, font_size=7)

    order = {"Звезды":0,"Восходящие":1,"Спящие":2,"Кризис":3,"Недобор":4}
    def _nps_key(z):
        v = z["nps26"].replace("+","").replace("−","-").replace(",",".")
        try: return -float(v)
        except: return 999
    sorted_z = sorted(ZHK_FULL, key=lambda z: (order.get(z["seg"], 5), _nps_key(z)))

    col_w = [14, 118, 30, 22, 30, 32]
    hdr = ["№","ЖК","NPS26","n26","вклад","сегмент"]
    # 2 колонки по 48
    rows1 = [[str(i+1), z["zhk"], z["nps26"], z["n26"], z["vklad"], z["seg"]]
             for i,z in enumerate(sorted_z[:48])]
    rows2 = [[str(i+49), z["zhk"], z["nps26"], z["n26"], z["vklad"], z["seg"]]
             for i,z in enumerate(sorted_z[48:])]
    table_compact(c, x3, y-38, col_w, hdr, rows1,
                  header_fill=VIOLET_DARK, row_h=7.6, font_size=5.8,
                  nps_col=2, contrib_col=4, seg_col=5)
    x4 = x3 + sum(col_w) + 12
    table_compact(c, x4, y-38, col_w, hdr, rows2,
                  header_fill=VIOLET_DARK, row_h=7.6, font_size=5.8,
                  nps_col=2, contrib_col=4, seg_col=5)

# ---------- Лист 5 Витрина ЖК ----------
def body_p5(c):
    y = PAGE_H-80
    # Фильтры — компактные чипы
    def field(x,y,label,w=110):
        box(c,x,y,w,18,fill=colors.white,stroke=GRAY_LIGHT,radius=4)
        c.setFillColor(GRAY_DARK); c.setFont("Graphik",7); c.drawString(x+6,y+6,label)
    gap = 6
    widths = [150, 96, 76, 76, 96, 86, 96]
    scale = (CONTENT_W - gap*(len(widths)-1)) / sum(widths)
    widths = [round(w*scale) for w in widths]
    labels = ["Поиск ЖК: [___________]","NPS от [__] до [__]","n от [__] до [__]",
              "Регион ▼","Мероприятие ▼","Сегмент ▼","Только n≥50 ☐"]
    xf = MARGIN
    for lbl,w_ in zip(labels, widths):
        field(xf, y-18, lbl, w=w_); xf += w_ + gap

    # ---- Полный список 96 ЖК в двух колонках ----
    y_body = y - 30
    # Заголовок с числом
    caption(c, MARGIN, y_body, "ВСЕ 96 ЖК ПЕРИМЕТРА · n26 · NPS 2024/2025/2026 · вклад пп · сегмент", color=VIOLET_DARK, font_size=9)
    y_body -= 8

    # Разделим 96 строк на 2 колонки по 48
    col_w = [16, 108, 32, 32, 32, 26, 34, 36]  # № · ЖК · NPS24 · NPS25 · NPS26 · n26 · вклад · сегмент = 316
    hdr = ["№","ЖК","NPS24","NPS25","NPS26","n26","вклад пп","сегмент"]

    # сортировка по номеру (алфавитный порядок фонда)
    ranked = sorted(ZHK_FULL, key=lambda z: z["num"])
    n_per_col = 48

    rows1=[]
    for z in ranked[:n_per_col]:
        rows1.append([str(z["num"]), z["zhk"], z["nps24"], z["nps25"], z["nps26"],
                      z["n26"], z["vklad"], z["seg"]])
    table_compact(c, MARGIN, y_body, col_w, hdr, rows1,
                  header_fill=VIOLET_DARK, row_h=8, font_size=6.0,
                  nps_col=4, contrib_col=6, seg_col=7)

    x2 = MARGIN + sum(col_w) + 20
    rows2=[]
    for z in ranked[n_per_col:]:
        rows2.append([str(z["num"]), z["zhk"], z["nps24"], z["nps25"], z["nps26"],
                      z["n26"], z["vklad"], z["seg"]])
    table_compact(c, x2, y_body, col_w, hdr, rows2,
                  header_fill=VIOLET_DARK, row_h=8, font_size=6.0,
                  nps_col=4, contrib_col=6, seg_col=7)

    # Правая колонка — легенда сегментов и правило n≥50
    x3 = x2 + sum(col_w) + 16
    caption(c, x3, y_body, "ЛЕГЕНДА СЕГМЕНТА", color=VIOLET_DARK, font_size=8)
    seg_desc = [
        ("Звезды",     GREEN,    "NPS ≥ 0, тренд ≥ 0"),
        ("Восходящие", GREEN,    "Δ ≥ +5 пп к 2025"),
        ("Спящие",     YELLOW,   "стабильно · без импульса"),
        ("Кризис",     RED,      "Δ ≤ −5 пп или NPS < −30"),
        ("Недобор",    GRAY_DARK,"n < 50 · интерпретация под вопросом"),
    ]
    for i,(name, col, desc) in enumerate(seg_desc):
        yy = y_body - 16 - i*22
        c.setFillColor(col); c.circle(x3+6, yy+3, 4, stroke=0, fill=1)
        c.setFillColor(INK); c.setFont("Graphik-Sb",8); c.drawString(x3+16, yy+2, name)
        c.setFillColor(GRAY_DARK); c.setFont("Graphik",7); c.drawString(x3+16, yy-8, desc)

    # Пояснение
    box(c, x3, 60, 130, 110, fill=BG_LIGHT, radius=6)
    caption(c, x3+8, 158, "КАК ЧИТАТЬ", color=VIOLET_DARK, font_size=8)
    notes = [
        "NPS 2024/2025 — база",
        "сопоставимости; фиксируем",
        "и не пересчитываем.",
        "",
        "NPS 2026 накопительный,",
        "обновляется ежечасно.",
        "",
        "Вклад ЖК считается",
        "ТОЛЬКО по данным 2026",
        "(n2026 / общий n × NPS ЖК).",
    ]
    for i,ln in enumerate(notes):
        caption(c, x3+8, 148-i*10, ln, color=INK, font_size=6.5)

# ---------- Лист 6 Драйверы ----------
def body_p6(c):
    y = PAGE_H-80
    h2(c, MARGIN, y, "Waterfall: NPS 2025 → NPS 2026 (взвешено по важности CSI)")
    # waterfall — верх страницы, оставляем полосу для подписей внизу
    wx1, wy1 = MARGIN, 340
    wx2, wy2 = MARGIN + CONTENT_W - 320, y - 30
    box(c, wx1, wy1, wx2-wx1, wy2-wy1, fill=BG_LIGHT, radius=8)
    steps = [
      ("2025",    -32,  VIOLET_DARK, "start"),
      ("Операт.", -4.1, RED,         "delta"),
      ("Лифты",   -3.8, RED,         "delta"),
      ("Уборка",  -2.5, RED,         "delta"),
      ("Инжен.",  -1.9, RED,         "delta"),
      ("Прочее",  -0.5, ORANGE,      "delta"),
      ("2026",    -53.5, VIOLET_DARK,"end"),
    ]
    def yv(v):
        vmin, vmax = -60, -20
        return wy1 + 24 + (v-vmin)/(vmax-vmin)*(wy2-wy1-40)
    inner_w = wx2 - wx1 - 40
    bw = inner_w / len(steps)
    cur = -32
    for i,(name, val, col, kind) in enumerate(steps):
        x = wx1 + 20 + i*bw
        if kind in ("start","end"):
            y_top = yv(val); y_base = yv(-60)
            c.setFillColor(col); c.rect(x+6, y_base, bw-12, y_top-y_base, stroke=0, fill=1)
            c.setFillColor(colors.white); c.setFont("Graphik-Sb",10)
            c.drawCentredString(x+bw/2, y_top-14, f"{val}")
        else:
            y_top = yv(cur); y_bot = yv(cur+val)
            c.setFillColor(col); c.rect(x+6, y_bot, bw-12, y_top-y_bot, stroke=0, fill=1)
            c.setFillColor(INK); c.setFont("Graphik-Sb",8)
            c.drawCentredString(x+bw/2, y_bot-11, f"{val:+.1f}")
            cur += val
        # подпись столбца под ящиком, ниже wy1
        c.setFillColor(GRAY_DARK); c.setFont("Graphik",8)
        c.drawCentredString(x+bw/2, wy1+8, name)

    # Матрица важность × Δ
    mx1 = wx2 + 20; my1 = 340; mx2 = MARGIN + CONTENT_W; my2 = wy2
    box(c, mx1, my1, mx2-mx1, my2-my1, fill=BG_LIGHT, radius=8)
    caption(c, mx1+10, my2-14, "Матрица «Важность × Δ YoY»", color=INK, font_size=10)
    ax = mx1 + (mx2-mx1)/2; ay = my1 + (my2-my1)/2
    c.setStrokeColor(GRAY_DARK); c.setDash(2,2)
    c.line(mx1+20, ay, mx2-10, ay); c.line(ax, my1+20, ax, my2-30)
    c.setDash()
    pts = [
      ("Лифты",       0.62, 0.85, RED),
      ("Операт.",     0.72, 0.90, RED),
      ("Уборка МОП",  0.55, 0.78, RED),
      ("Инжен.",      0.60, 0.68, RED),
      ("Информ.",     0.45, 0.55, ORANGE),
      ("Домофон",     0.40, 0.42, ORANGE),
      ("Соседи",      0.20, 0.25, GREEN),
      ("Мебель",      0.15, 0.18, GREEN),
    ]
    for name, xr, yr, col in pts:
        px = mx1 + 30 + xr*(mx2-mx1-60); py = my1 + 30 + yr*(my2-my1-60)
        c.setFillColor(col); c.circle(px, py, 5, stroke=0, fill=1)
        c.setFillColor(INK); c.setFont("Graphik",7); c.drawString(px+7, py-2, name)
    caption(c, mx1+22, my1+22, "− важ. низ")
    caption(c, mx2-72, my1+22, "важ. выс +")

    # Топ-5 «что чинить» — ниже waterfall, с зазором
    h2(c, MARGIN, 300, "Топ-5 «что чинить»")
    table(c, MARGIN, 290, [200, 60, 60, 80, 80],
      [
        ["Критерий","Важн.","Оценка","ЖК с оц.<3","Приоритет"],
        ["Оперативность заявок","9,1","2,89","67","P0"],
        ["Работа лифтов","8,9","2,90","58","P0"],
        ["Уборка МОП","8,6","2,98","52","P0"],
        ["Инженерные системы","8,4","3,10","44","P1"],
        ["Домофония","7,2","3,20","38","P1"],
      ], header_fill=VIOLET_DARK, row_h=18, highlight_col=4)

# ---------- Лист 7 Разрезы ----------
def body_p7(c):
    y = PAGE_H-80
    # 2 колонки: демография + гео/фонд
    h2(c, MARGIN, y, "Демография")
    table(c, MARGIN, y-4, [70, 50, 60, 100],
      [
        ["Пол","n","NPS","маркер"],
        ["Ж","980","−48",""],
        ["М","817","−60",""],
      ], header_fill=VIOLET_DARK, row_h=16)

    table(c, MARGIN, y-70, [70, 50, 60, 100],
      [
        ["Возраст","n","NPS","маркер"],
        ["< 20","15","−80","гипотеза"],
        ["21–30","340","−45",""],
        ["31–40","690","−52",""],
        ["41–50","430","−58",""],
        ["> 50","322","−67",""],
      ], header_fill=VIOLET_DARK, row_h=16, highlight_col=3)

    # Регион + класс
    h2(c, MARGIN+300, y, "География и фонд")
    table(c, MARGIN+300, y-4, [100, 50, 60],
      [["Регион","n","NPS"],
       ["Москва","1120","−55"],
       ["МО","420","−49"],
       ["СПб","170","−52"],
       ["Прочие","87","−60"],
      ], header_fill=VIOLET_DARK, row_h=16)
    table(c, MARGIN+300, y-90, [100, 50, 60],
      [["Класс фонда","n","NPS"],
       ["Стандарт","920","−62"],
       ["Комфорт","630","−48"],
       ["Бизнес","247","−38"],
      ], header_fill=VIOLET_DARK, row_h=16, highlight_col=2)

    # Гарантия — крупный блок
    h2(c, MARGIN+520, y, "Гарантия · крупнейший драйвер оттока")
    table(c, MARGIN+520, y-4, [140, 50, 60],
      [["Обращались","n","NPS"],
       ["Да","1168","−67"],
       ["Нет","629","−23"],
       ["из «да»: оц 4+5","—","−18"],
       ["из «да»: оц 1–2","—","−85"],
      ], header_fill=ORANGE, row_h=16, highlight_col=2)
    caption(c, MARGIN+520, y-90, "Δ Да − Нет: −44 пп", color=ORANGE, font_size=10)

    # ПИК-Сервис, ПИК-Аренда
    y2 = 180
    h2(c, MARGIN, y2+50, "ПИК-Сервис · awareness · usage · NPS")
    table(c, MARGIN, y2+42, [130, 60, 60],
      [["Показатель","2025","2026"],
       ["Awareness","74%","93%"],
       ["Usage","39%","41%"],
       ["NPS пользователей","","−31"],
       ["NPS непользователей","","−60"],
       ["Δ NPS (польз. − непольз.)","","+29 пп"],
      ], header_fill=VIOLET_DARK, row_h=16, highlight_col=2)

    h2(c, MARGIN+320, y2+50, "ПИК-Аренда · awareness · usage · NPS")
    table(c, MARGIN+320, y2+42, [130, 60, 60],
      [["Показатель","2025","2026"],
       ["Awareness","—","62%"],
       ["Usage","—","10%"],
       ["NPS пользователей","","−38"],
       ["NPS непользователей","","−55"],
       ["Δ NPS (польз. − непольз.)","","+17 пп"],
      ], header_fill=VIOLET_DARK, row_h=16, highlight_col=2)

    # Заметка
    caption(c, MARGIN+640, y2+50, "Малые группы < 30 — маркер «гипотеза»", color=GRAY_DARK)

# ---------- Лист 8 Списки для работы ----------
def body_p8(c):
    y = PAGE_H-80

    # ---- Триптих: 3 колонки — Просели / Недобор / Лояльники ----
    # Заголовок
    caption(c, MARGIN, y, "СПИСКИ ДЛЯ РАБОТЫ · все ЖК по трём разрезам одновременно",
            color=VIOLET_DARK, font_size=10)
    caption(c, MARGIN, y-11,
            "Просели — суточное падение NPS ≥ 5 пп. Недобор — n26 < 50. Лояльники — % ответов ≥ 7 из всех респондентов ЖК.",
            color=GRAY_DARK, font_size=7)

    col_gap = 10
    col_w_total = (CONTENT_W - col_gap*2) / 3

    # Проседания — фиктивно на основе тех, у кого NPS26 < NPS25 и n26 достаточен
    def _to_float(s):
        s = s.replace("+","").replace("−","-").replace(",",".")
        try: return float(s)
        except: return None
    proseli = []
    for z in ZHK_FULL:
        n26 = _to_float(z["nps26"]); n25 = _to_float(z["nps25"])
        if n26 is None or n25 is None: continue
        d = n26 - n25
        if d <= -5 and z["n26_int"] >= 10:
            proseli.append((z, d))
    proseli.sort(key=lambda t: t[1])  # сильнее просели вверху

    # Недобор — все с n26 < 50
    nedobor = sorted([z for z in ZHK_FULL if z["n26_int"] < 50], key=lambda z: z["n26_int"])

    # Лояльники — с положительным NPS 2026 или ростом
    loyal = []
    for z in ZHK_FULL:
        n26 = _to_float(z["nps26"])
        if n26 is None: continue
        if n26 >= 0 or (_to_float(z["nps25"]) is not None and n26 - _to_float(z["nps25"]) >= 5):
            loyal.append(z)

    # заголовки блоков
    hdr_y = y - 26
    def block_header(x, w, title, count, color):
        box(c, x, hdr_y-18, w, 22, fill=color, radius=6)
        c.setFillColor(colors.white); c.setFont("Graphik-Sb", 10)
        c.drawString(x+10, hdr_y-12, title)
        c.setFillColor(colors.white); c.setFont("Graphik-Sb", 10)
        c.drawRightString(x+w-10, hdr_y-12, f"{count}")

    x1 = MARGIN
    x2 = x1 + col_w_total + col_gap
    x3 = x2 + col_w_total + col_gap
    block_header(x1, col_w_total, "ПРОСЕЛИ · Δ ≤ −5 пп",           len(proseli), ORANGE)
    block_header(x2, col_w_total, "НЕДОБОР · n < 50",              len(nedobor), VIOLET_DARK)
    block_header(x3, col_w_total, "ЛОЯЛЬНИКИ · NPS ≥ 0 или Δ ≥ +5",len(loyal),   GREEN)

    # Таблицы
    ty = hdr_y - 24
    cw_prox = [16, 100, 32, 32, 30]
    scale = col_w_total / sum(cw_prox)
    cw = [round(v*scale, 1) for v in cw_prox]

    r_pros = [[str(i+1), z["zhk"], z["nps25"], z["nps26"], z["n26"]]
              for i,(z,_) in enumerate(proseli)]
    table_compact(c, x1, ty, cw, ["№","ЖК","NPS25","NPS26","n26"],
                  r_pros or [["—","(нет)","","",""]],
                  header_fill=ORANGE, row_h=9, font_size=6.2, nps_col=3)

    r_nedo = [[str(i+1), z["zhk"], z["n26"], z["nps26"], z["seg"]]
              for i,z in enumerate(nedobor)]
    table_compact(c, x2, ty, cw, ["№","ЖК","n26","NPS26","сегмент"],
                  r_nedo or [["—","(нет)","","",""]],
                  header_fill=VIOLET_DARK, row_h=5.6, font_size=5.2, seg_col=4, nps_col=3)

    r_loy = [[str(i+1), z["zhk"], z["nps25"], z["nps26"], z["n26"]]
             for i,z in enumerate(loyal)]
    table_compact(c, x3, ty, cw, ["№","ЖК","NPS25","NPS26","n26"],
                  r_loy or [["—","(нет)","","",""]],
                  header_fill=GREEN, row_h=9, font_size=6.2, nps_col=3)

    # Подсказка: что делать — узкая полоса внизу над колонтитулом
    hint_y = 40
    box(c, MARGIN, hint_y-4, CONTENT_W, 30, fill=BG_LIGHT, radius=4)
    caption(c, MARGIN+8, hint_y+18, "ЧТО ДЕЛАТЬ", color=VIOLET_DARK, font_size=7)
    hints = [
        "Просели · CRM критикам, встреча с УК ЖК, разбор причин в течение суток.  ·  Недобор · ЛКК-push + пост в ТГ ЖК; response rate до n ≥ 50; RR > 20% — риск.  ·  Лояльники · CRM (амбассадоры, бонус); нейтралы 7–8 → промоутеры 9–10.",
    ]
    for i,ln in enumerate(hints):
        caption(c, MARGIN+8, hint_y+8-i*8, ln, color=INK, font_size=6.5)

# ---------- Лист 9 Методика ----------
def body_p9(c):
    y = PAGE_H-90
    # Левая колонка: Формулы + Пороги — сужена до 480pt
    LEFT_W = 500
    h2(c, MARGIN, y, "Формулы")
    y -= 4
    table(c, MARGIN, y, [120, LEFT_W-120],
      [["Метрика","Формула"],
       ["NPS","% промоутеров (9–10) − % критиков (0–6); нейтралы (7–8) в знаменателе"],
       ["CSAT","доля (4+5) / общее × 100%    (не среднее)"],
       ["CSI","Σ (оценка × важность) / Σ (важность), нормировано / max × 100%"],
      ], header_fill=VIOLET_DARK, row_h=22, highlight_col=0)

    y -= 110
    h2(c, MARGIN, y, "Пороги и правила")
    y -= 4
    table(c, MARGIN, y, [200, LEFT_W-200],
      [["Правило","Действие"],
       ["n ≥ 50 на ЖК","обязательно для сравнения; ниже — «недостаточно данных»"],
       ["Малые группы < 30","маркер «гипотеза», не выводы"],
       ["Вклад ЖК в общий NPS","считается ТОЛЬКО по данным 2026"],
       ["Прошлый год в сводке","не показываем; только KPI, факт, цель"],
      ], header_fill=VIOLET_DARK, row_h=20, highlight_col=1)

    # Правая колонка — справа от левой с зазором
    x = MARGIN + LEFT_W + 30
    ry = PAGE_H - 90
    h2(c, x, ry, "Источники")
    ry -= 18
    for line in [
        "Testograf survey 712624 (2026)",
        "Testograf survey 662501 (2025)",
        "Testograf survey 599226 (2024)",
        "Справочник фонда: fond-01.09.26.xlsx",
        "Кластеризация: pplx_sdk.llm.extract",
    ]:
        caption(c, x, ry, line, font_size=9, color=INK); ry -= 13

    ry -= 12
    h2(c, x, ry, "Прогноз (плашка на «Листе»)")
    ry -= 18
    for line, col in [
        ("Взвешенное объединение:", INK),
        ("1) экстраполяция дневного темпа 2026;", INK),
        ("2) исторический профиль сентября 2025.", INK),
        ("Взвешивание: 50/50 · 70/30 · 100/0 — согласовать.", ORANGE),
    ]:
        caption(c, x, ry, line, font_size=9, color=col); ry -= 13

    ry -= 12
    h2(c, x, ry, "Лог сборки")
    ry -= 18
    caption(c, x, ry,    "Последнее обновление: 10.09.2026 18:54 MSK", font_size=9, color=INK); ry -= 13
    caption(c, x, ry,    "Статус Testograf API: 429 (лимит исчерпан)", font_size=9, color=RED)

# ---------- Лист 10 Голос жителя ----------
def body_p10(c):
    y = PAGE_H-80
    h2(c, MARGIN, y, "Уровень 1 · три сегмента респондентов")
    caption(c, MARGIN, y-13, "клик по плитке → список ЖК → голос жителя", color=GRAY_DARK, font_size=8)

    # три плитки с процентами (шапка листа)
    tw = 180; th = 60; ty = y-24-th
    kpi_tile(c, MARGIN,           ty, tw, th, "промоутеры (9–10)", "16,7%", "P=300", tile_color=GREEN, value_size=26)
    kpi_tile(c, MARGIN+tw+16,     ty, tw, th, "нейтралы (7–8)",    "13,1%", "N=236", tile_color=YELLOW, value_color=INK, value_size=26)
    kpi_tile(c, MARGIN+2*(tw+16), ty, tw, th, "критики (0–6)",     "70,2%", "D=1 261",  tile_color=RED, value_size=26)

    # Правая справка о механике
    xr = MARGIN + 3*(tw+16) - 16 + 20
    caption(c, xr, y-16, "УРОВЕНЬ 2 → все 96 ЖК", color=VIOLET_DARK, font_size=8)
    caption(c, xr, y-26, "% критиков внутри ЖК", color=GRAY_DARK, font_size=7)
    caption(c, xr, y-36, "сортировка: тяжелейшие критики вверху", color=GRAY_DARK, font_size=7)

    # ---- Все 96 ЖК с процентом критиков ----
    y2 = ty - 14
    h2(c, MARGIN, y2, "Уровень 2 · все 96 ЖК · % критиков (0–6) от респондентов ЖК")
    y2 -= 8

    # Считаем % критиков из responses_long.csv
    import csv
    from collections import defaultdict
    total = defaultdict(int); crit = defaultdict(int)
    try:
        with open("/home/user/workspace/nps_pipeline/db/responses_long.csv", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                if row.get("year") != "2026": continue
                zhk = (row.get("zhk") or "").strip()
                nps = row.get("nps_uk")
                if not zhk or not nps: continue
                try: v = int(float(nps))
                except: continue
                total[zhk] += 1
                if v <= 6: crit[zhk] += 1
    except Exception as e:
        pass

    rows = []
    for z in ZHK_FULL:
        t = total.get(z["zhk"], 0)
        c_ = crit.get(z["zhk"], 0)
        pct = f"{c_*100/t:.0f}%".replace(".",",") if t else "—"
        marker = "n<50" if t < 50 else ""
        rows.append((z["zhk"], pct, c_, t, marker, c_*100/t if t else -1))

    rows.sort(key=lambda r: -r[5])  # по % критиков сверху
    # 3 колонки по 32 строки = 96
    n_per_col = 32
    col_w = [14, 110, 30, 24, 22, 30]
    hdr = ["№","ЖК","% крит","крит.","n26","маркер"]
    for k in range(3):
        chunk = rows[k*n_per_col : (k+1)*n_per_col]
        rr = []
        for i,(zhk, pct, c_, t, marker, _) in enumerate(chunk):
            rr.append([str(k*n_per_col+i+1), zhk, pct, str(c_), str(t), marker])
        x = MARGIN + k*(sum(col_w)+16)
        table_compact(c, x, y2, col_w, hdr, rr,
                      header_fill=RED, row_h=9, font_size=6.0)

    # Уровень 3 — компактный пример под таблицей
    y3 = y2 - (n_per_col+1)*9 - 22
    h2(c, MARGIN, y3, "Уровень 3 · пример: «Бутово парк 2» × критики · 65 из 78 (84%)")
    y3 -= 8
    lw = 420; rw = CONTENT_W - lw - 20
    box(c, MARGIN, y3-70, lw, 70, fill=BG_LIGHT, radius=6)
    caption(c, MARGIN+10, y3-12, "Свободные ответы (65)", color=INK, font_size=8)
    quotes = [
        "«На заявки не реагируют месяцами. Уборка на этажах ужасная...»",
        "«Только замена персонала, который только обещает и не делает...»",
        "«Работа только на бумаге. Фотографируют, но не делают.»",
    ]
    for i,q in enumerate(quotes):
        caption(c, MARGIN+10, y3-24-i*12, q, color=INK, font_size=7)

    rx = MARGIN + lw + 20
    box(c, rx, y3-70, rw, 70, fill=BG_LIGHT, radius=6)
    caption(c, rx+10, y3-12, "Кластеры проблем", color=INK, font_size=8)
    clusters = [
        ("Качество исполнения работ", 18, RED),
        ("Оперативность заявок",      15, RED),
        ("Работа лифтов",             12, ORANGE),
        ("Чистота МОП",               11, ORANGE),
        ("Инженерные системы",         9, VIOLET_DARK),
    ]
    max_v = clusters[0][1]; bar_max = rw - 220
    for i,(name, v, col) in enumerate(clusters):
        yy = y3 - 24 - i*10
        c.setFillColor(INK); c.setFont("Graphik", 7); c.drawString(rx+10, yy, name)
        w = bar_max * v/max_v
        c.setFillColor(col); c.rect(rx+150, yy-1, w, 6, stroke=0, fill=1)
        c.setFillColor(GRAY_DARK); c.setFont("Graphik-Sb", 7)
        c.drawString(rx+150+w+4, yy, str(v))

# =================================================
def build():
    import os
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    c = canvas.Canvas(OUT, pagesize=landscape(A4))
    # cover
    page_00_cover(c)
    # 10 листов
    page_content(c, 1, "Лист · сводка компании", "витрина верхнего уровня", body_p1)
    page_content(c, 2, "Прогноз NPS 2026",       "три опорные точки: −32 (2025), −35 (KPI), −30 (цель)", body_p2)
    page_content(c, 3, "Пульс дня",              "динамика по дням · слоты · коммуникации", body_p3)
    page_content(c, 4, "ЖК: карта и сегменты",   "BCG-матрица · стратегии по сегментам", body_p4)
    page_content(c, 5, "ЖК: витрина",            "список с фильтрами · разворачиваемые под-строки", body_p5)
    page_content(c, 6, "Драйверы NPS · CSI-факторы YoY", "waterfall · матрица важность × Δ · топ «что чинить»", body_p6)
    page_content(c, 7, "Разрезы: факторы NPS",   "демография · фонд · гарантия · ПИК-Сервис · ПИК-Аренда", body_p7)
    page_content(c, 8, "Списки для работы",      "просели · недобор · критичные CSI · лояльники в CRM", body_p8)
    page_content(c, 9, "Методика",               "формулы · пороги · источники · прогноз", body_p9)
    page_content(c, 10, "Голос жителя",          "три сегмента · drill-down до открытых ответов", body_p10)
    c.save()
    print("Готово:", OUT)

if __name__ == "__main__":
    build()
