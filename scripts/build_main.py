# -*- coding: utf-8 -*-
import sys
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import (XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION,
                             XL_TICK_MARK, XL_MARKER_STYLE)
from pptx.oxml.ns import qn
from build_lib import *

SRC = "source/KazBioFert_Board_Presentation_20260709_ORIGINAL.pptx"
OUT = "KazBioFert_Board_Presentation_FINAL_20260709.pptx"

prs = Presentation(SRC)

# ---------- chart styling helpers ----------
def _chart_font(chart, sz=8, color=INK):
    chart.font.size = Pt(sz); chart.font.name = 'Calibri'; chart.font.color.rgb = color

def style_cat(chart, sz=8.5, color=INK):
    ax = chart.category_axis
    ax.tick_labels.font.size = Pt(sz); ax.tick_labels.font.name='Calibri'
    ax.tick_labels.font.color.rgb = color
    ax.has_major_gridlines = False
    ax.major_tick_mark = XL_TICK_MARK.NONE
    ax.format.line.color.rgb = BORDER; ax.format.line.width = Pt(0.75)

def style_val(chart, sz=8, fmt='0', gridcolor=MIST, mx=None, mn=None):
    ax = chart.value_axis
    ax.tick_labels.font.size = Pt(sz); ax.tick_labels.font.name='Calibri'
    ax.tick_labels.font.color.rgb = GRAY
    ax.tick_labels.number_format = fmt; ax.tick_labels.number_format_is_linked = False
    ax.has_major_gridlines = True
    ax.major_gridlines.format.line.color.rgb = gridcolor
    ax.major_gridlines.format.line.width = Pt(0.75)
    ax.format.line.fill.background()
    ax.major_tick_mark = XL_TICK_MARK.NONE
    if mx is not None: ax.maximum_scale = mx
    if mn is not None: ax.minimum_scale = mn

def legend(chart, sz=8.5):
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.legend.font.size = Pt(sz); chart.legend.font.name='Calibri'
    chart.legend.font.color.rgb = INK

def col_labels(plot, sz=9, color=INK, fmt='0', pos=XL_LABEL_POSITION.OUTSIDE_END):
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.number_format = fmt; dl.number_format_is_linked = False
    dl.font.size = Pt(sz); dl.font.bold = True; dl.font.name='Calibri'; dl.font.color.rgb = color
    try: dl.position = pos
    except Exception: pass

def add_chart(slide, ctype, x, y, w, h, cats, series):
    cd = CategoryChartData(); cd.categories = cats
    for nm, vals in series: cd.add_series(nm, vals)
    gf = slide.shapes.add_chart(ctype, Pt(x), Pt(y), Pt(w), Pt(h), cd)
    ch = gf.chart; ch.has_title = False; _chart_font(ch)
    return ch

def header(s, kicker, headline, sub, kcolor=COPPER, on_dark=False):
    tcol = WHITE if on_dark else NAVY
    scol = LBLUE if on_dark else RGBColor(0x53,0x62,0x72)
    kcol = TAN if on_dark else kcolor
    simple(s, 40, 26, 641, 16, kicker, 11, True, kcol, 'Calibri', spc=1.4)
    simple(s, 40, 46, 648, 28, headline, 21, True, tcol, 'Arial')
    simple(s, 40, 73, 648, 18, sub, 10.5, False, scol, 'Calibri')
    hrule(s, 97, 40, 641, (RGBColor(0x2A,0x3D,0x5C) if on_dark else BORDER), 0.75)

def insight(s, prefix, body, y=304):
    rect(s, 40, y, 640, 44, fill=NAVY, round_=0.06)
    text(s, 56, y, 610, 44, [[
        {'t':prefix+'  ','sz':9.5,'b':True,'c':TAN,'f':'Calibri'},
        {'t':body,'sz':9.5,'b':False,'c':WHITE,'f':'Calibri'}]],
        anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.05)

# ============================================================ DIVIDER
def slide_divider():
    s = add_blank(prs, NAVY)
    rect(s, 470, 0, 250, 405, fill=NAVYDK)
    line_seg(s, 470, 0, 0, 405, RGBColor(0x24,0x3A,0x59), 1.0)
    rect(s, 56, 92, 30, 3, fill=TAN)
    simple(s, 56, 100, 500, 16, 'MARKET INTELLIGENCE  ·  РАЗДЕЛ', 11, True, TAN, 'Calibri', spc=2.4)
    simple(s, 54, 126, 520, 46, 'АНАЛИЗ РЫНКА', 40, True, WHITE, 'Arial')
    simple(s, 54, 170, 520, 46, 'УДОБРЕНИЙ', 40, True, WHITE, 'Arial')
    simple(s, 56, 226, 400, 20,
           'Мировой рынок · Казахстан · Ценовое позиционирование ОМУ · Инвестиционный тезис',
           11.5, False, LBLUE, 'Calibri', line_spacing=1.12)
    rect(s, 56, 262, 236, 22, fill=NAVYDK, line=RGBColor(0x35,0x52,0x74), line_w=1.0, round_=0.5)
    simple(s, 56, 263, 236, 20, 'Данные актуализированы на 30.06.2026', 9, True, TAN,
           'Calibri', PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    kpi_chip(s, 56, 302, 196, 56, '208', 'млн т', 'мировой спрос на NPK (пит. в-ва), 2025',
             big_c=TAN, on_dark=True, big_sz=24)
    kpi_chip(s, 262, 302, 196, 56, '1,3', 'млн т', 'ёмкость рынка удобрений РК, 2025 (+19% г/г)',
             big_c=LBLUE, on_dark=True, big_sz=24)
    kpi_chip(s, 468, 302, 212, 56, '≈52%', '', 'валовая маржа ОМУ · $233 себест. → $487 цена',
             big_c=WHITE, on_dark=True, big_sz=24)
    footer(s, 15, 'АНАЛИЗ РЫНКА · РАЗДЕЛ', on_dark=True)
    return s

# ============================================================ WORLD
def slide_world():
    s = add_blank(prs, WHITE)
    header(s, 'МИРОВОЙ РЫНОК УДОБРЕНИЙ  ·  IFA / WORLD BANK, 2026',
           'Мировое предложение удобрений превышает спрос',
           'Спрос ≈ 208 млн т питательных веществ (N+P+K) в 2025; рост 1–2%/год до ~224 млн т к 2029')
    rect(s, 40, 106, 392, 186, fill=WHITE, line=BORDER, round_=0.05, shadow=True)
    card_head(s, 54, 116, 360, 'Мировой баланс, млн т питательных веществ (N+P₂O₅+K₂O)')
    cats = ['2019','2020','2021','2022','2023','2024','2025','2026п']
    prod = [200, 202, 205, 198, 208, 212, 216, 219]
    cons = [190, 193, 195, 185, 198, 204, 208, 211]
    ch = add_chart(s, XL_CHART_TYPE.LINE_MARKERS, 46, 134, 380, 150, cats,
                   [('Производство', prod), ('Потребление', cons)])
    s0, s1 = ch.series[0], ch.series[1]
    s0.smooth = False; s1.smooth = False
    s0.format.line.color.rgb = STEEL; s0.format.line.width = Pt(2.5)
    s1.format.line.color.rgb = COPPER; s1.format.line.width = Pt(2.5)
    for ser, col in ((s0, STEEL), (s1, COPPER)):
        ser.marker.style = XL_MARKER_STYLE.CIRCLE; ser.marker.size = 5
        ser.marker.format.fill.solid(); ser.marker.format.fill.fore_color.rgb = col
        ser.marker.format.line.color.rgb = WHITE; ser.marker.format.line.width = Pt(1)
    style_cat(ch, 7.5); style_val(ch, 7.5, mn=160, mx=230); legend(ch, 8.5)
    rect(s, 442, 106, 238, 86, fill=WHITE, line=BORDER, round_=0.06, shadow=True)
    card_head(s, 456, 115, 210, 'Структура мирового спроса, 2025')
    bx, by, bw, bh = 456, 137, 210, 20
    cx = bx
    for lab, pct, col in [('N',56,STEEL),('P',23,TAN),('K',21,GRAY)]:
        w = bw*pct/100.0
        rect(s, cx, by, w, bh, fill=col)
        simple(s, cx, by, w, bh, '%d%%'%pct, 8.5, True, WHITE, 'Calibri', PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
        cx += w
    text(s, 456, 164, 214, 20, [[
        {'t':'■ ','sz':9,'b':True,'c':STEEL,'f':'Calibri'},{'t':'Азот 116   ','sz':8.3,'c':INK,'f':'Calibri'},
        {'t':'■ ','sz':9,'b':True,'c':TAN,'f':'Calibri'},{'t':'Фосфор 48   ','sz':8.3,'c':INK,'f':'Calibri'},
        {'t':'■ ','sz':9,'b':True,'c':GRAY,'f':'Calibri'},{'t':'Калий 44','sz':8.3,'c':INK,'f':'Calibri'}]])
    kpi_chip(s, 442, 198, 238, 44, '204→208', 'млн т', 'мировое потребление, 2024→2025 (рекорд)', big_c=STEEL, big_sz=20)
    kpi_chip(s, 442, 247, 238, 44, '+1–2%', '/год', 'прогноз роста спроса до 224 млн т к 2029 (IFA)', big_c=COPPER, big_sz=22)
    insight(s, 'ЦЕНЫ 06.2026:',
            'карбамид ≈ $700/т, DAP ≈ $720/т (+60–80% с начала года) — рост от перебоев через Ормузский пролив, '
            'а не от дефицита; локальное сырьё защищает маржу ОМУ. Ослабление ожидается в 2027 (World Bank).')
    source_note(s, 'Источники: IFA Medium-Term Fertilizer Outlook 2025–2029; World Bank Commodity Markets, 06.2026. Основа: питательные вещества (N+P₂O₅+K₂O).')
    footer(s, 16, 'АНАЛИЗ РЫНКА · МИРОВОЙ РЫНОК')
    return s

# ============================================================ KAZAKHSTAN
def slide_kz():
    s = add_blank(prs, WHITE)
    header(s, 'РЫНОК КАЗАХСТАНА  ·  БНС РК / ENERGYPROM, 2025',
           'Казахстан — импортозависимый и недоудобренный рынок',
           'Ёмкость ≈ 1,3 млн т в 2025 (+19% г/г); внесение 4–5 кг/га против агрономической нормы 80–100 кг/га')
    rect(s, 40, 106, 392, 186, fill=WHITE, line=BORDER, round_=0.05, shadow=True)
    card_head(s, 54, 115, 360, 'Внесение удобрений, кг действ. в-ва на га пашни')
    simple(s, 250, 116, 176, 12, 'агронорма 80–100 кг/га', 8, True, COPPER, 'Calibri', PP_ALIGN.RIGHT)
    cats = ['Казахстан','Афганистан','Узбекистан','Бразилия']
    vals = [5, 7, 150, 344]
    ch = add_chart(s, XL_CHART_TYPE.COLUMN_CLUSTERED, 46, 138, 380, 146, cats, [('кг/га', vals)])
    plot = ch.plots[0]; plot.gap_width = 55; ser = ch.series[0]
    for i, col in enumerate([COPPER, STEEL, STEEL, STEEL]):
        pt = ser.points[i]; pt.format.fill.solid(); pt.format.fill.fore_color.rgb = col
    col_labels(plot, 8.5, INK); ch.has_legend = False
    style_cat(ch, 8); style_val(ch, 7.5, mx=380)
    kpi_chip(s, 442, 106, 238, 44, '1,3', 'млн т', 'ёмкость рынка РК, 2025 (+19% г/г)', big_c=STEEL, big_sz=22)
    kpi_chip(s, 442, 154, 238, 44, '43,6%', '', 'самообеспеченность по азоту (фосфор — 98,5%)', big_c=COPPER, big_sz=22)
    kpi_chip(s, 442, 202, 238, 44, '≈ 6%', '', 'доля завода (80 тыс т ОМУ) в потреблении РК', big_c=STEEL, big_sz=22)
    rect(s, 442, 250, 238, 42, fill=PANEL, line=BORDER, round_=0.10)
    simple(s, 454, 255, 214, 11, 'СТРУКТУРА ИМПОРТА (стоимость)', 7.4, True, GRAY, 'Calibri', spc=0.5)
    text(s, 454, 270, 214, 20, [[
        {'t':'Россия ','sz':10,'b':True,'c':NAVY,'f':'Calibri'},{'t':'73%','sz':10,'b':True,'c':STEEL,'f':'Calibri'},
        {'t':'    ·    Узбекистан ','sz':10,'b':False,'c':INK,'f':'Calibri'},{'t':'11%','sz':10,'b':True,'c':STEEL,'f':'Calibri'}]])
    insight(s, 'ИМПОРТОЗАМЕЩЕНИЕ:',
            'внесение в 15–20× ниже нормы и рост импорта (552 → 700 тыс т за 2023–2025) формируют структурный, '
            'растущий спрос на доступные ОМУ; завод (80 тыс т) закрывает ≈ 11% импорта РК.')
    source_note(s, 'Источники: Бюро нацстатистики РК; EnergyProm / ElDala, 2025; FAO, World Bank (внесение). Показатели — товарный тоннаж.')
    footer(s, 17, 'АНАЛИЗ РЫНКА · КАЗАХСТАН')
    return s

# ============================================================ PRICING & COST
def slide_price():
    s = add_blank(prs, WHITE)
    header(s, 'ЦЕНОВОЕ ПОЗИЦИОНИРОВАНИЕ  ·  ЭКОНОМИКА ОМУ',
           'Цена ОМУ $487/т ниже импортного паритета $630/т',
           'Цена реализации по фин. модели $487/т; себестоимость $233/т (сырьё $194 + производство $39); валовая маржа ≈ 52%')
    rect(s, 40, 106, 232, 186, fill=WHITE, line=BORDER, round_=0.05, shadow=True)
    card_head(s, 54, 115, 200, 'Экономика 1 т ОМУ, $/т')
    yB = 270; topY = 150; maxv = 630.0; scale = (yB - topY) / maxv
    for bx, v, col, lab in [(62,233,GREENT,'Себест.'),(120,487,STEEL,'Цена'),(178,630,NAVY,'Импорт')]:
        h = v*scale; y = yB - h
        rect(s, bx, y, 42, h, fill=col, round_=0.04)
        simple(s, bx-6, y-16, 54, 14, '$%d'%v, 9.5, True, col, 'Calibri', PP_ALIGN.CENTER)
        simple(s, bx-6, yB+3, 54, 12, lab, 7.8, False, GRAY, 'Calibri', PP_ALIGN.CENTER)
    line_seg(s, 54, yB, 200, 0, BORDER, 0.75)
    rect(s, 280, 106, 150, 186, fill=NAVY, round_=0.05, shadow=True)
    simple(s, 292, 130, 126, 34, '≈52%', 30, True, TAN, 'Arial', PP_ALIGN.LEFT)
    simple(s, 293, 166, 126, 14, 'валовая маржа', 10, True, WHITE, 'Calibri')
    simple(s, 293, 182, 126, 12, '$254 на тонну ОМУ', 8.5, False, LBLUE, 'Calibri')
    line_seg(s, 293, 208, 118, 0, RGBColor(0x35,0x52,0x74), 0.75)
    simple(s, 292, 220, 126, 26, '−23%', 22, True, TAN, 'Arial')
    simple(s, 293, 250, 126, 24, 'к цене импортного\nкомплексного аналога', 8.5, False, LBLUE, 'Calibri', line_spacing=1.0)
    rect(s, 440, 106, 240, 186, fill=WHITE, line=BORDER, round_=0.05, shadow=True)
    simple(s, 452, 114, 220, 12, 'ЦЕНЫ КОНКУРЕНТОВ, $/т', 8.4, True, GRAY, 'Calibri', spc=0.5)
    rows = [('Продукт','NPK','Kusto','Alem'),
            ('Карбамид','N 56','441','467'),
            ('Аммофос','P 52','475','823'),
            ('Сульфоаммофос','20-20-14S','492','584'),
            ('Аммиачная селитра','N 34','303','343')]
    tb = s.shapes.add_table(len(rows), 4, Pt(450), Pt(132), Pt(220), Pt(150)).table
    tb.first_row = False; tb.horz_banding = False; clean_table(tb)
    for j, wv in enumerate([92, 54, 40, 34]): tb.columns[j].width = Pt(wv)
    for i, row in enumerate(rows):
        tb.rows[i].height = Pt(29)
        for j, val in enumerate(row):
            c = tb.cell(i, j)
            c.margin_left = Pt(4); c.margin_right = Pt(3); c.margin_top = Pt(0); c.margin_bottom = Pt(0)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            if i == 0:
                c.fill.solid(); c.fill.fore_color.rgb = NAVY; col = WHITE; b = True; sz = 8.3
            else:
                c.fill.solid(); c.fill.fore_color.rgb = WHITE if i % 2 else PANEL; col = INK; b = (j == 0); sz = 8.4
            p = c.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if j < 2 else PP_ALIGN.RIGHT
            r = p.add_run(); r.text = val
            r.font.size = Pt(sz); r.font.bold = b; r.font.name = 'Calibri'; r.font.color.rgb = col
            if i > 0 and j >= 2: r.font.color.rgb = STEEL; r.font.bold = True
            cell_border(c, ('bottom',), (RGBColor(0x35,0x52,0x74) if i==0 else BORDER), 0.75)
    insight(s, 'ПРЕИМУЩЕСТВО:',
            'ОМУ на ≈ 23% дешевле импортных комплексных удобрений (~$630/т) при эффективности питательных веществ '
            '1:3 (ФАО) и локальном сырье — устойчивая юнит-экономика независимо от глобальных цен.')
    source_note(s, 'Цена реализации — по фин. модели FEM v11.5 ($487/т, к выходу ≈57% валовой маржи); рыночная средняя ≈$450/т; импорт ~$630/т (World Bank, 06.2026); эффективность 1:3 (ФАО).')
    footer(s, 18, 'АНАЛИЗ РЫНКА · ЦЕНЫ И ЭКОНОМИКА')
    return s

# ============================================================ INVESTMENT THESIS
def slide_thesis():
    s = add_blank(prs, NAVY)
    header(s, 'ИНВЕСТИЦИОННЫЙ ТЕЗИС  ·  MARKET → PROJECT',
           'Рынок конвертируется в защищённую выручку проекта',
           'Импортозамещение + ценовое преимущество + локальное сырьё = устойчивая маржинальная выручка',
           on_dark=True)
    W = 151
    xs = [40, 40+W+12, 40+2*(W+12), 40+3*(W+12)]
    kpis = [('1,3','млн т','ёмкость рынка РК — адресный спрос',TAN),
            ('≈11%','','доля завода в импорте РК (80 тыс т)',LBLUE),
            ('≈24','млрд ₸','пиковая выручка (~$36 млн) — фин. модель',WHITE),
            ('≈57%','','валовая маржа на выходе (фин. модель)',TAN)]
    for x,(big,unit,lab,col) in zip(xs,kpis):
        kpi_chip(s, x, 108, W, 64, big, unit, lab, big_c=col, on_dark=True, big_sz=23)
    rect(s, 40, 186, 640, 92, fill=NAVYDK, line=RGBColor(0x24,0x3A,0x59), line_w=1.0, round_=0.04)
    bullets = [
        'Импортозамещение — государственный приоритет: спрос структурный и растущий, не привязан к сырьевому суперциклу.',
        'Ценовое преимущество: ОМУ на ≈ 23% дешевле импортных комплексных удобрений (~$630/т) при эффективности 1:3 (ФАО).',
        'Локальная себестоимость ($233/т) и короткое логистическое плечо ($25/т) защищают маржу от волатильности мировых цен.']
    by = 196
    for b in bullets:
        rect(s, 56, by+4, 5, 5, fill=TAN)
        simple(s, 70, by, 596, 22, b, 9.7, False, RGBColor(0xD6,0xE0,0xEC), 'Calibri', line_spacing=1.02)
        by += 27
    rect(s, 40, 290, 640, 46, fill=TAN, round_=0.06)
    text(s, 56, 290, 610, 46, [[
        {'t':'ПРОЕКТНАЯ ЭКОНОМИКА:  ','sz':10,'b':True,'c':NAVY,'f':'Calibri'},
        {'t':'IRR 31%  ·  NPV ≈ 30,8 млрд ₸  ·  окупаемость ≈ 5 лет  —  подтверждается рыночной позицией (цена ОМУ $450/т).',
         'sz':10,'b':True,'c':RGBColor(0x3A,0x2A,0x12),'f':'Calibri'}]], anchor=MSO_ANCHOR.MIDDLE)
    source_note(s, 'Фин. модель FEM v11.5 (4,95 МВт · 80 тыс т ОМУ · цена $487/т · 80/20): IRR проекта 31,08%, NPV 30,77 млрд ₸, WACC 12,71%, окупаемость 5,1 г.', on_dark=True)
    footer(s, 19, 'АНАЛИЗ РЫНКА · ИНВЕСТИЦИОННЫЙ ТЕЗИС', on_dark=True)
    return s

# ---- build & insert ----
n0 = len(prs.slides._sldIdLst)
slide_divider(); slide_world(); slide_kz(); slide_price(); slide_thesis()
for k in range(5):
    move_slide(prs, n0 + k, 14 + k)

# ---- align existing slides to the financial model (FEM v11.5) ----
for sl in prs.slides:
    for sh in sl.shapes:
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if '75 000' in r.text:            # product capacity -> model 80 000 t/y
                        r.text = r.text.replace('75 000', '80 000')

# ---- renumber footer page numbers ----
for idx, sl in enumerate(prs.slides):
    for sh in sl.shapes:
        if sh.has_text_frame:
            t = sh.text_frame.text.strip()
            try: tp = Emu(sh.top).pt; lf = Emu(sh.left).pt
            except Exception: continue
            if t.isdigit() and tp > 372 and lf > 640:
                sh.text_frame.paragraphs[0].runs[0].text = str(idx + 1)

prs.save(OUT)
print("saved", OUT, "slides:", len(prs.slides._sldIdLst))
