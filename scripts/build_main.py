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
    cd = CategoryChartData()
    cd.categories = cats
    for nm, vals in series:
        cd.add_series(nm, vals)
    gf = slide.shapes.add_chart(ctype, Pt(x), Pt(y), Pt(w), Pt(h), cd)
    ch = gf.chart
    ch.has_title = False
    _chart_font(ch)
    return ch

# ============================================================
# SLIDE A — SECTION DIVIDER
# ============================================================
def slide_divider():
    s = add_blank(prs, NAVY)
    # faint deep panel on right for depth
    rect(s, 470, 0, 250, 405, fill=NAVYDK)
    line_seg(s, 470, 0, 0, 405, RGBColor(0x24,0x3A,0x59), 1.0)
    # eyebrow
    simple(s, 56, 92, 500, 16, 'MARKET INTELLIGENCE', 11, True, TAN, 'Calibri', spc=2.6)
    # title
    simple(s, 54, 116, 520, 46, 'АНАЛИЗ РЫНКА', 40, True, WHITE, 'Arial')
    simple(s, 54, 160, 520, 46, 'УДОБРЕНИЙ', 40, True, WHITE, 'Arial')
    line_seg(s, 58, 214, 132, 0, TAN, 2.5)
    simple(s, 56, 228, 470, 34,
           'Мировой рынок · Рынок Казахстана · Ценовое позиционирование и себестоимость ОМУ',
           12.5, False, LBLUE, 'Calibri', line_spacing=1.15)
    # date pill
    p = rect(s, 56, 268, 232, 22, fill=NAVYDK, line=RGBColor(0x35,0x52,0x74), line_w=1.0, round_=0.5)
    simple(s, 56, 269, 232, 20, 'Данные актуализированы на 30.06.2026', 9, True, TAN,
           'Calibri', PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    # KPI row
    kpi_chip(s, 56, 306, 196, 54, '208', 'млн т', 'мировой спрос на NPK (пит. в-ва), 2025',
             big_c=TAN, on_dark=True, big_sz=25)
    kpi_chip(s, 262, 306, 196, 54, '1,3', 'млн т', 'ёмкость рынка удобрений РК, 2025 (+19% г/г)',
             big_c=LBLUE, on_dark=True, big_sz=25)
    kpi_chip(s, 468, 306, 212, 54, '$233', '/т', 'себестоимость ОМУ NPK 8-21 против цены $450',
             big_c=WHITE, on_dark=True, big_sz=25)
    footer(s, 15, 'АНАЛИЗ РЫНКА · РАЗДЕЛ', on_dark=True)
    return s

# ============================================================
# SLIDE B — WORLD MARKET
# ============================================================
def slide_world():
    s = add_blank(prs, WHITE)
    eyebrow(s, 'МИРОВОЙ РЫНОК УДОБРЕНИЙ  ·  IFA / WORLD BANK, 2026', COPPER)
    title(s, 'Мировое предложение удобрений превышает спрос')
    subtitle(s, 'Спрос ≈ 208 млн т питательных веществ (N+P+K) в 2025; рост 1–2%/год до ~224 млн т к 2029 (IFA)')

    # left chart card
    rect(s, 40, 104, 392, 198, fill=WHITE, line=BORDER, round_=0.05, shadow=True)
    simple(s, 54, 114, 360, 16, 'Мировой баланс, млн т питательных веществ (N+P₂O₅+K₂O)',
           9.5, True, NAVY, 'Calibri')
    cats = ['2019','2020','2021','2022','2023','2024','2025','2026п']
    prod = [200, 202, 205, 198, 208, 212, 216, 219]
    cons = [190, 193, 195, 185, 198, 204, 208, 211]
    ch = add_chart(s, XL_CHART_TYPE.LINE_MARKERS, 46, 132, 380, 150, cats,
                   [('Производство', prod), ('Потребление', cons)])
    s0, s1 = ch.series[0], ch.series[1]
    s0.smooth = False; s1.smooth = False
    s0.format.line.color.rgb = STEEL; s0.format.line.width = Pt(2.25)
    s1.format.line.color.rgb = COPPER; s1.format.line.width = Pt(2.25)
    for ser, col in ((s0, STEEL), (s1, COPPER)):
        ser.marker.style = XL_MARKER_STYLE.CIRCLE; ser.marker.size = 5
        ser.marker.format.fill.solid(); ser.marker.format.fill.fore_color.rgb = col
        ser.marker.format.line.color.rgb = WHITE; ser.marker.format.line.width = Pt(1)
    style_cat(ch, 7.5); style_val(ch, 7.5, mn=160, mx=230); legend(ch, 8.5)

    # right: structure card
    rect(s, 442, 104, 238, 96, fill=WHITE, line=BORDER, round_=0.06, shadow=True)
    simple(s, 456, 114, 210, 14, 'Структура мирового спроса, 2025', 9.5, True, NAVY, 'Calibri')
    # segmented 100% bar N56 / P23 / K21
    bx, by, bw, bh = 456, 138, 210, 20
    segs = [('N', 56, STEEL), ('P', 23, TAN), ('K', 21, GRAY)]
    cx = bx
    for lab, pct, col in segs:
        w = bw*pct/100.0
        rect(s, cx, by, w, bh, fill=col)
        simple(s, cx, by, w, bh, '%d%%'%pct, 8.5, True, WHITE, 'Calibri', PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
        cx += w
    text(s, 456, 166, 214, 26, [[
        {'t':'■ ','sz':9,'b':True,'c':STEEL,'f':'Calibri'},{'t':'Азот (N) 116 млн т   ','sz':8.5,'c':INK,'f':'Calibri'},
        {'t':'■ ','sz':9,'b':True,'c':TAN,'f':'Calibri'},{'t':'Фосфор 48   ','sz':8.5,'c':INK,'f':'Calibri'},
        {'t':'■ ','sz':9,'b':True,'c':GRAY,'f':'Calibri'},{'t':'Калий 44','sz':8.5,'c':INK,'f':'Calibri'}]],
        line_spacing=1.1)

    kpi_chip(s, 442, 208, 238, 44, '208', 'млн т', 'мировое потребление пит. в-в, 2025 (+2% г/г)', big_c=STEEL, big_sz=23)
    kpi_chip(s, 442, 256, 238, 44, '+1–2%', '/год', 'прогноз роста спроса до 224 млн т к 2029 (IFA)', big_c=COPPER, big_sz=23)

    # bottom price-shock band
    rect(s, 40, 312, 640, 44, fill=NAVY, round_=0.06)
    text(s, 56, 312, 610, 44, [[
        {'t':'ЦЕНЫ НА 30.06.2026:  ','sz':9.5,'b':True,'c':TAN,'f':'Calibri'},
        {'t':'карбамид ≈ $700/т, аммофос (DAP) ≈ $720/т — рост ~60–80% с начала года на фоне перебоев поставок через Ормузский пролив; ослабление ожидается в 2027 (World Bank).',
         'sz':9.5,'b':False,'c':WHITE,'f':'Calibri'}]], anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.05)
    footer(s, 16, 'АНАЛИЗ РЫНКА · МИРОВОЙ РЫНОК')
    return s

# ============================================================
# SLIDE C — KAZAKHSTAN MARKET
# ============================================================
def slide_kz():
    s = add_blank(prs, WHITE)
    eyebrow(s, 'РЫНОК КАЗАХСТАНА  ·  БНС РК / ENERGYPROM, 2025', COPPER)
    title(s, 'Казахстан — импортозависимый и недоудобренный рынок')
    subtitle(s, 'Ёмкость рынка ≈ 1,3 млн т в 2025 (+19% г/г); внесение 4–5 кг/га против агрономической нормы 80–100 кг/га')

    # left chart card — application rate
    rect(s, 40, 104, 392, 198, fill=WHITE, line=BORDER, round_=0.05, shadow=True)
    simple(s, 54, 114, 360, 16, 'Внесение удобрений, кг действ. в-ва на га пашни', 9.5, True, NAVY, 'Calibri')
    cats = ['Казахстан','Афганистан','Узбекистан','Бразилия']
    vals = [5, 7, 150, 344]
    ch = add_chart(s, XL_CHART_TYPE.COLUMN_CLUSTERED, 46, 134, 380, 146, cats, [('кг/га', vals)])
    plot = ch.plots[0]; plot.gap_width = 60
    ser = ch.series[0]
    # per-point colors: KZ copper (highlight), others steel
    from pptx.oxml.ns import qn as _qn
    pts_cols = [COPPER, STEEL, STEEL, STEEL]
    for i, col in enumerate(pts_cols):
        pt = ser.points[i]
        pt.format.fill.solid(); pt.format.fill.fore_color.rgb = col
    col_labels(plot, 8.5, INK); ch.has_legend = False
    style_cat(ch, 8); style_val(ch, 7.5, mx=380)

    # right KPI stack
    kpi_chip(s, 442, 104, 238, 44, '1,3', 'млн т', 'ёмкость рынка удобрений РК, 2025 (+19% г/г)', big_c=STEEL, big_sz=23)
    kpi_chip(s, 442, 152, 238, 44, '43,6%', '', 'самообеспеченность по азоту (фосфор — 98,5%)', big_c=COPPER, big_sz=23)
    kpi_chip(s, 442, 200, 238, 44, '≈ 6%', '', 'доля завода (75 тыс т ОМУ) в потреблении РК', big_c=STEEL, big_sz=23)
    # imports structure mini
    rect(s, 442, 248, 238, 54, fill=PANEL, line=BORDER, round_=0.10)
    simple(s, 454, 254, 214, 12, 'СТРУКТУРА ИМПОРТА (стоимость)', 7.6, True, GRAY, 'Calibri', spc=0.6)
    text(s, 454, 270, 214, 26, [[
        {'t':'Россия ','sz':10,'b':True,'c':NAVY,'f':'Calibri'},{'t':'73%','sz':10,'b':True,'c':STEEL,'f':'Calibri'},
        {'t':'    ·    Узбекистан ','sz':10,'b':False,'c':INK,'f':'Calibri'},{'t':'11%','sz':10,'b':True,'c':STEEL,'f':'Calibri'}]])

    rect(s, 40, 312, 640, 44, fill=NAVY, round_=0.06)
    text(s, 56, 312, 610, 44, [[
        {'t':'ВЫВОД:  ','sz':9.5,'b':True,'c':TAN,'f':'Calibri'},
        {'t':'сверхнизкое внесение и рост импорта (552 → 700 тыс т за 2023–2025) формируют устойчивый спрос на доступные ОМУ и прямой потенциал импортозамещения.',
         'sz':9.5,'b':False,'c':WHITE,'f':'Calibri'}]], anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.05)
    footer(s, 17, 'АНАЛИЗ РЫНКА · КАЗАХСТАН')
    return s

# ============================================================
# SLIDE D — PRICING & COST
# ============================================================
def slide_price():
    s = add_blank(prs, WHITE)
    eyebrow(s, 'ЦЕНОВОЕ ПОЗИЦИОНИРОВАНИЕ  ·  СЕБЕСТОИМОСТЬ ОМУ', COPPER)
    title(s, 'Себестоимость ОМУ $233/т при цене реализации $450/т')
    subtitle(s, 'ОМУ NPK 8-21: сырьё $194 + производство (с пост. расходами) $39 = $233/т; валовая маржа ≈ 48%')

    # left cost-bridge chart card
    rect(s, 40, 104, 330, 198, fill=WHITE, line=BORDER, round_=0.05, shadow=True)
    simple(s, 54, 114, 300, 16, 'Экономика 1 т ОМУ, $/т', 9.5, True, NAVY, 'Calibri')
    cats = ['Себест.\nОМУ','Цена\nреализ.','Импорт\nкомплекс.']
    vals = [233, 450, 630]
    ch = add_chart(s, XL_CHART_TYPE.COLUMN_CLUSTERED, 46, 134, 318, 146, cats, [('$/т', vals)])
    plot = ch.plots[0]; plot.gap_width = 70
    ser = ch.series[0]
    for i, col in enumerate([GREENT, STEEL, NAVY]):
        pt = ser.points[i]; pt.format.fill.solid(); pt.format.fill.fore_color.rgb = col
    col_labels(plot, 9.5, INK, fmt='$0'); ch.has_legend = False
    style_cat(ch, 8); style_val(ch, 7.5, fmt='$0', mx=700)

    # right competitor table card
    rect(s, 380, 104, 300, 198, fill=WHITE, line=BORDER, round_=0.05, shadow=True)
    simple(s, 394, 112, 280, 14, 'ЦЕНЫ КОНКУРЕНТОВ, $/т', 8.6, True, GRAY, 'Calibri', spc=0.6)
    rows = [
        ('Продукт','NPK','Kusto','Alem'),
        ('Карбамид','N 56','441','467'),
        ('Аммофос','P 52','475','823'),
        ('Сульфоаммофос','20-20-14S','492','584'),
        ('Аммиачная селитра','N 34','303','343'),
    ]
    tb = s.shapes.add_table(len(rows), 4, Pt(392), Pt(130), Pt(276), Pt(150)).table
    tb.first_row = False; tb.horz_banding = False
    clean_table(tb)
    widths = [104, 62, 56, 54]
    for j, wv in enumerate(widths):
        tb.columns[j].width = Pt(wv)
    for i, row in enumerate(rows):
        tb.rows[i].height = Pt(30 if i==0 else 30)
        for j, val in enumerate(row):
            c = tb.cell(i, j)
            c.margin_left = Pt(5); c.margin_right = Pt(3); c.margin_top = Pt(1); c.margin_bottom = Pt(1)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            if i == 0:
                c.fill.solid(); c.fill.fore_color.rgb = NAVY
                col = WHITE; b = True; sz = 8.5
            else:
                c.fill.solid(); c.fill.fore_color.rgb = WHITE if i % 2 else PANEL
                col = INK; b = (j == 0); sz = 8.6
            p = c.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if j < 2 else PP_ALIGN.RIGHT
            r = p.add_run(); r.text = val
            r.font.size = Pt(sz); r.font.bold = b; r.font.name = 'Calibri'
            r.font.color.rgb = col
            if i > 0 and j >= 2:
                r.font.color.rgb = STEEL; r.font.bold = True
            cell_border(c, ('bottom',), (RGBColor(0x35,0x52,0x74) if i==0 else BORDER), 0.75)

    rect(s, 40, 312, 640, 44, fill=NAVY, round_=0.06)
    text(s, 56, 312, 610, 44, [[
        {'t':'ПРЕИМУЩЕСТВО:  ','sz':9.5,'b':True,'c':TAN,'f':'Calibri'},
        {'t':'ОМУ дешевле импортных комплексных удобрений (~$630/т) и минеральных аналогов при эффективности питательных веществ 1:3 (ФАО); логистика ≈ $25/т.',
         'sz':9.5,'b':False,'c':WHITE,'f':'Calibri'}]], anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.05)
    footer(s, 18, 'АНАЛИЗ РЫНКА · ЦЕНЫ И СЕБЕСТОИМОСТЬ')
    return s

# ---- build & insert ----
n0 = len(prs.slides._sldIdLst)
slide_divider(); slide_world(); slide_kz(); slide_price()
# they are appended at indices n0..n0+3 ; move to positions 14,15,16,17 (after slide 14)
for k in range(4):
    move_slide(prs, n0 + k, 14 + k)

# ---- renumber footer page-number boxes sequentially ----
count = 0
for idx, sl in enumerate(prs.slides):
    count = idx + 1
    for sh in sl.shapes:
        if sh.has_text_frame:
            t = sh.text_frame.text.strip()
            try: tp = Emu(sh.top).pt; lf = Emu(sh.left).pt
            except Exception: continue
            if t.isdigit() and tp > 372 and lf > 640:
                run = sh.text_frame.paragraphs[0].runs[0]
                run.text = str(count)

prs.save(OUT)
print("saved", OUT, "slides:", len(prs.slides._sldIdLst))
