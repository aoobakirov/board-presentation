# -*- coding: utf-8 -*-
"""Consistency pass on the user's v2 deck + new 'Producers of Kazakhstan' slide.
- Insert a styled producers slide after the Kazakhstan market slide
- Rebrand KazBioFert->QazBioFert, BBFF->SBFF across the whole deck
- Renumber footer page numbers sequentially
- De-duplicate the two identical 'ПРОГНОЗ БЮДЖЕТА' budget titles
- Fix the renamed thesis slide's footer label
Operates on source/base_manual_v2.pptx so all manual edits are preserved.
"""
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_TICK_MARK
from build_lib import (rect, text, simple, line_seg, kpi_chip, footer, source_note,
                       hrule, card_head, add_blank, move_slide,
                       NAVY, NAVYDK, STEEL, COPPER, TAN, INK, GRAY, LBLUE, WHITE,
                       BORDER, PANEL, MIST)

SRC = "source/base_manual_v2.pptx"
OUT = "KazBioFert_Board_Presentation_FINAL_20260709.pptx"
prs = Presentation(SRC)

# ---------- helpers ----------
def header(s, kicker, headline, sub):
    simple(s, 40, 26, 641, 16, kicker, 11, True, COPPER, 'Calibri', spc=1.4)
    simple(s, 40, 46, 648, 28, headline, 21, True, NAVY, 'Arial')
    simple(s, 40, 73, 648, 18, sub, 10.5, False, RGB(0x53,0x62,0x72), 'Calibri')
    hrule(s, 97, 40, 641, BORDER, 0.75)

def insight(s, prefix, body, y=304):
    rect(s, 40, y, 640, 44, fill=NAVY, round_=0.06)
    text(s, 56, y, 610, 44, [[
        {'t': prefix + '  ', 'sz': 9.5, 'b': True, 'c': TAN, 'f': 'Calibri'},
        {'t': body, 'sz': 9.5, 'b': False, 'c': WHITE, 'f': 'Calibri'}]],
        anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.05)

from pptx.dml.color import RGBColor as RGB

def producer_card(s, x, y, w, h, name, detail, accent=STEEL):
    rect(s, x, y, w, h, fill=WHITE, line=BORDER, line_w=1.0, round_=0.10, shadow=True)
    rect(s, x + 2, y + 5, 3.0, h - 10, fill=accent)
    simple(s, x + 14, y + 8, w - 22, 18, name, 13, True, NAVY, 'Calibri')
    simple(s, x + 14, y + 28, w - 24, h - 30, detail, 8.8, False, INK, 'Calibri', line_spacing=1.02)

# ---------- NEW SLIDE: producers of Kazakhstan ----------
def slide_producers():
    s = add_blank(prs, WHITE)
    header(s, 'РЫНОК КАЗАХСТАНА  ·  ПРОИЗВОДИТЕЛИ  ·  KAZPHOSPHATE / КАЗАЗОТ, 2025',
           'Внутреннее производство — два игрока; ниша ОМУ свободна',
           'Казфосфат и КазАзот дают ≈ 96% производства РК; гранулированные органоминеральные удобрения не выпускает никто')
    # left chart
    rect(s, 40, 106, 392, 186, fill=WHITE, line=BORDER, round_=0.05, shadow=True)
    card_head(s, 54, 115, 360, 'Крупнейшие производители РК, тыс т в год')
    cd = CategoryChartData()
    cd.categories = ['Казфосфат\n(аммофос, P)', 'КазАзот\n(селитра, N)']
    cd.add_series('тыс т', (750, 330))
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Pt(50), Pt(136), Pt(372), Pt(148), cd)
    ch = gf.chart; ch.has_title = False
    ch.font.size = Pt(8); ch.font.name = 'Calibri'; ch.font.color.rgb = INK
    plot = ch.plots[0]; plot.gap_width = 90
    ser = ch.series[0]
    for i, col in enumerate([STEEL, LBLUE]):
        ser.points[i].format.fill.solid(); ser.points[i].format.fill.fore_color.rgb = col
    plot.has_data_labels = True
    dl = plot.data_labels; dl.number_format = '0'; dl.number_format_is_linked = False
    dl.font.size = Pt(10); dl.font.bold = True; dl.font.name = 'Calibri'; dl.font.color.rgb = INK
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    ch.has_legend = False
    ca = ch.category_axis
    ca.tick_labels.font.size = Pt(8); ca.tick_labels.font.name = 'Calibri'; ca.tick_labels.font.color.rgb = INK
    ca.has_major_gridlines = False; ca.major_tick_mark = XL_TICK_MARK.NONE
    ca.format.line.color.rgb = BORDER; ca.format.line.width = Pt(0.75)
    va = ch.value_axis
    va.tick_labels.font.size = Pt(7.5); va.tick_labels.font.color.rgb = GRAY
    va.has_major_gridlines = True; va.major_gridlines.format.line.color.rgb = MIST
    va.major_gridlines.format.line.width = Pt(0.75); va.format.line.fill.background()
    va.major_tick_mark = XL_TICK_MARK.NONE; va.maximum_scale = 800
    # right cards
    producer_card(s, 442, 106, 238, 58, 'Казфосфат',
                  'Фосфорные (аммофос) · ~750 тыс т в 2025 (+58% г/г) · крупнейший в СНГ, значит. экспорт', STEEL)
    producer_card(s, 442, 168, 238, 58, 'КазАзот',
                  'Азотные (селитра, аммиак) · мощность ~330 тыс т/год · единственный в РК', LBLUE)
    rect(s, 442, 230, 238, 62, fill=PANEL, line=COPPER, line_w=1.25, round_=0.08)
    rect(s, 444, 235, 3.0, 52, fill=COPPER)
    simple(s, 456, 236, 214, 12, 'СВОБОДНАЯ НИША · ОМУ', 8.2, True, COPPER, 'Calibri', spc=0.6)
    simple(s, 456, 252, 214, 36, '0 производителей гранулированных ОМУ в РК → проект 75 тыс т без прямых внутр. конкурентов',
           8.8, False, INK, 'Calibri', line_spacing=1.02)
    insight(s, 'ВЫВОД:',
            'оба игрока выпускают минеральные N и P и значительную часть экспортируют; гранулированные ОМУ в РК '
            'не производит никто — конкуренция проекта не внутренняя, а импорт (~700 тыс т/год).')
    source_note(s, 'Источники: Kazphosphate / Kursiv, 2025 (аммофос 750 тыс т, +58%); АО «КазАзот» (мощность ~330 тыс т); '
                   'БНС РК; The Diplomat (два игрока ≈ 96% производства).')
    footer(s, 16, 'АНАЛИЗ РЫНКА · ПРОИЗВОДИТЕЛИ РК')
    return s

# find Kazakhstan market slide, insert producers slide right after it
kz_idx = next(i for i, sl in enumerate(prs.slides)
              if any(sh.has_text_frame and 'РЫНОК КАЗАХСТАНА' in sh.text_frame.text for sh in sl.shapes))
n0 = len(prs.slides._sldIdLst)
slide_producers()
move_slide(prs, n0, kz_idx + 1)

# ---------- rebrand everywhere ----------
def rebrand_runs():
    left = 0
    for sl in prs.slides:
        for sh in sl.shapes:
            if sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    for r in p.runs:
                        if 'KazBioFert' in r.text or 'BBFF' in r.text:
                            r.text = r.text.replace('KazBioFert', 'QazBioFert').replace('BBFF', 'SBFF')
                    # paragraph-level fallback if a token was split across runs
                    if ('KazBioFert' in p.text or 'BBFF' in p.text) and p.runs:
                        joined = ''.join(r.text for r in p.runs).replace('KazBioFert', 'QazBioFert').replace('BBFF', 'SBFF')
                        p.runs[0].text = joined
                        for r in p.runs[1:]:
                            r.text = ''
    for sl in prs.slides:
        for sh in sl.shapes:
            if sh.has_text_frame and ('KazBioFert' in sh.text_frame.text or 'BBFF' in sh.text_frame.text):
                left += 1
    return left
print("brand left after rebrand:", rebrand_runs())

# ---------- de-duplicate budget titles + fix labels ----------
def set_full_text(shape, new):
    p = shape.text_frame.paragraphs[0]
    if p.runs:
        p.runs[0].text = new
        for r in p.runs[1:]:
            r.text = ''
    else:
        p.add_run().text = new

for sl in prs.slides:
    tdims = [(len(sh.table.rows), len(sh.table.columns)) for sh in sl.shapes if sh.has_table]
    title_sh = next((sh for sh in sl.shapes if sh.has_text_frame
                     and sh.text_frame.text.strip() == 'ПРОГНОЗ БЮДЖЕТА НА 2026'), None)
    if not title_sh:
        continue
    if (31, 22) in tdims:      # Jan-Jun actuals
        set_full_text(title_sh, 'ИСПОЛНЕНИЕ БЮДЖЕТА · 1 ПОЛУГОДИЕ 2026')
        for sh in sl.shapes:
            if sh.has_text_frame and sh.text_frame.text.strip() == 'ПРОГНОЗ БЮДЖЕТА':
                set_full_text(sh, 'ИСПОЛНЕНИЕ · 1 ПОЛУГОДИЕ')
    elif (34, 9) in tdims:     # Jul-Dec plan
        set_full_text(title_sh, 'ПРОГНОЗ БЮДЖЕТА · 2 ПОЛУГОДИЕ 2026')
        for sh in sl.shapes:
            if sh.has_text_frame and sh.text_frame.text.strip() == 'ПРОГНОЗ БЮДЖЕТА':
                set_full_text(sh, 'ПРОГНОЗ · 2 ПОЛУГОДИЕ')

# fix renamed thesis slide footer label
for sl in prs.slides:
    has_share = any(sh.has_text_frame and sh.text_frame.text.strip() == 'Потенциальная доля на рынке' for sh in sl.shapes)
    if has_share:
        for sh in sl.shapes:
            if sh.has_text_frame and 'ИНВЕСТИЦИОННЫЙ ТЕЗИС' in sh.text_frame.text:
                set_full_text(sh, 'АНАЛИЗ РЫНКА · ПОТЕНЦИАЛЬНАЯ ДОЛЯ')

# ---------- renumber footer page numbers ----------
for idx, sl in enumerate(prs.slides):
    for sh in sl.shapes:
        if sh.has_text_frame:
            t = sh.text_frame.text.strip()
            try:
                tp, lf = Emu(sh.top).pt, Emu(sh.left).pt
            except Exception:
                continue
            if t.isdigit() and tp > 372 and lf > 640:
                sh.text_frame.paragraphs[0].runs[0].text = str(idx + 1)

prs.save(OUT)
print("saved", OUT, "slides:", len(prs.slides._sldIdLst))
