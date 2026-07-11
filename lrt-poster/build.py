# -*- coding: utf-8 -*-
"""LRT safety poster: latinize ЛРТ->LRT, complete the Kazakh translation, and
redesign the section blocks into clean white cards with solid colored header
bands (white titles). Applied to both RU (slide 1) and KZ (slide 2)."""
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

SRC = "lrt-poster/source_LRT_ru_kz.pptx"
OUT = "lrt-poster/LRT_safety_poster_FINAL.pptx"
prs = Presentation(SRC)

ORANGE = RGBColor(0xE4,0x57,0x2E)
RED    = RGBColor(0xC4,0x32,0x1B)
BLUE   = RGBColor(0x0B,0x5E,0x86)
NAVY   = RGBColor(0x0E,0x3A,0x54)
WHITE  = RGBColor(0xFF,0xFF,0xFF)
CARDBD = RGBColor(0xE5,0xE9,0xED)

def find(slide, name):
    return next((sh for sh in slide.shapes if sh.name == name), None)

def set_runs(shape, parts):
    p = shape.text_frame.paragraphs[0]; r0 = p.runs[0]
    sz, bold, name = r0.font.size, r0.font.bold, r0.font.name
    r0.text = parts[0][0]; r0.font.color.rgb = parts[0][1]
    for r in p.runs[1:]: r.text = ''
    for t, c in parts[1:]:
        r = p.add_run(); r.text = t
        r.font.size, r.font.bold, r.font.name = sz, bold, name
        r.font.color.rgb = c

def set_text(shape, new, color=None):
    p = shape.text_frame.paragraphs[0]; p.runs[0].text = new
    if color is not None: p.runs[0].font.color.rgb = color
    for r in p.runs[1:]: r.text = ''

def set_radius(shape, val):
    g = shape._element.find('.//' + qn('a:prstGeom'))
    gd = g.find('.//' + qn('a:gd')) if g is not None else None
    if gd is not None: gd.set('fmla', 'val %d' % val)

def style_card(shape, radius=7000):
    shape.fill.solid(); shape.fill.fore_color.rgb = WHITE
    shape.line.color.rgb = CARDBD; shape.line.width = Pt(1.0)
    set_radius(shape, radius)

def add_band(slide, card, color, h=50, radius=7000):
    b = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, card.left, card.top, card.width, Pt(h))
    b.fill.solid(); b.fill.fore_color.rgb = color
    b.line.fill.background(); b.shadow.inherit = False
    set_radius(b, radius)
    sp = b._element; sp.getparent().remove(sp); card._element.addnext(sp)  # z: above card, below icon/title
    return b

SECTIONS = [('Shape 3', 'Shape 4', 'Text 5', ORANGE),
            ('Shape 10', 'Shape 11', 'Text 12', RED),
            ('Shape 17', 'Shape 18', 'Text 19', BLUE)]

title_tail = {0: ': правила безопасности', 1: ': қауіпсіздік ережелері'}
kz_fix = {
    'Text 1':  'Астананың жердегі метросы жолаушыларына арналған қарапайым ережелер',
    'Text 15': 'Газ баллоны, пиротехника, қару.',
    'Text 20': 'Сары сызықтың артында тұрып, пойызды күтіңіз.',
    'Text 21': 'Алдымен шыққандарды өткізіп, содан кейін кіріңіз.',
    'Text 22': 'Тұтқадан берік ұстаныңыз, әсіресе баламен.',
}

SW = Emu(prs.slide_width).pt
for i, s in enumerate(prs.slides):
    t0 = find(s, 'Text 0'); set_runs(t0, [('LRT', ORANGE), (title_tail[i], NAVY)]); t0.width = Pt(480)
    bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(0), Pt(0), Pt(SW), Pt(7))
    bar.fill.solid(); bar.fill.fore_color.rgb = NAVY; bar.line.fill.background(); bar.shadow.inherit = False
    for card_n, circ_n, title_n, color in SECTIONS:
        card, circ, title = find(s, card_n), find(s, circ_n), find(s, title_n)
        style_card(card)
        add_band(s, card, color)
        set_text(title, title.text_frame.text.strip(), WHITE)
        # keep the icon badge the band colour (blends in) so the white icon glyph stays visible
        circ.fill.solid(); circ.fill.fore_color.rgb = color
        circ.line.fill.background()
    if i == 1:
        for nm, new in kz_fix.items():
            set_text(find(s, nm), new)

prs.save(OUT)
print("saved", OUT, "slides:", len(prs.slides._sldIdLst))
