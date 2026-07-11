# -*- coding: utf-8 -*-
"""LRT safety poster: latinize ЛРТ->LRT (as a title accent), complete/fix the
Kazakh translation to match the Russian, and apply safe design polish."""
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

SRC = "lrt-poster/source_LRT_ru_kz.pptx"
OUT = "lrt-poster/LRT_safety_poster_FINAL.pptx"
prs = Presentation(SRC)

ORANGE = RGBColor(0xE4,0x57,0x2E)
NAVY   = RGBColor(0x0E,0x3A,0x54)

def find(slide, name):
    return next((sh for sh in slide.shapes if sh.name == name), None)

def set_runs(shape, parts):
    """parts = [(text, rgb), ...]; keeps first run's size/bold/font."""
    p = shape.text_frame.paragraphs[0]
    r0 = p.runs[0]
    sz, bold, name = r0.font.size, r0.font.bold, r0.font.name
    r0.text = parts[0][0]; r0.font.color.rgb = parts[0][1]
    for r in p.runs[1:]:
        r.text = ''
    for t, c in parts[1:]:
        r = p.add_run(); r.text = t
        r.font.size, r.font.bold, r.font.name = sz, bold, name
        r.font.color.rgb = c

def set_text(shape, new):
    p = shape.text_frame.paragraphs[0]
    p.runs[0].text = new
    for r in p.runs[1:]:
        r.text = ''

def set_radius(shape, val):
    g = shape._element.find('.//' + qn('a:prstGeom'))
    gd = g.find('.//' + qn('a:gd')) if g is not None else None
    if gd is not None:
        gd.set('fmla', 'val %d' % val)

# ---- titles: ЛРТ -> LRT, with LRT in accent orange ----
titles = {0: ': правила безопасности', 1: ': қауіпсіздік ережелері'}
for i, tail in titles.items():
    t0 = find(prs.slides[i], 'Text 0')
    set_runs(t0, [('LRT', ORANGE), (tail, NAVY)])

# ---- Kazakh translation completion (slide 2) to match Russian ----
kz = prs.slides[1]
kz_fix = {
    'Text 1':  'Астананың жердегі метросы жолаушыларына арналған қарапайым ережелер',
    'Text 15': 'Газ баллоны, пиротехника, қару.',
    'Text 20': 'Сары сызықтың артында тұрып, пойызды күтіңіз.',
    'Text 21': 'Алдымен шыққандарды өткізіп, содан кейін кіріңіз.',
    'Text 22': 'Тұтқадан берік ұстаныңыз, әсіресе баламен.',
}
for name, new in kz_fix.items():
    set_text(find(kz, name), new)

# ---- safe design polish on both slides ----
SW = Emu(prs.slide_width).pt
for s in prs.slides:
    # full-bleed top accent bar
    bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(0), Pt(0), Pt(SW), Pt(6))
    bar.fill.solid(); bar.fill.fore_color.rgb = NAVY
    bar.line.fill.background(); bar.shadow.inherit = False
    # send the bar to back-ish then in front is fine; keep after logos so it's visible on top edge
    # harmonize the three section-card corner radii
    for nm in ['Shape 3', 'Shape 10', 'Shape 17']:
        sh = find(s, nm)
        if sh is not None:
            set_radius(sh, 6000)

prs.save(OUT)
print("saved", OUT, "slides:", len(prs.slides._sldIdLst))
