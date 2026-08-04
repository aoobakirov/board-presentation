# -*- coding: utf-8 -*-
"""Сборка board-презентации из отрендеренных слайдов (HTML → PNG → PPTX).
Конвейер: strategy/deck/deck.html  →  node strategy/deck/render.mjs (Chromium, 2×)  →  png/  →  этот скрипт."""
from pptx import Presentation
from pptx.util import Inches
import glob, os

prs=Presentation(); prs.slide_width=Inches(13.333); prs.slide_height=Inches(7.5)
blank=prs.slide_layouts[6]
pngs=sorted(glob.glob(os.path.join(os.path.dirname(__file__),'deck','png','slide-*.png')))
assert pngs, 'нет PNG — сначала: node strategy/deck/render.mjs'
for p in pngs:
    s=prs.slides.add_slide(blank)
    s.shapes.add_picture(p,0,0,width=prs.slide_width,height=prs.slide_height)
out=os.path.join(os.path.dirname(__file__),'WayStar_Стратегия_презентация.pptx')
prs.save(out)
print('saved',out,'·',len(prs.slides),'слайдов')
