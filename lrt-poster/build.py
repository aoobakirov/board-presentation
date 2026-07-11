# -*- coding: utf-8 -*-
"""LRT safety poster — bold redesign: navy gradient hero, big section index
numbers, clean colored header bands. RU (slide 1) + KZ (slide 2). Keeps images,
completes the Kazakh translation, latinizes ЛРТ->LRT."""
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

SRC = "lrt-poster/source_LRT_ru_kz.pptx"
OUT = "lrt-poster/LRT_safety_poster_FINAL.pptx"
prs = Presentation(SRC)

ORANGE=RGBColor(0xE4,0x57,0x2E); RED=RGBColor(0xC4,0x32,0x1B); BLUE=RGBColor(0x0B,0x5E,0x86)
NAVY=RGBColor(0x0E,0x3A,0x54); NAVY2=RGBColor(0x16,0x5A,0x7C); GOLD=RGBColor(0xFD,0xB5,0x15)
WHITE=RGBColor(0xFF,0xFF,0xFF); CARDBD=RGBColor(0xE5,0xE9,0xED)
LT=RGBColor(0xC6,0xD8,0xE4); MUT=RGBColor(0x93,0xAB,0xBA)
DY=20  # shift content down to make room for the hero

def find(s,n): return next((x for x in s.shapes if x.name==n),None)
def rid(sh):
    p=sh._element.getparent();
    if p is not None: p.remove(sh._element)
def to_back(sh):
    el=sh._element; tree=el.getparent(); tree.remove(el)
    anchor=tree.find(qn('p:grpSpPr'))
    anchor.addnext(el)
def set_radius(sh,val):
    g=sh._element.find('.//'+qn('a:prstGeom')); gd=g.find('.//'+qn('a:gd')) if g is not None else None
    if gd is not None: gd.set('fmla','val %d'%val)
def grad(shape,c1,c2,ang=2700000):
    spPr=shape._element.spPr
    for t in ('a:noFill','a:solidFill','a:gradFill','a:blipFill','a:pattFill'):
        e=spPr.find(qn(t))
        if e is not None: spPr.remove(e)
    g=spPr.makeelement(qn('a:gradFill'),{}); lst=g.makeelement(qn('a:gsLst'),{})
    for pos,c in [(0,c1),(100000,c2)]:
        gs=lst.makeelement(qn('a:gs'),{'pos':str(pos)})
        cl=gs.makeelement(qn('a:srgbClr'),{'val':'%02X%02X%02X'%(c[0],c[1],c[2])})
        gs.append(cl); lst.append(gs)
    g.append(lst); g.append(g.makeelement(qn('a:lin'),{'ang':str(ang),'scaled':'1'}))
    ln=spPr.find(qn('a:ln'))
    (ln.addprevious(g) if ln is not None else spPr.append(g))
def set_runs(shape,parts,anchor=None):
    tf=shape.text_frame; p=tf.paragraphs[0]; r0=p.runs[0]
    sz,bold,name=r0.font.size,r0.font.bold,r0.font.name
    r0.text=parts[0][0]; r0.font.color.rgb=parts[0][1]
    for r in p.runs[1:]: r.text=''
    for t,c in parts[1:]:
        r=p.add_run(); r.text=t; r.font.size,r.font.bold,r.font.name=sz,bold,name; r.font.color.rgb=c
    if anchor is not None: tf.vertical_anchor=anchor
def set_text(shape,new,color=None):
    p=shape.text_frame.paragraphs[0]; p.runs[0].text=new
    if color is not None: p.runs[0].font.color.rgb=color
    for r in p.runs[1:]: r.text=''
def rect(s,x,y,w,h,fill=None,line=None,lw=1.0,round_=None,grad2=None):
    shp=MSO_SHAPE.ROUNDED_RECTANGLE if round_ is not None else MSO_SHAPE.RECTANGLE
    sh=s.shapes.add_shape(shp,Pt(x),Pt(y),Pt(w),Pt(h))
    if round_ is not None: set_radius(sh,round_)
    if grad2: grad(sh,grad2[0],grad2[1])
    elif fill is not None: sh.fill.solid(); sh.fill.fore_color.rgb=fill
    else: sh.fill.background()
    if line is not None: sh.line.color.rgb=line; sh.line.width=Pt(lw)
    else: sh.line.fill.background()
    sh.shadow.inherit=False; return sh
def txt(s,x,y,w,h,runs,sz,bold,color,align=PP_ALIGN.LEFT,anchor=MSO_ANCHOR.TOP,font='Calibri'):
    tb=s.shapes.add_textbox(Pt(x),Pt(y),Pt(w),Pt(h)); tf=tb.text_frame; tf.word_wrap=True
    tf.vertical_anchor=anchor
    for m in ('left','right','top','bottom'): setattr(tf,'margin_'+m,0)
    p=tf.paragraphs[0]; p.alignment=align
    for t,c in (runs if isinstance(runs,list) else [(runs,color)]):
        r=p.add_run(); r.text=t; r.font.size=Pt(sz); r.font.bold=bold; r.font.name=font; r.font.color.rgb=c
    return tb

SECTIONS=[('Shape 3','Shape 4','Text 5','Image 2',ORANGE,'1'),
          ('Shape 10','Shape 11','Text 12','Image 7',RED,'2'),
          ('Shape 17','Shape 18','Text 19','Image 12',BLUE,'3')]
title_tail={0:': правила безопасности',1:': қауіпсіздік ережелері'}
sub_txt={0:'Простые правила для пассажиров наземного метро Астаны',1:'Астананың жердегі метросы жолаушыларына арналған қарапайым ережелер'}
iss_txt={0:'Прокуратура города Астаны',1:'Астана қаласының прокуратурасы'}
kz_fix={'Text 15':'Газ баллоны, пиротехника, қару.','Text 20':'Сары сызықтың артында тұрып, пойызды күтіңіз.',
        'Text 21':'Алдымен шыққандарды өткізіп, содан кейін кіріңіз.','Text 22':'Тұтқадан берік ұстаныңыз, әсіресе баламен.'}

SW=Emu(prs.slide_width).pt
for i,s in enumerate(prs.slides):
    # 1) shift existing content down to clear the hero
    for sh in list(s.shapes):
        if Emu(sh.top).pt>=129: sh.top=Pt(Emu(sh.top).pt+DY)
    # 2) HERO background (gradient) + gold accent
    hero=rect(s,0,0,SW,140,grad2=(NAVY,NAVY2)); to_back(hero)
    rect(s,0,140,SW,5,fill=GOLD)
    # 3) white chips behind logos
    for logo,(cx,cy,cw,ch) in [(find(s,'Image 0'),(30,20,78,78)),(find(s,'Image 1'),(466,24,99,71))]:
        if logo is not None:
            chip=rect(s,cx,cy,cw,ch,fill=WHITE,round_=12000)
            el=chip._element; el.getparent().remove(el); logo._element.addprevious(el)
    # 4) hero texts (reuse existing title/subtitle/issuer shapes, restyle white)
    t0=find(s,'Text 0'); t0.left,t0.top,t0.width,t0.height=Pt(118),Pt(40),Pt(352),Pt(30)
    set_runs(t0,[('LRT',ORANGE),(title_tail[i],WHITE)],anchor=MSO_ANCHOR.MIDDLE)
    for r in t0.text_frame.paragraphs[0].runs: r.font.size=Pt(21)
    t1=find(s,'Text 1'); t1.left,t1.top,t1.width,t1.height=Pt(36),Pt(100),Pt(524),Pt(18)
    set_text(t1,sub_txt[i],LT)
    for r in t1.text_frame.paragraphs[0].runs: r.font.size=Pt(11.5)
    t2=find(s,'Text 2'); t2.left,t2.top,t2.width=Pt(36),Pt(119),Pt(524)
    set_text(t2,iss_txt[i],MUT)
    # 5) sections: white card + colored band + big index number badge
    for card_n,badge_n,title_n,icon_n,color,num in SECTIONS:
        card=find(s,card_n); badge=find(s,badge_n); title=find(s,title_n); icon=find(s,icon_n)
        card.fill.solid(); card.fill.fore_color.rgb=WHITE
        card.line.color.rgb=CARDBD; card.line.width=Pt(1.0); set_radius(card,7000)
        band=rect(s,Emu(card.left).pt,Emu(card.top).pt,Emu(card.width).pt,52,fill=color,round_=7000)
        el=band._element; el.getparent().remove(el); card._element.addnext(el)  # above card, below badge/title
        # numbered badge (reuse the circle shape): white disc + colored number
        badge.fill.solid(); badge.fill.fore_color.rgb=WHITE; badge.line.fill.background()
        badge.width=Pt(40); badge.height=Pt(40)
        tf=badge.text_frame; tf.word_wrap=False; tf.vertical_anchor=MSO_ANCHOR.MIDDLE
        for m in ('left','right','top','bottom'): setattr(tf,'margin_'+m,0)
        p=tf.paragraphs[0]; p.alignment=PP_ALIGN.CENTER
        r=p.add_run(); r.text=num; r.font.size=Pt(22); r.font.bold=True; r.font.name='Calibri'; r.font.color.rgb=color
        if icon is not None: rid(icon)   # drop small section glyph in favour of the number
        set_text(title,title.text_frame.text.strip(),WHITE)
    # 6) translation + LRT already handled in titles; KZ bullet completion
    if i==1:
        set_text(find(s,'Text 1'),sub_txt[1],LT)
        for nm,new in kz_fix.items(): set_text(find(s,nm),new)

prs.save(OUT); print("saved",OUT)
