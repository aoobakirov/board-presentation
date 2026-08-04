# -*- coding: utf-8 -*-
"""ТОО «WayStar Group» — board-презентация стратегии группы на 5 лет (16:9)."""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

NAVY=RGBColor(0x13,0x29,0x4B); STEEL=RGBColor(0x2F,0x66,0x90); COPPER=RGBColor(0xC4,0x77,0x2F)
INK=RGBColor(0x2A,0x37,0x46); GRAY=RGBColor(0x6B,0x7A,0x8D); GREEN=RGBColor(0x2E,0x7D,0x46)
WHITE=RGBColor(0xFF,0xFF,0xFF); LIGHT=RGBColor(0xF1,0xF5,0xF9); LINE=RGBColor(0xCF,0xDA,0xE4)
MIST=RGBColor(0xBF,0xD0,0xE0); PANEL=RGBColor(0xEE,0xF3,0xF8); MINT=RGBColor(0xE8,0xF3,0xEC)

prs=Presentation(); prs.slide_width=Inches(13.333); prs.slide_height=Inches(7.5)
BLANK=prs.slide_layouts[6]; W=prs.slide_width

def slide(): return prs.slides.add_slide(BLANK)

def box(sl,l,t,w,h,fill=None,line=None,rounded=False,line_w=1.0):
    shp=sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE,l,t,w,h)
    shp.shadow.inherit=False
    if fill is None: shp.fill.background()
    else: shp.fill.solid(); shp.fill.fore_color.rgb=fill
    if line is None: shp.line.fill.background()
    else: shp.line.color.rgb=line; shp.line.width=Pt(line_w)
    return shp

def txt(sl,l,t,w,h,paras,anchor=MSO_ANCHOR.TOP,align=PP_ALIGN.LEFT,wrap=True):
    tb=sl.shapes.add_textbox(l,t,w,h); tf=tb.text_frame; tf.word_wrap=wrap
    tf.vertical_anchor=anchor
    tf.margin_left=0;tf.margin_right=0;tf.margin_top=0;tf.margin_bottom=0
    for i,pr in enumerate(paras):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        p.alignment=pr.get('align',align)
        if 'sa' in pr: p.space_after=Pt(pr['sa'])
        if 'sb' in pr: p.space_before=Pt(pr['sb'])
        if 'ls' in pr: p.line_spacing=pr['ls']
        for (t_,sz,bd,cl) in pr['runs']:
            r=p.add_run(); r.text=t_; f=r.font; f.size=Pt(sz); f.bold=bd; f.color.rgb=cl; f.name='Calibri'
    return tb

def header(sl,title,n,kicker=None):
    box(sl,0,0,W,Inches(1.0),fill=NAVY)
    box(sl,0,Inches(1.0),W,Inches(0.055),fill=COPPER)
    if kicker:
        txt(sl,Inches(0.6),Inches(0.14),Inches(12.1),Inches(0.25),[{'runs':[(kicker.upper(),10,True,MIST)]}])
        txt(sl,Inches(0.6),Inches(0.36),Inches(12.1),Inches(0.55),[{'runs':[(title,23,True,WHITE)]}],anchor=MSO_ANCHOR.MIDDLE)
    else:
        txt(sl,Inches(0.6),Inches(0.1),Inches(12.1),Inches(0.82),[{'runs':[(title,25,True,WHITE)]}],anchor=MSO_ANCHOR.MIDDLE)
    box(sl,Inches(0.6),Inches(7.06),Inches(12.13),Pt(1),fill=LINE)
    txt(sl,Inches(0.6),Inches(7.12),Inches(9),Inches(0.3),[{'runs':[('WayStar Group · Стратегия группы 2026–2031',9,False,GRAY)]}])
    txt(sl,Inches(11.9),Inches(7.12),Inches(0.83),Inches(0.3),[{'runs':[(str(n),9,True,GRAY)]}],align=PP_ALIGN.RIGHT)

def tiles(sl,l,t,total_w,items,h=Inches(1.35),gap=Inches(0.22)):
    n=len(items); w=int((total_w-gap*(n-1))/n)
    for i,(big,label,color) in enumerate(items):
        x=int(l+i*(w+gap))
        box(sl,x,t,w,h,fill=LIGHT,rounded=True)
        box(sl,x,t+Inches(0.12),Inches(0.09),h-Inches(0.24),fill=color)
        txt(sl,x+Inches(0.28),t+Inches(0.16),w-Inches(0.4),Inches(0.6),[{'runs':[(big,25,True,color)]}])
        txt(sl,x+Inches(0.28),t+Inches(0.78),w-Inches(0.42),h-Inches(0.85),[{'runs':[(label,11,False,INK)]}])

def card(sl,l,t,w,h,title,lines,accent=STEEL,tsize=13,lsize=10.5,fill=WHITE):
    box(sl,l,t,w,h,fill=fill,line=LINE,rounded=True)
    paras=[{'runs':[(title,tsize,True,accent)],'sa':5}]
    for ln in lines:
        if isinstance(ln,tuple):
            paras.append({'runs':[('▪  ',lsize,False,accent),(ln[0],lsize,True,INK),(ln[1],lsize,False,INK)],'sa':3,'ls':1.02})
        else:
            paras.append({'runs':[('▪  ',lsize,False,accent),(ln,lsize,False,INK)],'sa':3,'ls':1.02})
    txt(sl,l+Inches(0.24),t+Inches(0.2),w-Inches(0.46),h-Inches(0.34),paras)

def set_cell(cell,text,size,bold,color,align=PP_ALIGN.LEFT):
    cell.margin_left=Inches(0.09);cell.margin_right=Inches(0.06);cell.margin_top=Inches(0.02);cell.margin_bottom=Inches(0.02)
    cell.vertical_anchor=MSO_ANCHOR.MIDDLE
    p=cell.text_frame.paragraphs[0]; p.alignment=align
    r=p.add_run(); r.text=str(text); f=r.font; f.size=Pt(size); f.bold=bold; f.color.rgb=color; f.name='Calibri'

def ptable(sl,l,t,w,h,headers,rows,col_w=None,fs=10,hfs=10,aligns=None):
    from pptx.util import Emu
    gt=sl.shapes.add_table(len(rows)+1,len(headers),l,t,w,h).table
    gt.first_row=True; gt.horz_banding=False
    for j,htext in enumerate(headers):
        cell=gt.cell(0,j); cell.fill.solid(); cell.fill.fore_color.rgb=NAVY
        set_cell(cell,htext,hfs,True,WHITE, aligns[j] if aligns else PP_ALIGN.LEFT)
    for i,row in enumerate(rows,1):
        for j,val in enumerate(row):
            cell=gt.cell(i,j); cell.fill.solid()
            cell.fill.fore_color.rgb=LIGHT if i%2 else WHITE
            bold = isinstance(val,tuple) and val[1]
            v = val[0] if isinstance(val,tuple) else val
            set_cell(cell,v,fs,bold,INK, aligns[j] if aligns else PP_ALIGN.LEFT)
    if col_w:
        for j,wd in enumerate(col_w): gt.columns[j].width=wd
    return gt

def chain(sl,l,t,w,steps,accent=NAVY,h=Inches(0.95)):
    n=len(steps); gap=Inches(0.1); aw=Inches(0.28)
    bw=int((w-gap*2*(n-1)-aw*(n-1))/n)
    x=int(l)
    for i,s in enumerate(steps):
        box(sl,x,t,bw,h,fill=PANEL,line=accent,rounded=True,line_w=1.25)
        txt(sl,x+Inches(0.12),t,bw-Inches(0.24),h,[{'runs':[(s,10.5,True,accent)]}],anchor=MSO_ANCHOR.MIDDLE,align=PP_ALIGN.CENTER)
        x+=bw
        if i<n-1:
            ar=sl.shapes.add_shape(MSO_SHAPE.CHEVRON,int(x+gap),int(t+h/2-Inches(0.2)),aw,Inches(0.4))
            ar.fill.solid(); ar.fill.fore_color.rgb=COPPER; ar.line.fill.background(); ar.shadow.inherit=False
            x+=int(gap*2+aw)

# ============ 1. ТИТУЛ ============
s=slide()
box(s,0,0,W,prs.slide_height,fill=NAVY)
box(s,0,Inches(4.55),W,Inches(0.06),fill=COPPER)
txt(s,Inches(0.9),Inches(1.5),Inches(11.5),Inches(0.5),[{'runs':[('WAYSTAR GROUP',18,True,MIST)]}])
txt(s,Inches(0.9),Inches(2.1),Inches(11.5),Inches(2.2),
    [{'runs':[('Стратегия группы',44,True,WHITE)],'sa':4},
     {'runs':[('на 5 лет · 2026–2031',30,True,RGBColor(0xAF,0xC6,0xDC))]}])
txt(s,Inches(0.9),Inches(4.75),Inches(11.5),Inches(0.6),
    [{'runs':[('Логистика  ·  Трейдинг  ·  Цифровые решения  ·  Международный центр в Дубае',15,False,MIST)]}])
txt(s,Inches(0.9),Inches(6.7),Inches(11.5),Inches(0.4),
    [{'runs':[('Конфиденциально · Проект для совета директоров',10,False,RGBColor(0x8FA,0x00,0x00) if False else RGBColor(0x7F,0x93,0xA8))]}])

# ============ 2. ТЕЗИС ============
s=slide(); header(s,'Стратегический тезис',2,kicker='Резюме')
tiles(s,Inches(0.6),Inches(1.35),Inches(12.13),
      [('150 000 т/мес','Объём перевозок',STEEL),('200+','Маршрутов',STEEL),
       ('3','Дивизиона группы',NAVY),('$50 млн','Оборот Дубая · год 1',COPPER)])
card(s,Inches(0.6),Inches(3.0),Inches(5.95),Inches(3.75),'Кто мы',
     [('','Вертикально интегрированная группа: логистика + трейдинг + цифровые решения.'),
      ('','Контролируем цепочку «производитель → порт → экспортный рынок».'),
      ('','Прямые контракты с переработчиками, экспорт до борта в порту Актау.'),
      ('','Собственный и привлечённый мультимодальный парк.')],accent=NAVY)
card(s,Inches(6.78),Inches(3.0),Inches(5.95),Inches(3.75),'Куда идём',
     [('Лидер ','агро-экспортной логистики и трейдинга региона.'),
      ('Дубай — ','международный сбыт и торговое финансирование, $50 млн в год 1.'),
      ('Парк ','пищевых цистерн 20 → 5 000 под управлением (№1 в нише).'),
      ('Управляем ','по валовой прибыли, а не «оборотом ради оборота».')],accent=GREEN,fill=MINT)

# ============ 3. КТО МЫ: 3 ДИВИЗИОНА ============
s=slide(); header(s,'Три дивизиона и вертикальная интеграция',3,kicker='Кто мы')
cw=Inches(3.9)
card(s,Inches(0.6),Inches(1.35),cw,Inches(2.9),'Логистика',
     ['ЖД, авто, авиа, мультимодальные перевозки по РК, СНГ и на экспорт.',
      'Парк: пищевые цистерны, крытые вагоны, полувагоны, зерновозы.',
      '150 000 т/мес · 200+ маршрутов · 20 собственных единиц.'],accent=STEEL)
card(s,Inches(4.72),Inches(1.35),cw,Inches(2.9),'Трейдинг',
     ['Растительные масла, зерно, масличные, навалочные грузы.',
      'Прямые контракты с переработчиками, прозрачная цена.',
      'Экспорт до борта судна в порту Актау.'],accent=COPPER)
card(s,Inches(8.83),Inches(1.35),cw,Inches(2.9),'Цифровые решения',
     ['Планирование отгрузок и контроль движения груза.',
      'Прослеживаемость на каждом этапе маршрута.',
      'Маржа на рейс и на сделку в реальном времени.'],accent=NAVY)
txt(s,Inches(0.6),Inches(4.5),Inches(12.1),Inches(0.35),[{'runs':[('Цепочка создания стоимости',13,True,NAVY)]}])
chain(s,Inches(0.6),Inches(4.95),Inches(12.13),
      ['Закуп у переработчика','Мультимодальная логистика','Экспорт: Актау / Китай','Сбыт через Дубай','Цифровой контроль'])
txt(s,Inches(0.6),Inches(6.15),Inches(12.1),Inches(0.7),
    [{'runs':[('Преимущество: ',12,True,COPPER),('единый ответственный за всю поставку и дополнительная маржа на каждом звене — того, что нет у разрозненных подрядчиков.',12,False,INK)]}])

# ============ 4. РЫНОК ============
s=slide(); header(s,'Рынок: окно возможностей',4,kicker='Диагностика')
tiles(s,Inches(0.6),Inches(1.35),Inches(12.13),
      [('+54%','Экспорт растит. масла · №2 в Китай',GREEN),('$7 млрд','Агроэкспорт Казахстана',STEEL),
       ('×2','Молоко к ~2028',STEEL),('×3–4','Транскаспийский маршрут к 2030',COPPER)])
ptable(s,Inches(0.6),Inches(3.05),Inches(12.13),Inches(3.0),
       ['Драйвер','Показатель','Значение для группы'],
       [['Растительное масло','Пр-во +17,4%; экспорт $944 млн (+54%); цель топ-3','Флагман трейдинга и грузопоток'],
        ['Зерно и масличные','Агроэкспорт $7 млрд; рост переработки','Трейдинг зерна, загрузка зерновозов'],
        ['Молоко / пищепром','Удвоение до 1,2 млн т; инвестиции +57,7%','База для пищевых цистерн'],
        ['Логистика РК','$29→36 млрд; автоперевозки +18,2%','Растущий рынок перевозок'],
        ['Дубай','Мировой торгово-финансовый центр','Рынки сбыта и финансирование']],
       col_w=[Inches(2.7),Inches(5.4),Inches(4.03)],fs=10,hfs=10)
txt(s,Inches(0.6),Inches(6.25),Inches(12.1),Inches(0.6),
    [{'runs':[('Вывод: ',12,True,COPPER),('агроэкспорт РК растёт двузначно, а группа уже владеет и логистикой, и трейдингом — осталось масштабировать цепочку и открыть сбыт через Дубай.',12,False,INK)]}])

# ============ 5. АМБИЦИЯ ============
s=slide(); header(s,'Амбиция группы 2031',5,kicker='Куда идём')
box(s,Inches(0.6),Inches(1.35),Inches(12.13),Inches(1.15),fill=MINT,line=LINE,rounded=True)
txt(s,Inches(0.9),Inches(1.5),Inches(11.5),Inches(0.9),
    [{'runs':[('Видение.  ',13,True,GREEN),('Интегрированный лидер агро-экспортной логистики и трейдинга Казахстана и региона: от контракта с переработчиком до поставки на рынки Азии, Ближнего Востока и Африки — с цифровым контролем и международным центром в Дубае.',13,False,INK)]}],anchor=MSO_ANCHOR.MIDDLE)
ptable(s,Inches(0.6),Inches(2.75),Inches(12.13),Inches(3.5),
       ['Направление','Сегодня','Цель 2031'],
       [['Трейдинг','масло/зерно через Актау',('сбыт через Дубай на Ближний Восток, Африку, Ю. Азию',True)],
        ['Дубай','—',('действующий центр, оборот $50 млн (год 1) → рост',True)],
        ['Парк пищевых цистерн','20 единиц',('≈ 5 000 под управлением · №1 в РК',True)],
        ['Логистика','мультимодальный парк','развитие вагонов и зерновозов'],
        ['Цифровые решения','внутренняя платформа','контроль всей цепочки; продукт для рынка'],
        ['Операционная маржа','базовая',('рост за счёт интеграции и масштаба',True)]],
       col_w=[Inches(3.5),Inches(3.8),Inches(4.83)],fs=10.5,hfs=10.5)

# ============ 6. БИЗНЕС-МОДЕЛЬ ============
s=slide(); header(s,'Бизнес-модель: вертикальная интеграция',6,kicker='Как это работает')
chain(s,Inches(0.6),Inches(1.5),Inches(12.13),
      ['Закуп','Логистика','Порт Актау / Китай','Дубай — сбыт','Финансирование'],h=Inches(1.0))
card(s,Inches(0.6),Inches(2.95),Inches(3.9),Inches(3.4),'Трейдинг',
     ['Оборот и доступ к марже товара.','Прямые контракты, прозрачная цена.','Тонкая маржа — нужен оборотный капитал.'],accent=COPPER)
card(s,Inches(4.72),Inches(2.95),Inches(3.9),Inches(3.4),'Логистика',
     ['Маржа перевозки и надёжность.','Отгрузки под график судна.','Специализация: цистерны, зерновозы.'],accent=STEEL)
card(s,Inches(8.83),Inches(2.95),Inches(3.9),Inches(3.4),'Синергия КЗ ↔ Дубай',
     ['Казахстан: сырьё и логистика (низкая себестоимость).','Дубай: сбыт, валюта, торговое финансирование.','Цифра связывает дивизионы.'],accent=GREEN,fill=MINT)

# ============ 7. ГДЕ ИГРАЕМ / КАК ПОБЕЖДАЕМ ============
s=slide(); header(s,'Где играем и как побеждаем',7,kicker='Стратегический выбор')
card(s,Inches(0.6),Inches(1.35),Inches(5.95),Inches(5.4),'Где играем',
     [('Товары: ','масло (флагман), зерно, масличные, пищевые наливные, навалочные.'),
      ('База: ','масличный и зерновой пояс Казахстана.'),
      ('Ворота: ','порт Актау (Каспий), Хоргос/Достык (Китай).'),
      ('Рынки: ','Ближний Восток, Африка, Южная Азия — через Дубай.'),
      ('Фокус: ','агро-товарные потоки; не распыляемся на весь генеральный груз.')],accent=NAVY,lsize=11.5)
card(s,Inches(6.78),Inches(1.35),Inches(5.95),Inches(5.4),'Как побеждаем',
     [('Интеграция: ','единый ответственный, маржа на всей цепочке.'),
      ('Парк: ','собственный + привлечённый мультимодальный, 200+ маршрутов.'),
      ('Дубай: ','рынки сбыта + торговое финансирование + валюта.'),
      ('Контракты: ','прямая работа с переработчиками — прозрачная цена.'),
      ('Цифра: ','планирование, прослеживаемость, маржа на рейс/сделку.'),
      ('Специализация: ','пищевые цистерны и зерновозы — барьер и премия.')],accent=GREEN,lsize=11.5,fill=MINT)

# ============ 8. СТОЛПЫ ============
s=slide(); header(s,'Шесть стратегических столпов',8,kicker='Стратегия')
P=[('1 · Трейдинг-экспорт','Масло, зерно, масличные; масштаб контрактов; сбыт через Актау и Дубай.',COPPER),
   ('2 · Дубай — центр','Офис логистика+трейдинг, $50 млн (год 1); рынки и финансирование.',GREEN),
   ('3 · Логистика и парк','Мультимодальный парк; цистерны 20 → 5 000; зерновозы.',STEEL),
   ('4 · Цифровые решения','Платформа планирования и контроля; маржа на рейс/сделку.',NAVY),
   ('5 · Финансовый двигатель','Торговое финансирование, лизинг парка, Банк развития/агро, капитал Дубая.',STEEL),
   ('6 · Люди и управление','Международная команда; мотивация 75/25 и операционная модель.',NAVY)]
cw=Inches(3.9); ch=Inches(2.5)
for i,(tt,bd,ac) in enumerate(P):
    col=i%3; rowi=i//3
    x=Inches(0.6)+col*(cw+Inches(0.21)); y=Inches(1.4)+rowi*(ch+Inches(0.25))
    box(s,x,y,cw,ch,fill=WHITE,line=LINE,rounded=True)
    box(s,x,y,cw,Inches(0.12),fill=ac,rounded=False)
    txt(s,x+Inches(0.24),y+Inches(0.28),cw-Inches(0.46),Inches(0.6),[{'runs':[(tt,13.5,True,ac)]}])
    txt(s,x+Inches(0.24),y+Inches(0.95),cw-Inches(0.46),ch-Inches(1.1),[{'runs':[(bd,11.5,False,INK)],'ls':1.05}])

# ============ 9. ДУБАЙ ============
s=slide(); header(s,'Дубай — ход этого года',9,kicker='Интернационализация')
card(s,Inches(0.6),Inches(1.35),Inches(6.0),Inches(3.05),'Роль центра',
     [('Торговый деск: ','сбыт масла и зерна на Ближний Восток, Африку, Ю. Азию.'),
      ('Финансы: ','аккредитивы, торговое финансирование, валютная выручка.'),
      ('Витрина: ','нейтральная юрисдикция и репутация для контрагентов.')],accent=STEEL,lsize=12)
tiles(s,Inches(6.85),Inches(1.35),Inches(5.88),
      [('$50 млн','Оборот · год 1',COPPER),('КЗ → Дубай','Сырьё и логистика → сбыт и деньги',GREEN)],h=Inches(3.05),gap=Inches(0.2))
box(s,Inches(0.6),Inches(4.65),Inches(12.13),Inches(2.05),fill=RGBColor(0xFB,0xF2,0xE6),line=COPPER,rounded=True)
txt(s,Inches(0.9),Inches(4.85),Inches(11.5),Inches(1.7),
    [{'runs':[('Осторожно (взгляд консультанта).  ',13,True,COPPER)],'sa':4},
     {'runs':[('$50 млн — это оборот, а не прибыль. Маржа трейдинга тонкая (единицы процентов) и требует оборотного капитала и управления риском цены и контрагента. ',12,False,INK)]},
     {'runs':[('Успех Дубая = торговое финансирование + хеджирование + надёжная логистика группы, а не только объём сделок.',12,True,INK)]}],anchor=MSO_ANCHOR.MIDDLE)

# ============ 10. ФИНАНСЫ ============
s=slide(); header(s,'Финансовая рамка',10,kicker='Экономика · иллюстративно')
ptable(s,Inches(0.6),Inches(1.4),Inches(12.13),Inches(3.15),
       ['Оборот, $ млн','Тек.','Год 1','Год 2','Год 3','Год 4','Год 5'],
       [['Логистика','40','55','75','100','130','165'],
        ['Трейдинг (Казахстан)','55','70','95','120','150','185'],
        ['— из них Дубай (новый)','0','50','100','150','210','280'],
        ['Цифровые решения','2','3','5','8','12','18'],
        [('Оборот группы',True),('~97',True),('~178',True),('~275',True),('~378',True),('~502',True),('~648',True)]],
       col_w=[Inches(3.7),Inches(1.4),Inches(1.4),Inches(1.4),Inches(1.4),Inches(1.4),Inches(1.43)],fs=10.5,hfs=10,
       aligns=[PP_ALIGN.LEFT,PP_ALIGN.CENTER,PP_ALIGN.CENTER,PP_ALIGN.CENTER,PP_ALIGN.CENTER,PP_ALIGN.CENTER,PP_ALIGN.CENTER])
box(s,Inches(0.6),Inches(4.8),Inches(12.13),Inches(1.9),fill=MINT,line=LINE,rounded=True)
txt(s,Inches(0.9),Inches(5.0),Inches(11.5),Inches(1.55),
    [{'runs':[('Главный принцип: управляем по валовой прибыли, а не по обороту.  ',13,True,GREEN)],'sa':4},
     {'runs':[('Трейдинг наращивает оборот при тонкой марже; логистика и цифра дают маржу. Числа иллюстративны (заданная точка — Дубай $50 млн в год 1) — заменить на факт: оборот по дивизионам, юнит-экономику цистерны, маржу трейдинга, ставки финансирования.',12,False,INK)]}],anchor=MSO_ANCHOR.MIDDLE)

# ============ 11. ПАРК ============
s=slide(); header(s,'Парк: 20 → 5 000 под управлением',11,kicker='Логистический столп')
tiles(s,Inches(0.6),Inches(1.4),Inches(12.13),
      [('20','Сегодня',NAVY),('450','Год 2',STEEL),('1 300','Год 3',STEEL),('2 800','Год 4',STEEL),('5 000','Год 5',COPPER)],h=Inches(1.3))
ptable(s,Inches(0.6),Inches(3.05),Inches(12.13),Inches(2.1),
       ['Единиц под управлением','Сейчас','Год 1','Год 2','Год 3','Год 4','Год 5'],
       [['Собственные','20','60','180','500','1 000','1 500'],
        ['Лизинг','0','40','150','500','1 200','2 200'],
        ['Партнёрские / управляемые','0','20','120','300','600','1 300'],
        [('Всего парк',True),('20',True),('120',True),('450',True),('1 300',True),('2 800',True),('5 000',True)]],
       col_w=[Inches(3.7),Inches(1.4),Inches(1.4),Inches(1.4),Inches(1.4),Inches(1.4),Inches(1.43)],fs=10.5,hfs=10,
       aligns=[PP_ALIGN.LEFT]+[PP_ALIGN.CENTER]*6)
txt(s,Inches(0.6),Inches(5.45),Inches(12.1),Inches(1.3),
    [{'runs':[('Умный капитал.  ',13,True,COPPER)],'sa':3},
     {'runs':[('5 000 собственных цистерн ≈ $750 млн — нереально «в лоб». Путь: компактное собственное ядро + лизинг + партнёрская сеть + покупка региональных перевозчиков. Это достижимо и капиталоэффективно.',12,False,INK)]}])

# ============ 12. РИСКИ ============
s=slide(); header(s,'Риски и их снятие',12,kicker='Управление рисками')
ptable(s,Inches(0.6),Inches(1.4),Inches(12.13),Inches(5.2),
       ['Риск','Влияние','Как снимаем'],
       [['Оборотный капитал и цена товара (трейдинг)',('Критическое',True),'Аккредитивы, хеджирование, лимиты на сделку и контрагента'],
        ['Запуск Дубая (новый рынок, команда)','Высокое','Опытные трейдеры, поэтапный старт, логистика группы как преимущество'],
        ['Капитал на парк + оборотку','Высокое','Разные источники: лизинг для парка, торговое финансирование для трейдинга'],
        ['Дефицит водителей и инфраструктуры','Высокое','Ранний набор, мойки с опережением, оплата 75/25'],
        ['Узкие места коридора / порт Актау','Среднее','Диверсификация маршрутов, партнёрства с портом'],
        ['Валюта и ставки','Среднее','Валютная выручка Дубая, фиксация лизинга, хеджирование'],
        ['Распыление на 3 дивизиона + новая страна',('Высокое',True),'Приоритеты по горизонтам, единая операционная модель, управление по прибыли']],
       col_w=[Inches(4.6),Inches(1.9),Inches(5.63)],fs=10,hfs=10,
       aligns=[PP_ALIGN.LEFT,PP_ALIGN.CENTER,PP_ALIGN.LEFT])

# ============ 13. ДОРОЖНАЯ КАРТА ============
s=slide(); header(s,'Дорожная карта: три горизонта',13,kicker='План')
H=[('Горизонт 1 · Годы 1–2','«Дубай и ядро»',
    ['Открыть Дубай, оборот $50 млн.','Навести порядок в управлении.','Настроить финансирование и лизинг.','Парк 20 → ~450; запуск платформы.'],STEEL),
   ('Горизонт 2 · Годы 3–4','«Масштаб»',
    ['Масштаб трейдинга и направлений сбыта.','Парк ~450 → ~2 800; сеть моек.','Покупка 1–2 перевозчиков.','Развитие вагонного парка.'],STEEL),
   ('Горизонт 3 · Год 5','«Лидер №1»',
    ['Оборот кратно выше базы; маржа ~12%.','№1 в пищевых цистернах РК.','Дубай — полноценный центр.','Готовность к новому циклу.'],GREEN)]
cw=Inches(3.9)
for i,(tt,sub,lines,ac) in enumerate(H):
    x=Inches(0.6)+i*(cw+Inches(0.21))
    fillc = MINT if ac==GREEN else WHITE
    box(s,x,Inches(1.4),cw,Inches(5.1),fill=fillc,line=LINE,rounded=True)
    box(s,x,Inches(1.4),cw,Inches(0.9),fill=ac,rounded=False)
    txt(s,x+Inches(0.24),Inches(1.5),cw-Inches(0.46),Inches(0.75),
        [{'runs':[(tt,12.5,True,WHITE)]},{'runs':[(sub,11,False,RGBColor(0xEA,0xF1,0xF7))]}],anchor=MSO_ANCHOR.MIDDLE)
    paras=[{'runs':[('▪  ',11,False,ac),(ln,11.5,False,INK)],'sa':7,'ls':1.05} for ln in lines]
    txt(s,x+Inches(0.26),Inches(2.55),cw-Inches(0.5),Inches(3.7),paras)

# ============ 14. РЕШЕНИЕ / КОНТАКТЫ ============
s=slide()
box(s,0,0,W,prs.slide_height,fill=NAVY)
box(s,0,Inches(1.05),W,Inches(0.055),fill=COPPER)
txt(s,Inches(0.9),Inches(0.4),Inches(11.5),Inches(0.6),[{'runs':[('Решение к совету',26,True,WHITE)]}])
asks=[('1','Одобрить запуск международного центра в Дубае (оборот $50 млн, год 1).'),
      ('2','Утвердить программу финансирования: лизинг парка + торговое финансирование трейдинга.'),
      ('3','Назначить владельца выручки и запустить операционную модель управления.'),
      ('4','Утвердить цель «№1 в пищевых цистернах» и траекторию парка 20 → 5 000.')]
y=Inches(1.5)
for nnum,tt in asks:
    box(s,Inches(0.9),y,Inches(0.55),Inches(0.55),fill=COPPER,rounded=True)
    txt(s,Inches(0.9),y,Inches(0.55),Inches(0.55),[{'runs':[(nnum,18,True,WHITE)]}],anchor=MSO_ANCHOR.MIDDLE,align=PP_ALIGN.CENTER)
    txt(s,Inches(1.65),y,Inches(10.9),Inches(0.6),[{'runs':[(tt,14,False,WHITE)]}],anchor=MSO_ANCHOR.MIDDLE)
    y=y+Inches(0.78)
box(s,Inches(0.9),Inches(5.05),Inches(11.5),Pt(1),fill=RGBColor(0x3A,0x52,0x74))
txt(s,Inches(0.9),Inches(5.35),Inches(11.5),Inches(1.6),
    [{'runs':[('ТОО «WayStar Group»',15,True,WHITE)],'sa':6},
     {'runs':[('г. Астана, р-н Есиль, ул. Сығанақ, 60/4, БЦ «Abu Dhabi Plaza»',11,False,MIST)],'sa':3},
     {'runs':[('+7 775 939 98 45   ·   info@waystarco.com   ·   www.waystar.kz',11,False,MIST)]}])

prs.save('strategy/WayStar_Стратегия_презентация.pptx')
print('saved strategy/WayStar_Стратегия_презентация.pptx ·', len(prs.slides.__iter__.__self__._sldIdLst), 'slides')
