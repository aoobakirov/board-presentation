# -*- coding: utf-8 -*-
"""Waystar Group — расчётный файл премий по модели 75/25 (.xlsx), только русские термины,
показатели под транспортно-логистическую компанию.
Оплата = 75% постоянная + 25% переменная. Переменная = (25%*оклад)*выполнение*коэф.компании*стоп-фактор, предел 150%."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

NAVY='13294B'; STEEL='2F6690'; COPPER='C4772F'; LIGHT='F1F5F9'; WHITE='FFFFFF'; GREEN='2E7D46'; YELLOW='FFF3C4'; PROF='E8F3EC'
thin=Side(style='thin',color='C9D3DE'); border=Border(left=thin,right=thin,top=thin,bottom=thin)
def fill(c): return PatternFill('solid',fgColor=c)
HDR=Font(name='Calibri',bold=True,color=WHITE,size=10)
BOLD=Font(name='Calibri',bold=True,size=10,color='2A3746')
REG=Font(name='Calibri',size=10,color='2A3746')
CTR=Alignment(horizontal='center',vertical='center',wrap_text=True)
LEFT=Alignment(horizontal='left',vertical='center',wrap_text=True)
RIGHT=Alignment(horizontal='right',vertical='center')

# №, ФИО, должность, тип(front/back), оклад, [(показатель, вес, прибыль?)]
EMP=[
 (1,'Аманбекова Айгуль','Отдел персонала (руководитель)','back',1246243,
   [('Чистая прибыль компании (общий)',0.25,1),('Удержание персонала / текучесть',0.25,0),
    ('Срок закрытия ключевых вакансий',0.20,0),('Вовлечённость персонала (опрос)',0.15,0),('Внедрение мотивации / бюджет фонда',0.15,0)]),
 (2,'Аубакиров Алимбек','Финансовый директор','front',4347889,
   [('Чистая прибыль компании к плану',0.35,1),('Запас денежных средств',0.20,0),('Оборот дебиторской задолженности',0.15,0),
    ('Точность бюджета',0.15,0),('Привлечение финансирования',0.15,0)]),
 (3,'Байльденова Шолпан','Менеджер по персоналу','back',614931,
   [('Чистая прибыль компании (общий)',0.25,1),('Срок закрытия вакансий',0.20,0),('Кадровый учёт без нарушений',0.20,0),
    ('Точность расчёта зарплаты',0.20,0),('Адаптация новых сотрудников',0.15,0)]),
 (4,'Данияр Алим','Генеральный директор','front',1246243,
   [('Чистая прибыль компании к плану',0.40,1),('Валовая прибыль / рентабельность',0.20,1),('Выручка к плану',0.20,0),
    ('Выполнение годового плана',0.15,0),('Удержание клиентов',0.05,0)]),
 (5,'Доолоткулов Данияр','Начальник отдела логистики','front',1246243,
   [('Валовая прибыль отдела перевозок к плану',0.40,1),('Соблюдение сроков доставки',0.20,0),
    ('Себестоимость перевозки на рейс/тонну',0.20,0),('Загрузка транспорта / порожний пробег',0.10,0),('Сохранность груза / претензии',0.10,0)]),
 (6,'Жунусова Камила','Менеджер по логистике','front',230896,
   [('Валовая маржа по своим перевозкам (≥K×стоимости)',0.45,1),('Объём перевозок к норме',0.15,0),
    ('Соблюдение сроков доставки',0.15,0),('Загрузка / порожний пробег',0.15,0),('Точность документов / сохранность',0.10,0)]),
 (7,'Сейтқадыр Дарын','Менеджер по логистике','front',236143,
   [('Валовая маржа по своим перевозкам (≥K×стоимости)',0.45,1),('Объём перевозок к норме',0.15,0),
    ('Соблюдение сроков доставки',0.15,0),('Загрузка / порожний пробег',0.15,0),('Точность документов / сохранность',0.10,0)]),
 (8,'Серикбаев Адиль','Президент','front',6000139,
   [('Чистая прибыль компании к плану',0.50,1),('Выручка к плану',0.15,0),('Развитие: направления/клиенты/география',0.20,0),
    ('Денежный поток / финансирование',0.10,0),('Удержание ключевой команды',0.05,0)]),
 (9,'Таджимурадова Асем','Менеджер по закрытию','front',378788,
   [('Валовая прибыль по закрытым перевозкам к плану',0.45,1),('Собираемость оплаты в срок',0.30,0),
    ('Срок закрытия документов',0.15,0),('Точность / отсутствие расхождений',0.10,0)]),
 (10,'Ташметов Сакен','Юрист','back',1246243,
   [('Чистая прибыль компании (общий)',0.25,1),('Взысканная дебиторка / споры',0.25,0),('Согласование договоров в срок',0.20,0),
    ('Юридические риски: штрафы',0.15,0),('Сопровождение сделок и перевозок',0.15,0)]),
 (11,'Унбетпаев Ернар','Операционный директор','front',1873611,
   [('Валовая прибыль по операциям к плану',0.40,1),('Соблюдение сроков (в срок и полностью)',0.20,0),
    ('Себестоимость операций на тонну/км',0.20,0),('Безопасность / охрана труда',0.10,0),('Загрузка / порожний пробег',0.10,0)]),
 (12,'Белова Полина','Менеджер по логистике (младший)','front',173011,
   [('Валовая маржа по своим перевозкам (≥K×стоимости)',0.45,1),('Объём перевозок к норме',0.15,0),
    ('Соблюдение сроков доставки',0.15,0),('Загрузка / порожний пробег',0.15,0),('Точность документов / сохранность',0.10,0)]),
 (13,'Сибагатов Хамардин','Менеджер по логистике (старший)','front',993718,
   [('Валовая маржа по своим перевозкам (≥K×стоимости)',0.45,1),('Объём перевозок к норме',0.15,0),
    ('Соблюдение сроков доставки',0.15,0),('Загрузка / порожний пробег',0.15,0),('Точность документов / сохранность',0.10,0)]),
]
FIX=0.75; VAR=0.25; CAP=1.5

wb=Workbook()

# ---------- Лист: Показатели ----------
det=wb.active; det.title='Показатели'; det.sheet_view.showGridLines=False
det['A1']='Waystar Group · Показатели эффективности (модель 75/25)'
det['A1'].font=Font(name='Calibri',bold=True,size=14,color=NAVY)
det['A2']='Введите фактическое выполнение по каждому показателю (жёлтые ячейки, 0–150%). Зелёные строки — цель по прибыли '\
          '(валовая/чистая). Итоговый процент выполнения = сумма(вес×выполнение) переносится в лист «Расчёт».'
det['A2'].font=Font(name='Calibri',size=9,italic=True,color='6B7A8D'); det.merge_cells('A2:F2')
dv=DataValidation(type='decimal',operator='between',formula1=0,formula2=1.5,allow_blank=True)
dv.error='Выполнение 0–150%'; dv.errorTitle='Диапазон'; det.add_data_validation(dv)
for col,w in {'A':4,'B':48,'C':10,'D':16,'E':16,'F':4}.items(): det.column_dimensions[col].width=w

subtotal_ref={}
r=4
for (idx,name,role,typ,oklad,kpis) in EMP:
    tag='ПЕРЕДОВАЯ ЛИНИЯ' if typ=='front' else 'ПОДДЕРЖКА'
    det.cell(r,2,'%d. %s — %s  [%s]'%(idx,name,role,tag))
    for col in range(1,7):
        det.cell(r,col).fill=fill(STEEL); det.cell(r,col).border=border
        det.cell(r,col).font=Font(name='Calibri',bold=True,size=10,color=WHITE)
    r+=1
    det.cell(r,2,'Показатель').font=HDR; det.cell(r,3,'Вес').font=HDR
    det.cell(r,4,'Выполнение, %').font=HDR; det.cell(r,5,'Вклад').font=HDR
    for col in range(1,7):
        det.cell(r,col).fill=fill(NAVY); det.cell(r,col).border=border; det.cell(r,col).alignment=CTR
    r+=1; first=r
    for (kpi,w,isp) in kpis:
        det.cell(r,2,kpi).font=(Font(name='Calibri',bold=True,size=10,color=GREEN) if isp else REG); det.cell(r,2).alignment=LEFT
        wc=det.cell(r,3,w); wc.number_format='0%'; wc.font=REG; wc.alignment=CTR
        ac=det.cell(r,4,1.0); ac.number_format='0%'; ac.font=Font(name='Calibri',size=10,color='1F4E79',bold=True)
        ac.fill=fill(YELLOW); ac.alignment=CTR; dv.add(ac)
        cc=det.cell(r,5,'=C%d*D%d'%(r,r)); cc.number_format='0%'; cc.font=REG; cc.alignment=CTR
        for col in range(1,7):
            det.cell(r,col).border=border
            if isp and col in (2,3,5): det.cell(r,col).fill=fill(PROF)
        r+=1
    last=r-1
    det.cell(r,2,'Итоговое выполнение').font=BOLD
    wsum=det.cell(r,3,'=SUM(C%d:C%d)'%(first,last)); wsum.number_format='0%'; wsum.font=BOLD; wsum.alignment=CTR
    tcell=det.cell(r,5,'=SUM(E%d:E%d)'%(first,last)); tcell.number_format='0%'
    tcell.font=Font(name='Calibri',bold=True,size=10,color=GREEN); tcell.alignment=CTR
    for col in range(1,7): det.cell(r,col).fill=fill(LIGHT); det.cell(r,col).border=border
    subtotal_ref[idx]=r
    r+=2
det.freeze_panes='A4'

# ---------- Лист: Расчёт ----------
calc=wb.create_sheet('Расчёт',0); calc.sheet_view.showGridLines=False
calc['A1']='Waystar Group · Расчёт оплаты 75/25 по показателям'
calc['A1'].font=Font(name='Calibri',bold=True,size=15,color=NAVY)
calc['A2']='Итого = 75%×Оклад + (25%×Оклад)×Выполнение×Коэффициент компании×Стоп-фактор   (переменная — предел 150%)'
calc['A2'].font=Font(name='Calibri',size=10,italic=True,color=COPPER)
calc['A4']='Коэффициент компании (по чистой прибыли, 0,7–1,3):'; calc['A4'].font=BOLD
mod=calc['F4']; mod.value=1.0; mod.number_format='0.00'; mod.font=Font(bold=True,color='1F4E79',size=11)
mod.fill=fill(YELLOW); mod.alignment=CTR; mod.border=border
calc['G4']='← 100% плана прибыли = 1,00'; calc['G4'].font=Font(size=9,italic=True,color='6B7A8D')
dvm=DataValidation(type='decimal',operator='between',formula1=0,formula2=1.3); calc.add_data_validation(dvm); dvm.add(mod)

heads=['№','Сотрудник','Должность','Линия','Оклад (100%), ₸','Постоянная 75%, ₸','Переменная целевая 25%, ₸',
       'Выполнение','Стоп-фактор','Переменная к выплате, ₸','ИТОГО к выплате, ₸','Разница к текущему, ₸']
hr=6
for j,txt in enumerate(heads,1):
    c=calc.cell(hr,j,txt); c.font=HDR; c.fill=fill(NAVY); c.border=border; c.alignment=CTR
widths=[4,24,27,8,14,14,16,10,9,16,16,15]
for j,w in enumerate(widths,1): calc.column_dimensions[get_column_letter(j)].width=w
dvg=DataValidation(type='list',formula1='"1,0.5,0"'); calc.add_data_validation(dvg)

row=hr+1; first_data=row
for (idx,name,role,typ,oklad,kpis) in EMP:
    calc.cell(row,1,idx).font=REG; calc.cell(row,1).alignment=CTR
    calc.cell(row,2,name).font=REG; calc.cell(row,2).alignment=LEFT
    calc.cell(row,3,role).font=REG; calc.cell(row,3).alignment=LEFT
    tg=calc.cell(row,4,'Передовая' if typ=='front' else 'Поддержка'); tg.alignment=CTR
    tg.font=Font(name='Calibri',bold=True,size=9,color=(GREEN if typ=='front' else STEEL))
    oc=calc.cell(row,5,oklad); oc.number_format='#,##0'; oc.font=REG; oc.alignment=RIGHT
    fx=calc.cell(row,6,'=E%d*%s'%(row,FIX)); fx.number_format='#,##0'; fx.font=REG; fx.alignment=RIGHT
    vt=calc.cell(row,7,'=E%d*%s'%(row,VAR)); vt.number_format='#,##0'; vt.font=REG; vt.alignment=RIGHT
    kc=calc.cell(row,8,"=Показатели!E%d"%subtotal_ref[idx]); kc.number_format='0%'
    kc.font=Font(name='Calibri',size=10,color=GREEN,bold=True); kc.alignment=CTR
    gc=calc.cell(row,9,1); gc.number_format='0.0'; gc.font=Font(color='1F4E79',bold=True); gc.alignment=CTR
    gc.fill=fill(YELLOW); dvg.add(gc)
    ve=calc.cell(row,10,'=G%d*MIN(%s,H%d*$F$4*I%d)'%(row,CAP,row,row)); ve.number_format='#,##0'; ve.font=BOLD; ve.alignment=RIGHT
    tot=calc.cell(row,11,'=F%d+J%d'%(row,row)); tot.number_format='#,##0'; tot.font=BOLD; tot.alignment=RIGHT
    dl=calc.cell(row,12,'=K%d-E%d'%(row,row)); dl.number_format='#,##0;[Red]-#,##0'; dl.font=REG; dl.alignment=RIGHT
    zb=LIGHT if (row-first_data)%2==0 else WHITE
    for j in range(1,13):
        calc.cell(row,j).border=border
        if j not in (8,9): calc.cell(row,j).fill=fill(zb)
    row+=1
last_data=row-1
calc.cell(row,2,'ИТОГО').font=Font(bold=True,color=WHITE,size=10)
for col,f in [(5,'E'),(6,'F'),(7,'G'),(10,'J'),(11,'K'),(12,'L')]:
    c=calc.cell(row,col,'=SUM(%s%d:%s%d)'%(f,first_data,f,last_data)); c.number_format='#,##0'
for j in range(1,13):
    calc.cell(row,j).fill=fill(STEEL); calc.cell(row,j).border=border
    calc.cell(row,j).font=Font(bold=True,color=WHITE,size=10)
    calc.cell(row,j).alignment=RIGHT if j in (5,6,7,10,11,12) else CTR
trow=row
row+=2
calc.cell(row,2,'Фонд (оклады) в год, ₸:').font=BOLD
calc.cell(row,5,'=E%d*12'%trow).number_format='#,##0'; calc.cell(row,5).font=Font(bold=True,color=NAVY); calc.cell(row,5).alignment=RIGHT
calc.cell(row,10,'Под результатом (25%) в год:').font=BOLD
calc.cell(row,11,'=G%d*12'%trow).number_format='#,##0'; calc.cell(row,11).font=Font(bold=True,color=COPPER); calc.cell(row,11).alignment=RIGHT
row+=1
calc.cell(row,2,'Итого к выплате в год, ₸:').font=BOLD
calc.cell(row,11,'=K%d*12'%trow).number_format='#,##0'; calc.cell(row,11).font=Font(bold=True,color=GREEN,size=11); calc.cell(row,11).alignment=RIGHT
row+=2
for note in [
 'Жёлтые ячейки — вводимые: Коэффициент компании (F4), Стоп-фактор по строке, выполнение показателей на листе «Показатели».',
 'Передовая линия — есть личная/командная цель по прибыли (валовая/чистая); Поддержка — доля переменной на чистой прибыли компании.',
 'Переменная = (25%×оклад) × выполнение × коэффициент компании × стоп-фактор, предел 150% целевой. При 100% и коэффициенте 1,0 «Итого» = текущему окладу.',
 'Оклад — начисленный (база деления 75/25). Пороги и коэффициент K целей по прибыли настраиваются по управленческой отчётности.']:
    calc.cell(row,2,'• '+note).font=Font(size=9,italic=True,color='6B7A8D')
    calc.merge_cells(start_row=row,start_column=2,end_row=row,end_column=12); row+=1

calc.freeze_panes='A7'
wb.save('motivation/Waystar_Калькулятор_премий.xlsx')
print('saved motivation/Waystar_Калькулятор_премий.xlsx')
