# -*- coding: utf-8 -*-
"""Waystar Group — interactive KPI premium calculator (.xlsx).
Sheet 1 'Калькулятор' — one row per employee, premium = f(KPI%, company modifier, gate).
Sheet 2 'Детализация KPI' — per-employee KPI blocks that roll up into the KPI% used on sheet 1."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

NAVY='13294B'; STEEL='2F6690'; COPPER='C4772F'; LIGHT='F1F5F9'; WHITE='FFFFFF'; GREEN='2E7D46'; YELLOW='FFF3C4'
thin=Side(style='thin',color='C9D3DE')
border=Border(left=thin,right=thin,top=thin,bottom=thin)
def fill(c): return PatternFill('solid',fgColor=c)
HDR=Font(name='Calibri',bold=True,color=WHITE,size=10)
BOLD=Font(name='Calibri',bold=True,size=10,color='2A3746')
REG=Font(name='Calibri',size=10,color='2A3746')
CTR=Alignment(horizontal='center',vertical='center',wrap_text=True)
LEFT=Alignment(horizontal='left',vertical='center',wrap_text=True)
RIGHT=Alignment(horizontal='right',vertical='center')

# employee: idx, name, role, grade, oklad_gross, target_coef, [(kpi, weight)]
EMP=[
 (1,'Аманбекова Айгуль','HR (директор по персоналу)','B',1246243,0.30,
   [('Закрытие ключевых вакансий в срок',0.25),('Удержание персонала / текучесть',0.25),
    ('eNPS / вовлечённость',0.20),('Запуск KPI-системы + обучение',0.20),('ФОТ-бюджет / точность payroll',0.10)]),
 (2,'Аубакиров Алимбек','CFO','A',4347889,0.40,
   [('Точность бюджета',0.20),('Ликвидность / runway',0.20),('DSO — дебиторка',0.20),
    ('Привлечение финансирования',0.20),('Закрытие периода + аудит',0.20)]),
 (3,'Байльденова Шолпан','HR менеджер','C',614931,0.20,
   [('Скорость подбора (time-to-hire)',0.25),('Кадровый документооборот',0.25),
    ('Точность payroll/табеля',0.20),('Онбординг / испыт. срок',0.15),('Поддержка HR-инициатив',0.15)]),
 (4,'Данияр Алим','CEO','A',1246243,0.40,
   [('Выручка vs план',0.25),('Операционная маржа / EBITDA',0.25),('Бизнес-план / OKR',0.20),
    ('Денежный цикл (CCC)',0.15),('Кросс-функц. цель / eNPS',0.15)]),
 (5,'Доолоткулов Данияр','Начальник отдела логистики','B',1246243,0.30,
   [('On-time delivery',0.30),('Стоимость доставки на тонну',0.25),('Загрузка транспорта',0.20),
    ('Потери / претензии',0.15),('План отдела / команда',0.10)]),
 (6,'Жунусова Камила','Менеджер по логистике','D',230896,0.15,
   [('Обработанные заявки/рейсы',0.25),('Экономия на фрахте / маржа',0.25),('On-time %',0.20),
    ('Точность документов',0.15),('Претензии / NPS',0.15)]),
 (7,'Сейтқадыр Дарын','Менеджер по логистике','D',236143,0.15,
   [('Обработанные заявки/рейсы',0.25),('Экономия на фрахте / маржа',0.25),('On-time %',0.20),
    ('Точность документов',0.15),('Претензии / NPS',0.15)]),
 (8,'Серикбаев Адиль','Президент','A',6000139,0.40,
   [('Выручка компании vs план',0.25),('EBITDA / чистая прибыль',0.25),('Стратег. инициативы',0.25),
    ('Финансирование / ДДС',0.15),('Удержание топ-команды',0.10)]),
 (9,'Таджимурадова Асем','Менеджер по закрытию','E',378788,0.35,
   [('Закрытые сделки/акты vs план',0.30),('Собираемость дебиторки',0.30),
    ('Срок закрытия документов',0.20),('Точность / 0 расхождений',0.20)]),
 (10,'Ташметов Сакен','Юрист','C',1246243,0.25,
   [('Договоры в SLA',0.25),('Взысканная дебиторка / споры',0.25),('Юр. риски: штрафы',0.20),
    ('Договорная защита / потери',0.15),('Поддержка сделок (DD)',0.15)]),
 (11,'Унбетпаев Ернар','Операционный директор','A',1873611,0.40,
   [('OTIF',0.25),('Удельная себестоимость операций',0.25),('Утилизация мощностей',0.20),
    ('Оборачиваемость запасов',0.15),('HSE / качество',0.15)]),
 (12,'Белова Полина','Менеджер по логистике (junior)','D',173011,0.15,
   [('Обработанные заявки/рейсы',0.25),('Экономия на фрахте / маржа',0.25),('On-time %',0.20),
    ('Точность документов',0.15),('Претензии / NPS',0.15)]),
 (13,'Сибагатов Хамардин','Менеджер по логистике (senior)','D',993718,0.20,
   [('Обработанные заявки/рейсы',0.25),('Экономия на фрахте / маржа',0.25),('On-time %',0.20),
    ('Точность документов',0.15),('Претензии / NPS',0.15)]),
]

wb=Workbook()

# ---------- Sheet 2 first (so sheet1 can reference); we'll reorder ----------
det=wb.active; det.title='Детализация KPI'
det.sheet_view.showGridLines=False
det['A1']='Waystar Group · Детализация KPI по сотрудникам'
det['A1'].font=Font(name='Calibri',bold=True,size=14,color=NAVY)
det['A2']='Введите фактическое достижение по каждому KPI в колонке «Достижение, %» (0–150%). '\
          'Итоговый % KPI = Σ(вес × достижение) автоматически переносится в лист «Калькулятор».'
det['A2'].font=Font(name='Calibri',size=9,italic=True,color='6B7A8D')
det.merge_cells('A2:F2')

dv=DataValidation(type='decimal',operator='between',formula1=0,formula2=1.5,allow_blank=True)
dv.error='Введите достижение от 0% до 150%'; dv.errorTitle='Диапазон'
det.add_data_validation(dv)

subtotal_ref={}  # idx -> cell ref of Итоговый %KPI on det sheet
r=4
det_widths={'A':4,'B':40,'C':10,'D':16,'E':16,'F':13}
for col,w in det_widths.items(): det.column_dimensions[col].width=w
for (idx,name,role,grade,oklad,tc,kpis) in EMP:
    det.cell(r,1,'%d'%idx).font=BOLD
    c=det.cell(r,2,'%s — %s'%(name,role)); c.font=Font(name='Calibri',bold=True,size=10,color=WHITE)
    for col in range(1,7):
        det.cell(r,col).fill=fill(STEEL); det.cell(r,col).border=border
        if col>1: det.cell(r,col).font=Font(name='Calibri',bold=True,size=10,color=WHITE)
    det.cell(r,1).font=Font(name='Calibri',bold=True,size=10,color=WHITE)
    r+=1
    hdr_row=r
    for col,txt in zip('BCDEF',['KPI','Вес','Достижение, %','Вклад','']):
        pass
    det.cell(r,2,'KPI').font=HDR; det.cell(r,2).fill=fill(NAVY)
    det.cell(r,3,'Вес').font=HDR; det.cell(r,3).fill=fill(NAVY)
    det.cell(r,4,'Достижение, %').font=HDR; det.cell(r,4).fill=fill(NAVY)
    det.cell(r,5,'Вклад (вес×дост.)').font=HDR; det.cell(r,5).fill=fill(NAVY)
    for col in (1,6): det.cell(r,col).fill=fill(NAVY)
    for col in range(1,7):
        det.cell(r,col).border=border; det.cell(r,col).alignment=CTR
    r+=1
    first=r
    for (kpi,w) in kpis:
        det.cell(r,2,kpi).font=REG; det.cell(r,2).alignment=LEFT
        wc=det.cell(r,3,w); wc.number_format='0%'; wc.font=REG; wc.alignment=CTR
        ac=det.cell(r,4,1.0); ac.number_format='0%'; ac.font=Font(name='Calibri',size=10,color='1F4E79',bold=True)
        ac.fill=fill(YELLOW); ac.alignment=CTR; dv.add(ac)
        cc=det.cell(r,5,'=C%d*D%d'%(r,r)); cc.number_format='0%'; cc.font=REG; cc.alignment=CTR
        for col in range(1,7): det.cell(r,col).border=border
        r+=1
    last=r-1
    det.cell(r,2,'Итоговый % KPI').font=BOLD
    tcell=det.cell(r,5,'=SUM(E%d:E%d)'%(first,last)); tcell.number_format='0%'
    tcell.font=Font(name='Calibri',bold=True,size=10,color=GREEN); tcell.alignment=CTR
    wsum=det.cell(r,3,'=SUM(C%d:C%d)'%(first,last)); wsum.number_format='0%'; wsum.font=BOLD; wsum.alignment=CTR
    for col in range(1,7):
        det.cell(r,col).fill=fill(LIGHT); det.cell(r,col).border=border
    subtotal_ref[idx]=r
    r+=2

# ---------- Sheet 1: Калькулятор ----------
calc=wb.create_sheet('Калькулятор',0)
calc.sheet_view.showGridLines=False
calc['A1']='Waystar Group · Калькулятор премий по KPI'
calc['A1'].font=Font(name='Calibri',bold=True,size=15,color=NAVY)
calc['A2']='Премия = Оклад × Целевой коэф. × Итоговый %KPI × Модификатор компании × Гейт'
calc['A2'].font=Font(name='Calibri',size=10,italic=True,color=COPPER)

# global inputs
calc['A4']='Модификатор компании (0,7–1,3):'; calc['A4'].font=BOLD
mod=calc['E4']; mod.value=1.0; mod.number_format='0.00'; mod.font=Font(bold=True,color='1F4E79',size=11)
mod.fill=fill(YELLOW); mod.alignment=CTR; mod.border=border
calc['F4']='← результат компании к плану (100%→1,00)'; calc['F4'].font=Font(size=9,italic=True,color='6B7A8D')
dvm=DataValidation(type='decimal',operator='between',formula1=0,formula2=1.3,allow_blank=False)
calc.add_data_validation(dvm); dvm.add(mod)

heads=['№','Сотрудник','Должность','Грейд','Оклад (мес.), ₸','Целевой коэф.',
       'Итоговый %KPI','Гейт','Премия (мес.), ₸','Совокупно (оклад+премия), ₸']
hr=6
for j,txt in enumerate(heads,1):
    c=calc.cell(hr,j,txt); c.font=HDR; c.fill=fill(NAVY); c.border=border; c.alignment=CTR
widths=[4,26,30,7,15,11,12,8,16,20]
for j,w in enumerate(widths,1): calc.column_dimensions[get_column_letter(j)].width=w

dvg=DataValidation(type='list',formula1='"1,0.5,0"',allow_blank=False)
calc.add_data_validation(dvg)

row=hr+1
first_data=row
for (idx,name,role,grade,oklad,tc,kpis) in EMP:
    calc.cell(row,1,idx).font=REG; calc.cell(row,1).alignment=CTR
    calc.cell(row,2,name).font=REG; calc.cell(row,2).alignment=LEFT
    calc.cell(row,3,role).font=REG; calc.cell(row,3).alignment=LEFT
    calc.cell(row,4,grade).font=BOLD; calc.cell(row,4).alignment=CTR
    oc=calc.cell(row,5,oklad); oc.number_format='#,##0'; oc.font=REG; oc.alignment=RIGHT
    cc=calc.cell(row,6,tc); cc.number_format='0%'; cc.font=REG; cc.alignment=CTR
    # KPI% pulled from detalization sheet
    kc=calc.cell(row,7,"='Детализация KPI'!E%d"%subtotal_ref[idx]); kc.number_format='0%'
    kc.font=Font(name='Calibri',size=10,color=GREEN,bold=True); kc.alignment=CTR
    gc=calc.cell(row,8,1); gc.number_format='0.0'; gc.font=Font(color='1F4E79',bold=True); gc.alignment=CTR
    gc.fill=fill(YELLOW); dvg.add(gc)
    pc=calc.cell(row,9,'=E%d*F%d*G%d*$E$4*H%d'%(row,row,row,row)); pc.number_format='#,##0'
    pc.font=BOLD; pc.alignment=RIGHT
    tc2=calc.cell(row,10,'=E%d+I%d'%(row,row)); tc2.number_format='#,##0'; tc2.font=REG; tc2.alignment=RIGHT
    zb= LIGHT if (row-first_data)%2==0 else WHITE
    for j in range(1,11):
        calc.cell(row,j).border=border
        if j not in (7,8):
            if calc.cell(row,j).fill.fgColor.rgb in (None,'00000000'): calc.cell(row,j).fill=fill(zb)
    row+=1
last_data=row-1
# totals
calc.cell(row,2,'ИТОГО').font=Font(bold=True,color=WHITE,size=10)
tsum=calc.cell(row,5,'=SUM(E%d:E%d)'%(first_data,last_data)); tsum.number_format='#,##0'
psum=calc.cell(row,9,'=SUM(I%d:I%d)'%(first_data,last_data)); psum.number_format='#,##0'
csum=calc.cell(row,10,'=SUM(J%d:J%d)'%(first_data,last_data)); csum.number_format='#,##0'
for j in range(1,11):
    calc.cell(row,j).fill=fill(STEEL); calc.cell(row,j).border=border
    calc.cell(row,j).font=Font(bold=True,color=WHITE,size=10)
    calc.cell(row,j).alignment=RIGHT if j in (5,9,10) else CTR
trow=row
# annualized note
row+=2
calc.cell(row,2,'Премиальный фонд в год (×12):').font=BOLD
ann=calc.cell(row,9,'=I%d*12'%trow); ann.number_format='#,##0'; ann.font=Font(bold=True,color=COPPER,size=11); ann.alignment=RIGHT
row+=1
calc.cell(row,2,'Совокупный ФОТ+премии в год (×12):').font=BOLD
ann2=calc.cell(row,10,'=J%d*12'%trow); ann2.number_format='#,##0'; ann2.font=Font(bold=True,color=NAVY,size=11); ann2.alignment=RIGHT
row+=2
for note in [
 'Жёлтые ячейки — вводимые: Модификатор компании (E4), Гейт по строке, и достижения KPI на листе «Детализация KPI».',
 'Гейт: 1,0 — нарушений нет; 0,5 — существенное нарушение; 0 — критическое (обнуление премии).',
 'Оклад — начисленный (база премирования). Пороги/цели KPI калибруются владельцами функций.',
 'Итоговый %KPI подтягивается с листа «Детализация KPI» (Σ вес×достижение).']:
    calc.cell(row,2,'• '+note).font=Font(size=9,italic=True,color='6B7A8D')
    calc.merge_cells(start_row=row,start_column=2,end_row=row,end_column=10)
    row+=1

calc.freeze_panes='A7'
det.freeze_panes='A4'
wb.save('motivation/Waystar_KPI_premium_calculator.xlsx')
print('saved motivation/Waystar_KPI_premium_calculator.xlsx')
