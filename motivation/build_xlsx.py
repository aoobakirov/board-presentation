# -*- coding: utf-8 -*-
"""ТОО «WayStar Group» — Матрица оплаты 75/25 и показатели эффективности (.xlsx).
Встроено в структуру пользовательского файла: листы «Матрица 75-25», «Показатели (детально)»,
«Справочник показателей». Только русские термины; показатели — под транспортную логистику.
По приказу: премия 0–100% лимита (25%). Начисленная премия = Лимит × мин(100%; факт показателей)."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

NAVY='13294B'; STEEL='2F6690'; COPPER='C4772F'; LIGHT='F1F5F9'; WHITE='FFFFFF'; GREEN='2E7D46'; YELLOW='FFF3C4'; PROF='E8F3EC'
thin=Side(style='thin',color='C9D3DE'); border=Border(left=thin,right=thin,top=thin,bottom=thin)
def fill(c): return PatternFill('solid',fgColor=c)
HDR=Font(name='Calibri',bold=True,color=WHITE,size=9.5)
BOLD=Font(name='Calibri',bold=True,size=10,color='2A3746')
REG=Font(name='Calibri',size=10,color='2A3746')
CTR=Alignment(horizontal='center',vertical='center',wrap_text=True)
LEFT=Alignment(horizontal='left',vertical='center',wrap_text=True)
RIGHT=Alignment(horizontal='right',vertical='center')
DATE='01.01.2026'

# №, ФИО, подразделение, должность, тип(front/back), оклад(ЦЕД 100%), примечание, [(показатель, вес, прибыль?)]
EMP=[
 (8,'Серикбаев Адиль','Руководство','Президент','front',6000139,'Отвечает за весь результат',
   [('Чистая прибыль компании к плану',0.50,1),('Выручка к плану',0.15,0),('Развитие: направления/клиенты/география',0.20,0),
    ('Денежный поток / финансирование',0.10,0),('Удержание ключевой команды',0.05,0)]),
 (4,'Данияр Алим','Руководство','Генеральный директор','front',1246243,'Операционный результат',
   [('Чистая прибыль компании к плану',0.40,1),('Валовая прибыль / рентабельность',0.20,1),('Выручка к плану',0.20,0),
    ('Выполнение годового плана',0.15,0),('Удержание клиентов',0.05,0)]),
 (2,'Аубакиров Алимбек','Финансы','Финансовый директор','front',4347889,'Прибыль и денежная устойчивость',
   [('Чистая прибыль компании к плану',0.35,1),('Запас денежных средств',0.20,0),('Оборот дебиторской задолженности',0.15,0),
    ('Точность бюджета',0.15,0),('Привлечение финансирования',0.15,0)]),
 (11,'Унбетпаев Ернар','Операции','Операционный директор','front',1873611,'Валовая прибыль по операциям ≥ K× стоимости',
   [('Валовая прибыль по операциям к плану',0.40,1),('Соблюдение сроков (в срок и полностью)',0.20,0),
    ('Себестоимость операций на тонну/км',0.20,0),('Безопасность / охрана труда',0.10,0),('Загрузка / порожний пробег',0.10,0)]),
 (5,'Доолоткулов Данияр','Логистика','Начальник отдела логистики','front',1246243,'Валовая прибыль отдела ≥ K× стоимости',
   [('Валовая прибыль отдела перевозок к плану',0.40,1),('Соблюдение сроков доставки',0.20,0),
    ('Себестоимость перевозки на рейс/тонну',0.20,0),('Загрузка / порожний пробег',0.10,0),('Сохранность груза / претензии',0.10,0)]),
 (13,'Сибагатов Хамардин','Логистика','Менеджер по логистике (старший)','front',993718,'Маржа ≥ K× стоимости; наставничество',
   [('Валовая маржа по своим перевозкам (≥K×стоимости)',0.45,1),('Объём перевозок к норме',0.15,0),
    ('Соблюдение сроков доставки',0.15,0),('Загрузка / порожний пробег',0.15,0),('Точность документов / сохранность',0.10,0)]),
 (6,'Жунусова Камила','Логистика','Менеджер по логистике','front',230896,'Маржа ≥ K× стоимости',
   [('Валовая маржа по своим перевозкам (≥K×стоимости)',0.45,1),('Объём перевозок к норме',0.15,0),
    ('Соблюдение сроков доставки',0.15,0),('Загрузка / порожний пробег',0.15,0),('Точность документов / сохранность',0.10,0)]),
 (7,'Сейтқадыр Дарын','Логистика','Менеджер по логистике','front',236143,'Маржа ≥ K× стоимости',
   [('Валовая маржа по своим перевозкам (≥K×стоимости)',0.45,1),('Объём перевозок к норме',0.15,0),
    ('Соблюдение сроков доставки',0.15,0),('Загрузка / порожний пробег',0.15,0),('Точность документов / сохранность',0.10,0)]),
 (12,'Белова Полина','Логистика','Менеджер по логистике (младший)','front',173011,'Маржа ≥ K× стоимости; пороги ниже',
   [('Валовая маржа по своим перевозкам (≥K×стоимости)',0.45,1),('Объём перевозок к норме',0.15,0),
    ('Соблюдение сроков доставки',0.15,0),('Загрузка / порожний пробег',0.15,0),('Точность документов / сохранность',0.10,0)]),
 (9,'Таджимурадова Асем','Расчёты','Менеджер по закрытию','front',378788,'Прибыль по закрытым перевозкам + собранная оплата',
   [('Валовая прибыль по закрытым перевозкам к плану',0.45,1),('Собираемость оплаты в срок',0.30,0),
    ('Срок закрытия документов',0.15,0),('Точность / отсутствие расхождений',0.10,0)]),
 (1,'Аманбекова Айгуль','Персонал','Руководитель отдела персонала','back',1246243,'Поддержка · доля от прибыли компании',
   [('Чистая прибыль компании (общий)',0.25,1),('Удержание персонала / текучесть',0.25,0),
    ('Срок закрытия ключевых вакансий',0.20,0),('Вовлечённость персонала (опрос)',0.15,0),('Внедрение мотивации / бюджет фонда',0.15,0)]),
 (3,'Байльденова Шолпан','Персонал','Менеджер по персоналу','back',614931,'Поддержка · доля от прибыли компании',
   [('Чистая прибыль компании (общий)',0.25,1),('Срок закрытия вакансий',0.20,0),('Кадровый учёт без нарушений',0.20,0),
    ('Точность расчёта зарплаты',0.20,0),('Адаптация новых сотрудников',0.15,0)]),
 (10,'Ташметов Сакен','Юридический','Юрист','back',1246243,'Поддержка · доля от прибыли компании',
   [('Чистая прибыль компании (общий)',0.25,1),('Взысканная дебиторка / споры',0.25,0),('Согласование договоров в срок',0.20,0),
    ('Юридические риски: штрафы',0.15,0),('Сопровождение сделок и перевозок',0.15,0)]),
]
FIX=0.75; VAR=0.25

wb=Workbook()

# =========== ЛИСТ 3: Показатели (детально) — строим первым для ссылок ===========
det=wb.active; det.title='Показатели (детально)'; det.sheet_view.showGridLines=False
det['A1']='ТОО «WayStar Group» · Показатели эффективности по сотрудникам (модель 75/25)'
det['A1'].font=Font(name='Calibri',bold=True,size=14,color=NAVY)
det['A2']='Введите фактическое выполнение по каждому показателю (жёлтые ячейки, 0–150%). Зелёные строки — цель по прибыли. '\
          'Итоговое выполнение = сумма(вес×выполнение) переносится в лист «Матрица 75-25» (столбец «Факт KPI, %»).'
det['A2'].font=Font(name='Calibri',size=9,italic=True,color='6B7A8D'); det.merge_cells('A2:F2')
dv=DataValidation(type='decimal',operator='between',formula1=0,formula2=1.5,allow_blank=True)
dv.error='Выполнение 0–150%'; dv.errorTitle='Диапазон'; det.add_data_validation(dv)
for col,w in {'A':4,'B':50,'C':10,'D':16,'E':16,'F':4}.items(): det.column_dimensions[col].width=w

subtotal_ref={}
r=4
for (idx,name,dep,role,typ,oklad,note,kpis) in EMP:
    tag='ПЕРЕДОВАЯ ЛИНИЯ' if typ=='front' else 'ПОДДЕРЖКА'
    det.cell(r,2,'%s — %s · %s  [%s]'%(name,role,dep,tag))
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
    det.cell(r,2,'Итоговое выполнение показателей').font=BOLD
    wsum=det.cell(r,3,'=SUM(C%d:C%d)'%(first,last)); wsum.number_format='0%'; wsum.font=BOLD; wsum.alignment=CTR
    tcell=det.cell(r,5,'=SUM(E%d:E%d)'%(first,last)); tcell.number_format='0%'
    tcell.font=Font(name='Calibri',bold=True,size=10,color=GREEN); tcell.alignment=CTR
    for col in range(1,7): det.cell(r,col).fill=fill(LIGHT); det.cell(r,col).border=border
    subtotal_ref[idx]=r
    r+=2
det.freeze_panes='A4'

# =========== ЛИСТ 1: Матрица 75-25 (структура пользователя) ===========
mx=wb.create_sheet('Матрица 75-25',0); mx.sheet_view.showGridLines=False
mx['A1']='ТОО «WayStar Group» · Матрица структуры оплаты труда 75/25 и показателей'
mx['A1'].font=Font(name='Calibri',bold=True,size=14,color=NAVY)
mx['A2']='Приложение 1 к приказу о введении оплаты «75% постоянная + 25% переменная» (с 01.01.2026). '\
         '«Факт KPI, %» подтягивается с листа «Показатели (детально)». Премия = Лимит × мин(100%; факт) — по приказу не выше лимита 25%.'
mx['A2'].font=Font(name='Calibri',size=9,italic=True,color='6B7A8D'); mx.merge_cells('A2:L2')
heads=['Подразделение','Должность','ФИО','Табельный №/ID','Целевой ежемесячный доход (100%), ₸',
       'Оклад (75%), ₸','Лимит премии (25%), ₸','Дата начала действия','Примечание',
       'Факт KPI, %','Начисленная премия, ₸','Итого (оклад+премия), ₸']
hr=4
for j,txt in enumerate(heads,1):
    c=mx.cell(hr,j,txt); c.font=HDR; c.fill=fill(NAVY); c.border=border; c.alignment=CTR
widths=[14,26,22,10,15,14,14,11,26,9,15,16]
for j,w in enumerate(widths,1): mx.column_dimensions[get_column_letter(j)].width=w

row=hr+1; first_data=row
for (idx,name,dep,role,typ,oklad,note,kpis) in EMP:
    mx.cell(row,1,dep).font=REG; mx.cell(row,1).alignment=LEFT
    mx.cell(row,2,role).font=REG; mx.cell(row,2).alignment=LEFT
    mx.cell(row,3,name).font=REG; mx.cell(row,3).alignment=LEFT
    mx.cell(row,4,idx).font=REG; mx.cell(row,4).alignment=CTR
    ce=mx.cell(row,5,oklad); ce.number_format='#,##0'; ce.font=REG; ce.alignment=RIGHT
    ok=mx.cell(row,6,'=E%d*%s'%(row,FIX)); ok.number_format='#,##0'; ok.font=REG; ok.alignment=RIGHT
    lim=mx.cell(row,7,'=E%d*%s'%(row,VAR)); lim.number_format='#,##0'; lim.font=REG; lim.alignment=RIGHT
    mx.cell(row,8,DATE).font=REG; mx.cell(row,8).alignment=CTR
    mx.cell(row,9,note).font=Font(name='Calibri',size=8.5,color='6B7A8D'); mx.cell(row,9).alignment=LEFT
    fk=mx.cell(row,10,"='Показатели (детально)'!E%d"%subtotal_ref[idx]); fk.number_format='0%'
    fk.font=Font(name='Calibri',size=10,color=GREEN,bold=True); fk.alignment=CTR
    pr=mx.cell(row,11,'=G%d*MIN(1,J%d)'%(row,row)); pr.number_format='#,##0'; pr.font=BOLD; pr.alignment=RIGHT
    it=mx.cell(row,12,'=F%d+K%d'%(row,row)); it.number_format='#,##0'; it.font=BOLD; it.alignment=RIGHT
    zb=LIGHT if (row-first_data)%2==0 else WHITE
    for j in range(1,13):
        mx.cell(row,j).border=border
        if j!=10: mx.cell(row,j).fill=fill(zb)
    row+=1
last_data=row-1
mx.cell(row,3,'ИТОГО').font=Font(bold=True,color=WHITE,size=10)
for col,f in [(5,'E'),(6,'F'),(7,'G'),(11,'K'),(12,'L')]:
    c=mx.cell(row,col,'=SUM(%s%d:%s%d)'%(f,first_data,f,last_data)); c.number_format='#,##0'
for j in range(1,13):
    mx.cell(row,j).fill=fill(STEEL); mx.cell(row,j).border=border
    mx.cell(row,j).font=Font(bold=True,color=WHITE,size=10)
    mx.cell(row,j).alignment=RIGHT if j in (5,6,7,11,12) else CTR
trow=row
row+=2
mx.cell(row,3,'В год (×12), ₸:').font=BOLD
mx.cell(row,5,'=E%d*12'%trow).number_format='#,##0'; mx.cell(row,5).font=Font(bold=True,color=NAVY); mx.cell(row,5).alignment=RIGHT
mx.cell(row,7,'Лимит премий/год (под результатом):').font=Font(bold=True,size=9); mx.merge_cells(start_row=row,start_column=7,end_row=row,end_column=6+1)
mx.cell(row,11,'=K%d*12'%trow).number_format='#,##0'; mx.cell(row,11).font=Font(bold=True,color=COPPER); mx.cell(row,11).alignment=RIGHT
mx.cell(row,12,'=L%d*12'%trow).number_format='#,##0'; mx.cell(row,12).font=Font(bold=True,color=GREEN); mx.cell(row,12).alignment=RIGHT
row+=2
for note in [
 'Целевой ежемесячный доход (100%) = действующий оклад начисленный; из него 75% — постоянная часть, 25% — лимит премии.',
 'Факт KPI, % — итоговое выполнение показателей с листа «Показатели (детально)» (вводится там, жёлтые ячейки).',
 'Начисленная премия = Лимит × мин(100%; Факт KPI). По приказу (п.4.3) премия не превышает лимит переменной части (25%).',
 'Переменная может быть снижена/не выплачена при нарушениях дисциплины, сроков/качества (приказ п.4.4) — корректируется вручную.',
 'Табельный № — временный (порядковый); заменить на реальные ID при внедрении.']:
    mx.cell(row,1,'• '+note).font=Font(size=9,italic=True,color='6B7A8D')
    mx.merge_cells(start_row=row,start_column=1,end_row=row,end_column=12); row+=1
mx.freeze_panes='A5'

# =========== ЛИСТ: Справочник показателей (транспортная логистика) ===========
sp=wb.create_sheet('Справочник показателей')
sp.sheet_view.showGridLines=False
sp['A1']='Справочник показателей эффективности (транспортно-логистическая компания)'
sp['A1'].font=Font(name='Calibri',bold=True,size=13,color=NAVY)
sp['A2']='Библиотека показателей по функциям — можно адаптировать под должности. Тип: индивидуальный / командный.'
sp['A2'].font=Font(name='Calibri',size=9,italic=True,color='6B7A8D'); sp.merge_cells('A2:F2')
sph=['Функция / роль','Показатель','Метрика','Рекоменд. вес, %','Источник факта','Тип']
for j,t in enumerate(sph,1):
    c=sp.cell(4,j,t); c.font=HDR; c.fill=fill(NAVY); c.border=border; c.alignment=CTR
LIB=[
 ('Руководство','Чистая прибыль компании к плану','% к плану','40–50','Управл. отчётность','команд.'),
 ('Руководство','Выручка к плану','% к плану','15–20','Отчёт о продажах','команд.'),
 ('Руководство','Развитие: новые направления/клиенты','% этапов','15–20','План развития','индив.'),
 ('Финансы','Запас денежных средств / кассовые разрывы','мес.','20','Движение средств','команд.'),
 ('Финансы','Период оборота дебиторской задолженности','дни','15–20','Реестр расчётов','команд.'),
 ('Финансы','Точность бюджета (отклонение факт/план)','%','15','План/факт','индив.'),
 ('Операции','Валовая прибыль по операциям','млн ₸','40','Себестоимость','команд.'),
 ('Операции','Соблюдение сроков (в срок и в полном объёме)','% рейсов','20','Учёт перевозок','команд.'),
 ('Операции','Себестоимость на тонну/километр','₸','20','Себестоимость','индив.'),
 ('Операции','Безопасность: происшествия / охрана труда','кол-во','10','Служба безопасности','команд.'),
 ('Логистика','Валовая маржа по своим перевозкам','₸ / % ≥K×стоим.','45','Себестоимость','индив.'),
 ('Логистика','Объём организованных перевозок к норме','рейсы / тонны','15','Учёт перевозок','индив.'),
 ('Логистика','Соблюдение сроков доставки','% в срок','15','Учёт перевозок','индив.'),
 ('Логистика','Загрузка транспорта / порожний пробег','%','10–15','Учёт перевозок','индив.'),
 ('Логистика','Сохранность груза и уровень претензий','% / кол-во','10','Отдел рекламаций','индив.'),
 ('Расчёты/закрытие','Валовая прибыль по закрытым перевозкам','₸','45','Учёт/отчётность','индив.'),
 ('Расчёты/закрытие','Собираемость оплаты в срок','% в срок','30','Реестр расчётов','индив.'),
 ('Расчёты/закрытие','Срок закрытия документов','дни','15','Документооборот','индив.'),
 ('Персонал','Удержание ключевого персонала / текучесть','%','25','Кадровый учёт','команд.'),
 ('Персонал','Срок закрытия вакансий','дни','20','Отчёт по подбору','индив.'),
 ('Персонал','Точность и своевременность расчёта зарплаты','%','20','Бухгалтерия','индив.'),
 ('Юридический','Взысканная дебиторская задолженность / споры','₸ / %','25','Претензии','индив.'),
 ('Юридический','Согласование договоров в срок','дни / %','20','Реестр договоров','индив.'),
 ('Все роли','Чистая прибыль компании (общий показатель)','% к плану','25','Управл. отчётность','команд.'),
]
sw=[16,42,20,14,20,10]
for j,w in enumerate(sw,1): sp.column_dimensions[get_column_letter(j)].width=w
rr=5
for i,rowv in enumerate(LIB):
    for j,v in enumerate(rowv,1):
        c=sp.cell(rr,j,v); c.font=REG; c.border=border
        c.alignment=CTR if j in (3,4,6) else LEFT
        c.fill=fill(LIGHT if i%2==0 else WHITE)
    rr+=1
sp.freeze_panes='A5'

wb.save('motivation/WayStar_Матрица_и_показатели.xlsx')
print('saved motivation/WayStar_Матрица_и_показатели.xlsx')
