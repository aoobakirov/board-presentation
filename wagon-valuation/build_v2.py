#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Build v2 valuation model (component-based cost approach) for 166 tank cars.
Keeps source data (Комплектация, Справочник) intact; drops heavy v1 Амортизация
sheet (kept in the original upload for reference) and adds Параметры / Оценка / Свод.
"""
import openpyxl, datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import column_index_from_string as ci, get_column_letter as gl

SRC = "source/Оценка_166_цистерн_v1_исходник.xlsx"
OUT = "Оценка_166_цистерн_v2.xlsx"
EVAL_DATE = datetime.datetime(2026, 8, 5)

# ---------- read source ----------
wb  = openpyxl.load_workbook(SRC, data_only=False)
wbv = openpyxl.load_workbook(SRC, data_only=True)
kv  = wbv["Комплектация"]        # values
av  = wbv["Амортизация"]         # v1 cached values (for справочно)

def pdate(s):
    if s in (None, ""): return None
    if isinstance(s, datetime.datetime): return s
    p = str(s).split(".")
    if len(p) != 3: return None
    try:  return datetime.datetime(int(p[2]), int(p[1]), int(p[0]))
    except Exception: return None

# per-wagon helpers computed in Python (source-derived data, written as values)
N = 166
KROW0 = 5           # Комплектация first data row
helpers = []        # list of dicts
for i in range(N):
    kr = KROW0 + i
    def g(col): return kv.cell(kr, ci(col)).value
    F  = pdate(g("F"))
    G  = pdate(g("G"))
    I  = pdate(g("I"))
    L  = pdate(g("L"))
    last = None
    cand = [d for d in (G, I) if d is not None]
    if cand: last = max(cand)
    helpers.append({
        "end_year": (F.year if F else None),
        "last_rep": last,
        "next_rep": L,
        "v1z": av.cell(16 + i, ci("Z")).value,
    })

# ---------- styling helpers ----------
ARIAL = "Arial"
def F_(sz=10, b=False, color="000000"): return Font(name=ARIAL, size=sz, bold=b, color=color)
BLUE   = F_(color="0000FF")          # hardcoded input
GREEN  = F_(color="008000")          # link to other sheet
BLACK  = F_()                        # formula
HEADER = F_(b=True, color="FFFFFF")
TITLE  = F_(14, b=True)
YELLOW = PatternFill("solid", fgColor="FFFF00")
HFILL  = PatternFill("solid", fgColor="305496")
SUBFILL= PatternFill("solid", fgColor="D9E1F2")
thin   = Side(style="thin", color="BFBFBF")
BORD   = Border(left=thin, right=thin, top=thin, bottom=thin)
PCT    = '0.0"%"'
RUB    = '#,##0'
CTR    = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

# ---------- drop heavy v1 sheet; fix Справочник summary refs ----------
if "Амортизация" in wb.sheetnames:
    del wb["Амортизация"]
sp = wb["Справочник"]
# unmerge anything in rows 10-15 first (title cells are merged)
for mr in list(sp.merged_cells.ranges):
    if mr.min_row >= 10 and mr.max_row <= 15:
        sp.unmerge_cells(str(mr))
# old summary rows 10-15 referenced Амортизация -> clear to avoid #REF
for r in range(10, 16):
    for c in range(1, 5):
        cell = sp.cell(r, c)
        if not isinstance(cell, openpyxl.cell.cell.MergedCell):
            cell.value = None

# =====================================================================
# Параметры (v2 assumptions)  — B column holds the live levers
# =====================================================================
if "Параметры" in wb.sheetnames: del wb["Параметры"]
pm = wb.create_sheet("Параметры", 1)
pm.sheet_view.showGridLines = False
P = [
    # (row, label, value, unit, note, style)
    (1, "Параметры методики v2 (компонентный затратный подход)", None, None, None, "title"),
    (2, "Дата оценки", EVAL_DATE, "", "Фиксированная дата расчёта", "blue"),
    (3, "Цена новой цистерны", "=Справочник!B8", "₽/вагон", "Ссылка на Справочник (100% нового вагона)", "green"),
    (4, "— Доли компонентов в стоимости нового вагона (Σ = 100%) —", None, None, None, "sub"),
    (5, "Доля: котёл + рама", 0.55, "", "ДОПУЩЕНИЕ — уточнить по калькуляции завода", "blue"),
    (6, "Доля: тележки (литьё: боковые рамы + надрессорные балки)", 0.18, "", "ДОПУЩЕНИЕ — уточнить", "blue"),
    (7, "Доля: колёсные пары (4 шт)", 0.20, "", "ДОПУЩЕНИЕ — уточнить", "blue"),
    (8, "Доля: прочее (автосцепка, тормоз, арматура)", 0.07, "", "ДОПУЩЕНИЕ — уточнить", "blue"),
    (9, "Контроль суммы долей (должно быть 1,00)", "=SUM(B5:B8)", "", "Если ≠ 1,00 — исправьте доли", "chk"),
    (10, "— Ресурсы и пределы —", None, None, None, "sub"),
    (11, "Нормативный срок службы (fallback), лет", 32, "лет", "Используется, если нет даты СС в данных", "blue"),
    (12, "Срок службы литья тележек, лет", 30, "лет", "ДОПУЩЕНИЕ — нормативный срок рам/балок", "blue"),
    (13, "Толщина обода новой КП", 70, "мм", "ДОПУЩЕНИЕ — полный диапазон износа обода", "blue"),
    (14, "Предельная толщина обода", "=Справочник!B6", "мм", "Ссылка на Справочник (эксплуатационный предел)", "green"),
    (15, "Съём металла при обточке", "=Справочник!B4", "мм", "Ссылка на Справочник", "green"),
    (16, "Мин. обод после обточки (порог годности)", "=Справочник!B3", "мм", "Ссылка на Справочник", "green"),
    (17, "— Стоимости ремонтов, ₽ —", None, None, None, "sub"),
    (18, "Стоимость предстоящего планового ремонта (ДР)", 300000, "₽", "ДОПУЩЕНИЕ — заменить фактической сметой ДР", "blue"),
    (19, "Стоимость замены одной колёсной пары", 220000, "₽/шт", "ДОПУЩЕНИЕ — цена новой/годной КП", "blue"),
]
pm.column_dimensions["A"].width = 52
pm.column_dimensions["B"].width = 16
pm.column_dimensions["C"].width = 9
pm.column_dimensions["D"].width = 50
for row in P:
    r, label, val, unit, note, style = row
    a = pm.cell(r, 1, label)
    if style == "title":
        a.font = TITLE; pm.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4); continue
    if style == "sub":
        a.font = F_(b=True); a.fill = SUBFILL
        pm.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        for c in range(1,5): pm.cell(r,c).fill = SUBFILL
        continue
    a.font = F_(); a.alignment = LEFT
    b = pm.cell(r, 2, val)
    if style == "blue":
        b.font = BLUE; b.fill = YELLOW
    elif style == "green":
        b.font = GREEN
    elif style == "chk":
        b.font = F_(b=True)
    if isinstance(val, (int,)) and unit == "₽" or unit == "₽/шт":
        b.number_format = RUB
    if r in (5,6,7,8,9):
        b.number_format = "0.00"
    if r == 2:
        b.number_format = "dd.mm.yyyy"
    pm.cell(r, 3, unit).font = F_(9)
    d = pm.cell(r, 4, note); d.font = F_(9, color="808080"); d.alignment = LEFT

# convenient absolute refs
PB = "Параметры!$B$"
def pmref(r): return f"{PB}{r}"

# =====================================================================
# Оценка  (v2 per-wagon)
# =====================================================================
if "Оценка" in wb.sheetnames: del wb["Оценка"]
ws = wb.create_sheet("Оценка", 2)
ws.sheet_view.showGridLines = False
ws.freeze_panes = "A3"

headers = [
    ("A","№ вагона"), ("B","Модель"), ("C","Год постройки"),
    ("D","Год оконч. срока службы"), ("E","Полный срок службы, лет"),
    ("F","Возраст вагона, лет"), ("G","Износ котла+рамы, %"), ("H","Остаточная котла+рамы, %"),
    ("I","Средний год литья тележек"), ("J","Возраст литья, лет"),
    ("K","Износ литья, %"), ("L","Остаточная литья, %"),
    ("M","Мин. обод по КП, мм"), ("N","Остаточная колёсных пар, %"),
    ("O","Остаточная прочего, %"),
    ("P","Компонентная стоимость, ₽"),
    ("Q","Последний ремонт"), ("R","Следующий плановый ремонт"),
    ("S","Использовано цикла ДР, %"),
    ("T","Остаточный пробег, км"), ("U","Норма пробега, км"), ("V","Использовано цикла пробега, %"),
    ("W","Позиция обслуживания, %"), ("X","Обязательство ремонта (ДР), ₽"),
    ("Y","КП не проходят обточку, шт"), ("Z","Обязательство замены КП, ₽"),
    ("AA","ИТОГОВАЯ стоимость v2, ₽"),
    ("AB","Цена продавца, ₽ (ввод)"), ("AC","Отклонение рынок−модель, ₽"),
    ("AD","Оценка v1 (справочно), ₽"),
]
# title
ws.merge_cells("A1:AD1")
t = ws.cell(1,1,"Оценка остаточной стоимости 166 цистерн — методика v2 (компонентный затратный подход). Синие/жёлтые ячейки-допущения — на листе «Параметры».")
t.font = F_(11, b=True); t.alignment = LEFT
ws.row_dimensions[1].height = 28
# header row 2
for col, name in headers:
    c = ws.cell(2, ci(col), name)
    c.font = HEADER; c.fill = HFILL; c.alignment = CTR; c.border = BORD
ws.row_dimensions[2].height = 42

def K(col, kr): return f"Комплектация!{col}{kr}"

for i in range(N):
    r  = 3 + i          # Оценка data row
    kr = KROW0 + i      # Комплектация row
    h  = helpers[i]
    def S(col, val):
        ws.cell(r, ci(col)).value = val
    # identity
    S("A", f"={K('A',kr)}")
    S("B", f"={K('C',kr)}")
    S("C", f"={K('E',kr)}")
    # body (котёл+рама) — per-wagon service life from data, fallback to norm
    ws.cell(r, ci("D")).value = h["end_year"]                       # data helper
    S("E", f'=IF(D{r}="",{pmref(11)},MAX(1,D{r}-C{r}))')
    S("F", f"=MAX(0,YEAR({pmref(2)})-C{r})")
    S("G", f"=MIN(100,F{r}/E{r}*100)")
    S("H", f"=100-G{r}")
    # bogie castings (2 надрес. балки + 4 боковые рамы)
    S("I", f"=AVERAGE({K('BJ',kr)},{K('BP',kr)},{K('BV',kr)},{K('CB',kr)},{K('CH',kr)},{K('CN',kr)})")
    S("J", f"=MAX(0,YEAR({pmref(2)})-I{r})")
    S("K", f"=MIN(100,J{r}/{pmref(12)}*100)")
    S("L", f"=100-K{r}")
    # wheelsets — per-KP residual over FULL rim range (обод_new -> предел), averaged
    def kp_resid(c1, c2):
        mn = f"MIN({K(c1,kr)},{K(c2,kr)})"
        wear = f"MAX(0,MIN(100,({pmref(13)}-{mn})/({pmref(13)}-{pmref(14)})*100))"
        return f"(100-{wear})"
    S("M", f"=MIN({K('AA',kr)},{K('AB',kr)},{K('AK',kr)},{K('AL',kr)},{K('AU',kr)},{K('AV',kr)},{K('BE',kr)},{K('BF',kr)})")
    S("N", f"=AVERAGE({kp_resid('AA','AB')},{kp_resid('AK','AL')},{kp_resid('AU','AV')},{kp_resid('BE','BF')})")
    # other
    S("O", f"=AVERAGE(H{r},L{r})")
    # component value
    S("P", f"={pmref(3)}*({pmref(5)}*H{r}+{pmref(6)}*L{r}+{pmref(7)}*N{r}+{pmref(8)}*O{r})/100")
    # maintenance cycle (dates written as real Excel dates -> light formulas)
    ws.cell(r, ci("Q")).value = h["last_rep"]
    ws.cell(r, ci("R")).value = h["next_rep"]
    S("S", f'=IF(OR(Q{r}="",R{r}="",R{r}=Q{r}),"",MAX(0,MIN(100,({pmref(2)}-Q{r})/(R{r}-Q{r})*100)))')
    # пробег может содержать текст «не ремонтируется по пробегу» / пустоту -> N/A
    S("T", f'=IF(ISNUMBER({K("P",kr)}),{K("P",kr)},"")')
    S("U", f'=IF(ISNUMBER({K("N",kr)}),{K("N",kr)},"")')
    S("V", f'=IF(OR(U{r}="",T{r}="",U{r}=0),"",MAX(0,MIN(100,(U{r}-T{r})/U{r}*100)))')
    S("W", f'=MAX(IF(S{r}="",0,S{r}),IF(V{r}="",0,V{r}))')
    S("X", f"={pmref(18)}*W{r}/100")
    # wheelsets failing turning (min rim of pair - съём < мин после обточки)
    def fail(c1, c2):
        return f"IF(MIN({K(c1,kr)},{K(c2,kr)})-{pmref(15)}<{pmref(16)},1,0)"
    S("Y", f"={fail('AA','AB')}+{fail('AK','AL')}+{fail('AU','AV')}+{fail('BE','BF')}")
    S("Z", f"=Y{r}*{pmref(19)}")
    # final
    S("AA", f"=MAX(0,P{r}-X{r}-Z{r})")
    S("AB", None)                                   # market input (yellow)
    S("AC", f'=IF(AB{r}="","",AB{r}-AA{r})')
    ws.cell(r, ci("AD")).value = h["v1z"]           # v1 static reference

# formats & fonts for data area
pctcols  = ["G","H","K","L","N","O","S","V","W"]
rubcols  = ["P","X","Z","AA","AB","AC","AD"]
intcols  = ["C","D","E","F","J","M","T","U","Y"]
grncols  = ["A","B","C"]                            # links to Комплектация
for i in range(N):
    r = 3 + i
    for col,_ in headers:
        c = ws.cell(r, ci(col)); c.border = BORD; c.font = BLACK
        c.alignment = Alignment(horizontal="center", vertical="center")
    for col in pctcols: ws.cell(r, ci(col)).number_format = PCT
    for col in rubcols: ws.cell(r, ci(col)).number_format = RUB
    for col in intcols: ws.cell(r, ci(col)).number_format = "0"
    ws.cell(r, ci("I")).number_format = "0"
    for col in grncols: ws.cell(r, ci(col)).font = GREEN
    ws.cell(r, ci("AD")).font = F_(color="808080")          # v1 grey
    ws.cell(r, ci("AA")).font = F_(b=True)                   # headline bold
    ws.cell(r, ci("AB")).fill = YELLOW; ws.cell(r, ci("AB")).font = BLUE
    ws.cell(r, ci("Q")).number_format = "dd.mm.yyyy"
    ws.cell(r, ci("R")).number_format = "dd.mm.yyyy"
# widths
wide = {"A":12,"B":13,"C":9,"D":11,"E":11,"F":10,"G":11,"H":12,"I":11,"J":10,"K":10,"L":11,
        "M":10,"N":13,"O":12,"P":15,"Q":13,"R":14,"S":11,"T":13,"U":12,"V":13,"W":12,
        "X":15,"Y":11,"Z":15,"AA":17,"AB":15,"AC":16,"AD":16}
for col,w in wide.items(): ws.column_dimensions[col].width = w

DATA_FIRST, DATA_LAST = 3, 3 + N - 1

# =====================================================================
# Свод
# =====================================================================
if "Свод" in wb.sheetnames: del wb["Свод"]
sv = wb.create_sheet("Свод", 3)
sv.sheet_view.showGridLines = False
sv.column_dimensions["A"].width = 46
sv.column_dimensions["B"].width = 18
sv.column_dimensions["C"].width = 60
rng = f"Оценка!$AA${DATA_FIRST}:$AA${DATA_LAST}"
v1rng = f"Оценка!$AD${DATA_FIRST}:$AD${DATA_LAST}"
rows = [
    ("Свод оценки — методика v2", None, None, "title"),
    ("Компонентный затратный подход: стоимость = Σ(доля компонента × остаточная %) − обязательства по ремонтам.", None, None, "note"),
    ("", None, None, "blank"),
    ("Показатель", "Значение", "Комментарий", "hdr"),
    ("Количество вагонов", f"=COUNT({rng})", "Оценённых по модели v2", "d"),
    ("Средняя стоимость v2, ₽", f"=AVERAGE({rng})", "Итоговая с учётом обязательств", "d"),
    ("Минимальная стоимость v2, ₽", f"=MIN({rng})", "", "d"),
    ("Максимальная стоимость v2, ₽", f"=MAX({rng})", "", "d"),
    ("Суммарная стоимость парка v2, ₽", f"=SUM({rng})", "Итог по 166 вагонам", "d"),
    ("", None, None, "blank"),
    ("Средняя оценка v1 (справочно), ₽", f"=AVERAGE({v1rng})", "Из исходной модели (возраст+обод, 50/50)", "d"),
    ("Разница средних v2 − v1, ₽", f"=AVERAGE({rng})-AVERAGE({v1rng})", "Влияние новой методики", "d"),
    ("Разница средних v2 − v1, %", f"=(AVERAGE({rng})-AVERAGE({v1rng}))/AVERAGE({v1rng})*100", "", "dp"),
    ("", None, None, "blank"),
    ("Что изменено в методике v2 против v1:", None, None, "hdr2"),
    ("1. Компонентная декомпозиция", "", "Котёл+рама / тележки(литьё) / КП / прочее вместо 50-на-50", "b"),
    ("2. Индивидуальный срок службы", "", "Износ котла — по дате СС из данных, не общий 32 г.", "b"),
    ("3. Литьё тележек как компонент", "", "Использованы годы боковых рам и надрессорных балок", "b"),
    ("4. Обод по полному диапазону", "", "Износ обода от новой (70 мм) до предела, а не в полосе 50→27", "b"),
    ("5. Пробег как ресурс", "", "Остаточный пробег включён в позицию обслуживания", "b"),
    ("6. ДР как обязательство в ₽", "", "Стоимость предстоящего ремонта вычитается в рублях", "b"),
    ("7. Замена негодных КП в ₽", "", "Негодные к обточке КП — вычет стоимости замены", "b"),
    ("8. Рыночная сверка", "", "Колонка «Цена продавца» для сравнения (затратный↔рыночный)", "b"),
    ("", None, None, "blank"),
    ("Ограничения — требуют данных/уточнения:", None, None, "hdr2"),
    ("• Толщина гребня", "", "В источнике = 0 (не собрана). Добавить в износ КП после замера.", "w"),
    ("• Доли компонентов и цены ремонтов", "", "Заданы как допущения на листе «Параметры» — заменить фактикой.", "w"),
    ("• Состояние котла (коррозия, толщина стенки)", "", "Нет в данных; износ тела аппроксимирован сроком службы.", "w"),
    ("• Цены продавца / рынок", "", "Колонка AB пуста — заполнить для рыночной сверки.", "w"),
    ("• Функциональный / внешний износ", "", "Не учтён (тип цистерны, ставка аренды, регуляторные лимиты).", "w"),
]
r = 1
for label, val, note, style in rows:
    if style == "blank": r += 1; continue
    a = sv.cell(r,1,label)
    if style == "title":
        a.font = TITLE; sv.merge_cells(start_row=r,start_column=1,end_row=r,end_column=3); r+=1; continue
    if style == "note":
        a.font = F_(10, color="808080"); sv.merge_cells(start_row=r,start_column=1,end_row=r,end_column=3); r+=1; continue
    if style in ("hdr","hdr2"):
        a.font = F_(b=True, color="FFFFFF" if style=="hdr" else "000000")
        fill = HFILL if style=="hdr" else SUBFILL
        for c in range(1,4):
            sv.cell(r,c).fill = fill
            if style=="hdr": sv.cell(r,c).font = HEADER
        if style=="hdr":
            sv.cell(r,2,val); sv.cell(r,3,note)
        r+=1; continue
    a.font = F_(b=(style=="b")); a.alignment = LEFT
    if val not in (None,""):
        b = sv.cell(r,2,val); b.font = F_(b=True)
        b.number_format = RUB if style=="d" else ('0.0"%"' if style=="dp" else RUB)
    c3 = sv.cell(r,3,note); c3.font = F_(9,color="808080"); c3.alignment = LEFT
    r += 1

wb.active = wb.sheetnames.index("Свод")
# force full recalculation when the file is opened (formulas carry no cached values)
wb.calculation.calcMode = "auto"
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print("saved", OUT, "sheets:", wb.sheetnames)
