# -*- coding: utf-8 -*-
"""Build management report Excel workbook from classified ledger."""
import json, re, collections
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

rows=json.load(open('/home/user/board-presentation/mgmt-report/ledger_classified.json'))
BAL={'БЦК':(75346.24,110840346.22),'Halyk':(222587.02,2526900.00),'RBK':(0.0,44256.05)}
PERIOD="15.06.2026 – 15.07.2026"
ENTITY='ТОО «WayStar Group» (БИН 240340014287)'

def norm(n):
    n=re.sub(r'^(Товарищество с ограниченной ответственностью|Акционерное общество|ТОО|АО|Индивидуальный предприниматель|ИП|Частная компания|Производственный кооператив)\s*','',n,flags=re.I)
    return n.strip(' "«»')

def sum_cat(cat): return sum(x['debit']+x['credit'] for x in rows if x['cat']==cat)
def sum_flow_in(f): return sum(x['credit'] for x in rows if x['flow']==f)
def sum_flow_out(f): return sum(x['debit'] for x in rows if x['flow']==f)

REVENUE=sum_flow_in('REVENUE'); REFUND=sum_flow_out('CONTRA')
COGS=sum_flow_out('COGS'); OPEX=sum_flow_out('OPEX'); TAX=sum_flow_out('TAX')
FIN_IN=sum_flow_in('FIN_INCOME')+sum(x['credit'] for x in rows if x['flow']=='FIN_COST')
FIN_OUT=sum_flow_out('FIN_COST')
NET_REV=REVENUE-REFUND; GROSS=NET_REV-COGS; EBITDA=GROSS-OPEX
NET_PROFIT=EBITDA-TAX+FIN_IN-FIN_OUT

opex_cats=collections.OrderedDict()
for x in rows:
    if x['flow']=='OPEX': opex_cats[x['cat']]=opex_cats.get(x['cat'],0)+x['debit']
cogs_cats=collections.OrderedDict()
for x in rows:
    if x['flow']=='COGS': cogs_cats[x['cat']]=cogs_cats.get(x['cat'],0)+x['debit']

rev_by=collections.defaultdict(lambda:[0.0,0])
for x in rows:
    if x['flow']=='REVENUE': rev_by[norm(x['cp'])][0]+=x['credit']; rev_by[norm(x['cp'])][1]+=1
clients=sorted(rev_by.items(), key=lambda kv:-kv[1][0])
tot_rev=sum(v[0] for v in rev_by.values())
cost_by=collections.defaultdict(float)
for x in rows:
    if x['flow'] in ('COGS','OPEX'): cost_by[norm(x['cp'])]+=x['debit']
suppliers=sorted(cost_by.items(), key=lambda kv:-kv[1])

NAVY="13294B"; BLUE="2F6690"; ORANGE="C4772F"; LGREY="EEF2F6"; GREEN="1E7A3D"; RED="B23A3A"
thin=Side(style='thin', color='C9D3DD')
border=Border(left=thin,right=thin,top=thin,bottom=thin)
NUM='#,##0;[Red]-#,##0'
def cell(ws,r,c,v,bold=False,fill=None,fontcolor="1A1A1A",size=10,align='left',fmt=None,italic=False,wrap=False,bd=True):
    x=ws.cell(r,c,v)
    x.font=Font(name='Calibri',size=size,bold=bold,color=fontcolor,italic=italic)
    x.alignment=Alignment(horizontal=align,vertical='center',wrap_text=wrap)
    if fill: x.fill=PatternFill('solid',fgColor=fill)
    if fmt: x.number_format=fmt
    if bd: x.border=border
    return x
def banner(ws,r,ncol,text,fill=NAVY,fontcolor="FFFFFF",size=11):
    cell(ws,r,1,text,bold=True,fill=fill,fontcolor=fontcolor,size=size)
    for c in range(2,ncol+1): cell(ws,r,c,"",fill=fill)
def hdr(ws,r,c,t): cell(ws,r,c,t,bold=True,fill=NAVY,fontcolor="FFFFFF",size=10,align='center',wrap=True)
def note(ws,r,text,ncol,rows_span=2):
    cell(ws,r,1,text,italic=True,size=9,fontcolor="777777",wrap=True,bd=False)
    ws.merge_cells(start_row=r,start_column=1,end_row=r+rows_span,end_column=ncol)

wb=Workbook()

# ===== 1. РЕЗЮМЕ =====
ws=wb.active; ws.title="Резюме"; ws.sheet_view.showGridLines=False
cell(ws,1,1,"УПРАВЛЕНЧЕСКАЯ ОТЧЁТНОСТЬ",bold=True,size=16,fontcolor=NAVY,bd=False)
cell(ws,2,1,ENTITY,size=11,fontcolor=BLUE,bd=False)
cell(ws,3,1,f"Период: {PERIOD}  ·  выписки 3 банков (БЦК, Halyk, Bank RBK)  ·  валюта: KZT",size=10,italic=True,bd=False)
cell(ws,4,1,"Метод: кассовый (по фактическому движению денег на счетах)",size=9,italic=True,fontcolor="777777",bd=False)
r=6; banner(ws,r,4,"КЛЮЧЕВЫЕ ПОКАЗАТЕЛИ (₸)")
kpis=[("Выручка (поступления от клиентов)",REVENUE,None,False),
 ("Себестоимость (ж/д тариф, вагоны, ремонт)",-COGS,RED,False),
 ("ВАЛОВАЯ ПРИБЫЛЬ",GROSS,GREEN,True),
 ("  валовая маржа, %",GROSS/NET_REV,None,False,'%'),
 ("Операционные расходы (SG&A)",-OPEX,RED,False),
 ("EBITDA",EBITDA,GREEN,True),
 ("Налоги уплаченные",-TAX,RED,False),
 ("Финансовые доходы (% по депозитам)",FIN_IN,None,False),
 ("Финансовые расходы (% и пени по кредитам)",-FIN_OUT,RED,False),
 ("ЧИСТАЯ ПРИБЫЛЬ (кассовая)",NET_PROFIT,GREEN,True),
 ("  чистая маржа, %",NET_PROFIT/NET_REV,None,False,'%')]
r+=1
for it in kpis:
    label,val,col,big=it[0],it[1],it[2],it[3]; pct=len(it)>4
    cell(ws,r,1,label,bold=big,size=11 if big else 10,fill=LGREY if big else None)
    for c in (2,3): cell(ws,r,c,"",fill=LGREY if big else None)
    cell(ws,r,4,val,bold=big,align='right',fmt=('0.0%' if pct else NUM),fontcolor=(col or "1A1A1A"),size=11 if big else 10,fill=LGREY if big else None)
    r+=1
r+=1; banner(ws,r,4,"ДВИЖЕНИЕ ДЕНЕЖНЫХ СРЕДСТВ (свод, ₸)")
op_net=NET_REV-COGS-OPEX-TAX+FIN_IN-FIN_OUT
inv_net=sum_flow_in('TREASURY')+FIN_IN-sum_flow_out('TREASURY')
fin_net=sum_flow_in('FINANCING')-sum_flow_out('FINANCING')-FIN_OUT
intern_net=sum_flow_in('INTERNAL')-sum_flow_out('INTERNAL')
other_net=sum_flow_in('OTHER_IN')
open_bal=sum(v[0] for v in BAL.values()); close_bal=sum(v[1] for v in BAL.values())
# reconcile fin: put FIN_OUT into operating already; keep financing pure principal
op_net=NET_REV-COGS-OPEX-TAX+FIN_IN-FIN_OUT
inv_net=sum_flow_in('TREASURY')-sum_flow_out('TREASURY')
fin_net=sum_flow_in('FINANCING')-sum_flow_out('FINANCING')
cfl=[("Денежные средства на начало периода",open_bal,True),
 ("Операционный денежный поток",op_net,False),
 ("Инвестиц./казначейские (депозиты, валюта, %)",inv_net+FIN_IN- FIN_IN,False),
 ("Финансовые потоки (займы, кредиты)",fin_net,False),
 ("Переводы на прочие счета компании (КЛС и пр.)",intern_net,False),
 ("Прочее",other_net,False),
 ("Денежные средства на конец периода",close_bal,True)]
r+=1
for label,val,tot in cfl:
    cell(ws,r,1,label,bold=tot,fill=LGREY if tot else None,size=10)
    for c in (2,3): cell(ws,r,c,"",fill=LGREY if tot else None)
    cell(ws,r,4,val,bold=tot,align='right',fmt=NUM,fill=LGREY if tot else None,fontcolor=(GREEN if val>0 else RED))
    r+=1
ws.column_dimensions['A'].width=48
for c in 'BCD': ws.column_dimensions[c].width=16

# ===== 2. P&L =====
ws=wb.create_sheet("P&L (кассовый)"); ws.sheet_view.showGridLines=False
cell(ws,1,1,"ОТЧЁТ О ПРИБЫЛЯХ И УБЫТКАХ (кассовый метод)",bold=True,size=14,fontcolor=NAVY,bd=False)
cell(ws,2,1,f"{ENTITY}  ·  {PERIOD}  ·  ₸",size=10,italic=True,bd=False)
r=4
for c,t in [(1,"Показатель"),(2,"Сумма, ₸"),(3,"% к чист. выручке")]: hdr(ws,r,c,t)
def pl_row(label,val,bold=False,indent=False,color="1A1A1A"):
    global r; r+=1
    cell(ws,r,1,("      " if indent else "")+label,bold=bold,fill=LGREY if bold else None,size=10)
    cell(ws,r,2,val,bold=bold,align='right',fmt=NUM,fill=LGREY if bold else None,fontcolor=color)
    cell(ws,r,3,(val/NET_REV if NET_REV else 0),bold=bold,align='right',fmt='0.0%',fill=LGREY if bold else None)
pl_row("Выручка от реализации услуг",REVENUE,bold=True)
pl_row("минус: возврат средств клиентам",-REFUND,indent=True)
pl_row("Чистая выручка",NET_REV,bold=True)
for cat,val in cogs_cats.items(): pl_row("минус: "+cat,-val,indent=True)
pl_row("Себестоимость услуг, итого",-COGS,bold=True,color=RED)
pl_row("ВАЛОВАЯ ПРИБЫЛЬ",GROSS,bold=True,color=GREEN)
for cat,val in sorted(opex_cats.items(),key=lambda kv:-kv[1]): pl_row("минус: "+cat,-val,indent=True)
pl_row("Операционные расходы, итого",-OPEX,bold=True,color=RED)
pl_row("EBITDA",EBITDA,bold=True,color=GREEN)
pl_row("минус: налоги и социальные платежи",-TAX,indent=True)
pl_row("плюс: финансовые доходы (% по депозитам)",FIN_IN,indent=True)
pl_row("минус: финансовые расходы (% и пени)",-FIN_OUT,indent=True)
pl_row("ЧИСТАЯ ПРИБЫЛЬ",NET_PROFIT,bold=True,color=GREEN)
r+=2
note(ws,r,"Отчёт построен на кассовой основе по банковским выпискам: «выручка» = поступления от клиентов, «расходы» = фактические платежи за период. Не учитывает начисления, дебиторскую/кредиторскую задолженность и амортизацию. Депозиты, займы и переводы между своими счетами исключены из P&L и отражены в ДДС.",3,2)
ws.column_dimensions['A'].width=54; ws.column_dimensions['B'].width=18; ws.column_dimensions['C'].width=15

# ===== 3. P&L ПО КЛИЕНТАМ =====
ws=wb.create_sheet("P&L по клиентам"); ws.sheet_view.showGridLines=False
cell(ws,1,1,"P&L В РАЗБИВКЕ ПО КЛИЕНТАМ",bold=True,size=14,fontcolor=NAVY,bd=False)
cell(ws,2,1,f"{ENTITY}  ·  {PERIOD}  ·  ₸",size=10,italic=True,bd=False)
note(ws,3,"Выручка — фактическая по каждому клиенту. Себестоимость и косвенные расходы распределены пропорционально доле клиента в выручке (в выписках нет привязки платежей поставщикам к конкретному клиенту), поэтому маржа по клиенту — оценочная.",7,1)
r=5
for c,t in enumerate(["Клиент","Выручка","Доля","Себест. (распр.)","Валовая приб.","GM %","Чистая приб. (распр.)"],1): hdr(ws,r,c,t)
overhead=OPEX+TAX+FIN_OUT-FIN_IN+REFUND
for name,(rev,n) in clients:
    r+=1; share=rev/tot_rev if tot_rev else 0
    alloc_cogs=COGS*share; gm=rev-alloc_cogs; net=gm-overhead*share
    cell(ws,r,1,name[:45],size=10); cell(ws,r,2,rev,align='right',fmt=NUM)
    cell(ws,r,3,share,align='right',fmt='0.0%'); cell(ws,r,4,-alloc_cogs,align='right',fmt=NUM,fontcolor=RED)
    cell(ws,r,5,gm,align='right',fmt=NUM); cell(ws,r,6,(gm/rev if rev else 0),align='right',fmt='0.0%')
    cell(ws,r,7,net,align='right',fmt=NUM,fontcolor=(GREEN if net>0 else RED))
r+=1
cell(ws,r,1,"ИТОГО",bold=True,fill=LGREY); cell(ws,r,2,tot_rev,bold=True,align='right',fmt=NUM,fill=LGREY)
cell(ws,r,3,1.0,bold=True,align='right',fmt='0.0%',fill=LGREY); cell(ws,r,4,-COGS,bold=True,align='right',fmt=NUM,fill=LGREY,fontcolor=RED)
cell(ws,r,5,tot_rev-COGS,bold=True,align='right',fmt=NUM,fill=LGREY); cell(ws,r,6,((tot_rev-COGS)/tot_rev),bold=True,align='right',fmt='0.0%',fill=LGREY)
cell(ws,r,7,tot_rev-COGS-overhead,bold=True,align='right',fmt=NUM,fill=LGREY,fontcolor=GREEN)
ws.column_dimensions['A'].width=42
for c in 'BDEG': ws.column_dimensions[c].width=17
for c in 'CF': ws.column_dimensions[c].width=9
ws.freeze_panes="A6"

# ===== 4. ДДС =====
ws=wb.create_sheet("ДДС (Cash Flow)"); ws.sheet_view.showGridLines=False
cell(ws,1,1,"ОТЧЁТ О ДВИЖЕНИИ ДЕНЕЖНЫХ СРЕДСТВ (прямой метод)",bold=True,size=14,fontcolor=NAVY,bd=False)
cell(ws,2,1,f"{ENTITY}  ·  {PERIOD}  ·  ₸",size=10,italic=True,bd=False)
r=4
for c,t in [(1,"Статья"),(2,"Поступления"),(3,"Выплаты"),(4,"Итого")]: hdr(ws,r,c,t)
def cf_row(label,inflow,outflow,bold=False):
    global r; r+=1
    cell(ws,r,1,("" if bold else "      ")+label,bold=bold,fill=LGREY if bold else None,size=10)
    cell(ws,r,2,inflow or None,align='right',fmt=NUM,fill=LGREY if bold else None)
    cell(ws,r,3,(-outflow if outflow else None),align='right',fmt=NUM,fill=LGREY if bold else None,fontcolor=RED if outflow else "1A1A1A")
    net=(inflow or 0)-(outflow or 0)
    cell(ws,r,4,net,align='right',fmt=NUM,bold=bold,fill=LGREY if bold else None,fontcolor=(GREEN if net>0 else RED))
cf_row("Остаток денежных средств на начало периода",open_bal,0,bold=True)
r+=1; banner(ws,r,4,"ОПЕРАЦИОННАЯ ДЕЯТЕЛЬНОСТЬ",fill=BLUE)
cf_row("Поступления от клиентов",REVENUE,0)
cf_row("Возврат средств клиентам",0,REFUND)
cf_row("Оплата перевозчикам / ж/д тариф / вагоны (COGS)",0,COGS)
cf_row("Операционные расходы (ЗП, аренда, банк, связь)",0,OPEX)
cf_row("Налоги и социальные платежи",0,TAX)
cf_row("Проценты и пени по кредитам",0,FIN_OUT)
cf_row("Проценты по депозитам",FIN_IN,0)
cf_row("Итого по операционной деятельности",REVENUE+FIN_IN,REFUND+COGS+OPEX+TAX+FIN_OUT,bold=True)
r+=1; banner(ws,r,4,"ИНВЕСТИЦИОННАЯ / КАЗНАЧЕЙСКАЯ ДЕЯТЕЛЬНОСТЬ",fill=BLUE)
cf_row("Снятие / возврат депозитов",sum_cat('Снятие/возврат депозита'),0)
cf_row("Размещение депозитов",0,sum_cat('Размещение депозита'))
cf_row("Покупка иностранной валюты (FX)",0,sum_cat('Покупка валюты (FX)'))
cf_row("Итого по инвестиц./казнач. деятельности",sum_flow_in('TREASURY'),sum_flow_out('TREASURY'),bold=True)
r+=1; banner(ws,r,4,"ФИНАНСОВАЯ ДЕЯТЕЛЬНОСТЬ",fill=BLUE)
cf_row("Займы полученные / возврат выданных займов",sum_flow_in('FINANCING'),0)
cf_row("Погашение и выдача займов",0,sum_flow_out('FINANCING'))
cf_row("Итого по финансовой деятельности",sum_flow_in('FINANCING'),sum_flow_out('FINANCING'),bold=True)
r+=1; banner(ws,r,4,"ВНУТРЕННИЕ ПЕРЕВОДЫ И ПРОЧЕЕ",fill=BLUE)
cf_row("Переводы между собственными счетами (в т.ч. на КЛС)",sum_flow_in('INTERNAL'),sum_flow_out('INTERNAL'))
cf_row("Прочие поступления",other_net,0)
r+=1; banner(ws,r,4,"ЧИСТОЕ ИЗМЕНЕНИЕ ДЕНЕЖНЫХ СРЕДСТВ")
cell(ws,r,4,close_bal-open_bal,bold=True,align='right',fmt=NUM,fill=NAVY,fontcolor="FFFFFF")
cf_row("Остаток денежных средств на конец периода",close_bal,0,bold=True)
r+=2; cell(ws,r,1,"Сверка по банковским счетам (₸):",bold=True,size=10,bd=False)
r+=1
for c,t in [(1,"Банк / счёт"),(2,"Остаток на начало"),(3,"Остаток на конец")]: hdr(ws,r,c,t)
cell(ws,r,4,"",fill=NAVY)
names={'БЦК':'АО «Банк ЦентрКредит»','Halyk':'АО «Народный Банк»','RBK':'АО «Bank RBK»'}
for b,(o,cl) in BAL.items():
    r+=1; cell(ws,r,1,names[b],size=10); cell(ws,r,2,o,align='right',fmt=NUM); cell(ws,r,3,cl,align='right',fmt=NUM); cell(ws,r,4,"")
r+=1
cell(ws,r,1,"ИТОГО по 3 счетам",bold=True,fill=LGREY); cell(ws,r,2,open_bal,bold=True,align='right',fmt=NUM,fill=LGREY)
cell(ws,r,3,close_bal,bold=True,align='right',fmt=NUM,fill=LGREY); cell(ws,r,4,"",fill=LGREY)
ws.column_dimensions['A'].width=54
for c in 'BCD': ws.column_dimensions[c].width=18

# ===== 5. ВЫРУЧКА ПО КЛИЕНТАМ =====
ws=wb.create_sheet("Выручка по клиентам"); ws.sheet_view.showGridLines=False
cell(ws,1,1,"ВЫРУЧКА ПО КЛИЕНТАМ",bold=True,size=14,fontcolor=NAVY,bd=False)
cell(ws,2,1,f"{PERIOD}  ·  ₸",size=10,italic=True,bd=False)
r=4
for c,t in [(1,"Клиент"),(2,"Операций"),(3,"Сумма, ₸"),(4,"Доля")]: hdr(ws,r,c,t)
for name,(rev,n) in clients:
    r+=1; cell(ws,r,1,name[:50],size=10); cell(ws,r,2,n,align='center')
    cell(ws,r,3,rev,align='right',fmt=NUM); cell(ws,r,4,rev/tot_rev,align='right',fmt='0.0%')
r+=1
cell(ws,r,1,"ИТОГО",bold=True,fill=LGREY); cell(ws,r,2,sum(v[1] for v in rev_by.values()),bold=True,align='center',fill=LGREY)
cell(ws,r,3,tot_rev,bold=True,align='right',fmt=NUM,fill=LGREY); cell(ws,r,4,1.0,bold=True,align='right',fmt='0.0%',fill=LGREY)
ws.column_dimensions['A'].width=48; ws.column_dimensions['B'].width=12; ws.column_dimensions['C'].width=18; ws.column_dimensions['D'].width=10
ws.freeze_panes="A5"

# ===== 6. РАСХОДЫ =====
ws=wb.create_sheet("Расходы"); ws.sheet_view.showGridLines=False
cell(ws,1,1,"РАСХОДЫ ПО КАТЕГОРИЯМ И ПОСТАВЩИКАМ",bold=True,size=14,fontcolor=NAVY,bd=False)
cell(ws,2,1,f"{PERIOD}  ·  ₸",size=10,italic=True,bd=False)
r=4; banner(ws,r,2,"По категориям",fill=BLUE)
r+=1
for c,t in [(1,"Категория"),(2,"Сумма, ₸")]: hdr(ws,r,c,t)
allcost=collections.OrderedDict()
for cat,v in cogs_cats.items(): allcost["COGS · "+cat]=v
for cat,v in sorted(opex_cats.items(),key=lambda kv:-kv[1]): allcost["OPEX · "+cat]=v
allcost["Налоги и соц. платежи"]=TAX
allcost["Проценты и пени по кредитам"]=FIN_OUT
for cat,v in allcost.items():
    r+=1; cell(ws,r,1,cat,size=10); cell(ws,r,2,v,align='right',fmt=NUM)
r+=1
cell(ws,r,1,"ИТОГО (без займов/депозитов/переводов)",bold=True,fill=LGREY); cell(ws,r,2,sum(allcost.values()),bold=True,align='right',fmt=NUM,fill=LGREY)
r+=2; banner(ws,r,2,"Топ-20 поставщиков (COGS + OPEX)",fill=BLUE)
r+=1
for c,t in [(1,"Поставщик"),(2,"Сумма, ₸")]: hdr(ws,r,c,t)
for name,v in suppliers[:20]:
    r+=1; cell(ws,r,1,name[:50],size=10); cell(ws,r,2,v,align='right',fmt=NUM)
ws.column_dimensions['A'].width=54; ws.column_dimensions['B'].width=18

# ===== 7. РЕЕСТР =====
ws=wb.create_sheet("Реестр операций"); ws.sheet_view.showGridLines=False
cell(ws,1,1,"РЕЕСТР ОПЕРАЦИЙ (классифицированный)",bold=True,size=13,fontcolor=NAVY,bd=False)
r=3
for c,t in enumerate(["Банк","Дата","Контрагент","БИН","Поступление","Выплата","КНП","Тип","Категория","Назначение платежа"],1): hdr(ws,r,c,t)
data=sorted(rows,key=lambda x:(x['date'][6:10]+x['date'][3:5]+x['date'][0:2], x['bank']))
for x in data:
    r+=1
    cell(ws,r,1,x['bank'],size=9); cell(ws,r,2,x['date'],size=9,align='center')
    cell(ws,r,3,x['cp'][:40],size=9); cell(ws,r,4,x.get('bin',''),size=9)
    cell(ws,r,5,x['credit'] or None,size=9,align='right',fmt=NUM)
    cell(ws,r,6,x['debit'] or None,size=9,align='right',fmt=NUM)
    cell(ws,r,7,x['knp'],size=9,align='center'); cell(ws,r,8,x['flow'],size=9)
    cell(ws,r,9,x['cat'],size=9); cell(ws,r,10,x['purpose'][:80],size=9)
for i,w in enumerate([8,11,34,15,15,15,7,11,32,60],1): ws.column_dimensions[get_column_letter(i)].width=w
ws.freeze_panes="A4"; ws.auto_filter.ref=f"A3:J{r}"

out='/home/user/board-presentation/mgmt-report/WayStar_Управленческая_отчетность_15.06-15.07.2026.xlsx'
wb.save(out)
print("SAVED:",out)
print(f"Revenue={REVENUE:,.0f} COGS={COGS:,.0f} Gross={GROSS:,.0f} ({GROSS/NET_REV*100:.1f}%) EBITDA={EBITDA:,.0f} Net={NET_PROFIT:,.0f}")
print(f"Clients={len(clients)} Suppliers={len(suppliers)} Rows={len(rows)}")
