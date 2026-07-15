# -*- coding: utf-8 -*-
"""Classify each transaction into P&L / Cashflow categories (v2)."""
import json, re, collections
rows=json.load(open('/home/user/board-presentation/mgmt-report/ledger.json'))
BAL={'БЦК':(75346.24,110840346.22),'Halyk':(222587.02,2526900.00),'RBK':(0.0,44256.05)}
WAY='240340014287'

TRANSPORT_KW=('железнодорожн','ж/д','ж-д','жд тариф','тариф','вагон','дислокац','подвижн',
    'перевозк','экспедир','оперирование','организации перевоз','логистич','рефконтейн','контейнер','порож')
def has(p,kws): return any(k in p for k in kws)

def classify(x):
    knp=(x['knp'] or '').strip()
    cp=x['cp'].lower(); p=x['purpose'].lower()
    inflow = x['credit']>0
    own = ('waystar' in cp) or (x.get('bin')==WAY) or (x.get('cp_bin')==WAY)
    state = any(s in cp for s in ('государственная корпорация','кгд','угд','дгд','комитет государственных доходов','налоговый комитет'))

    # 1. STATE / TAX (counterparty is a budget/social body)
    if state or knp in ('911','912','010','011','012','121','122','089','101','102'):
        return ('TAX','Налоги и социальные платежи')
    # 2. INTERNAL own-account transfers
    if knp in ('343','342') or ('своего текущ' in p) or ('на клс' in p) or ('перевод в клс' in p) \
       or ('перевод на клс' in p) or (own and 'перевод на другой сч' in p):
        return ('INTERNAL','Переводы между своими счетами')
    # 3. TREASURY: deposits
    if knp=='312' or 'размещение вклада' in p or 'размещение депозита' in p:
        return ('TREASURY','Размещение депозита')
    if knp in ('321','322') or 'снятие с' in p or 'выплата вклада' in p or 'частичная выплата вклада' in p:
        return ('TREASURY','Снятие/возврат депозита')
    # 4. TREASURY: FX
    if knp=='213' or 'покупка иностранной валюты' in p:
        return ('TREASURY','Покупка валюты (FX)')
    # 5. Deposit interest income
    if knp=='316' or ('вознаграждени' in p and ('вклад' in p or 'депозит' in p)):
        return ('FIN_INCOME','% по депозиту')
    # 6. FINANCING / loans
    if knp in ('411','421','424','419','413') or 'предоставление кредита' in p or 'выдача займа' in p \
       or 'выдачу краткосрочных займов' in p or 'погашение краткосрочных займов' in p \
       or 'возврат займа' in p or 'возврат фин' in p or 'финансовой помощи' in p \
       or 'погашение основного долга' in p or 'возврат кредита' in p or 'погашение кредита' in p \
       or 'погашение суммы займа' in p or 'погашение займа' in p:
        if inflow: return ('FINANCING','Займы полученные / возврат выданных')
        return ('FINANCING','Погашение / выдача займов')
    # 7. Interest & penalties on loans (finance cost)
    if knp in ('423','119') or 'погашение вознаграждени' in p or 'погашение начисленных пени' in p \
       or ('вознаграждени' in p and 'договор' in p):
        return ('FIN_COST','Проценты и пени по кредитам')
    # 8. SALARY
    if knp=='332' or 'заработ' in p or 'зарплат' in p or 'для зп' in cp or 'трудового отпуска' in p \
       or (('аванс' in p) and ('июн' in p or 'июл' in p or 'ма' in p)):
        return ('OPEX','Оплата труда')
    # 9. TRAVEL
    if knp in ('871','872') or 'суточные' in p or 'командиров' in p or 'по поездк' in p:
        return ('OPEX','Командировочные / поездки')
    # 10. BANK
    if knp in ('841','842') or 'интернет-банкинг' in p or 'за ведение счета' in p \
       or (('комисси' in p) and 'банк' in cp):
        return ('OPEX','Банковские услуги')
    # 11. INSURANCE
    if knp=='833' or 'страхован' in p:
        return ('OPEX','Страхование')
    # 12. RENT (office)
    if 'property management' in cp or 'aldar' in cp or ('аренд' in p and 'подвижн' not in p and 'вагон' not in p and 'состав' not in p):
        return ('OPEX','Аренда офиса')
    # 13. CONTRA refund to client
    if knp=='880' or 'возврат денежных средств' in p:
        return ('CONTRA','Возврат средств клиенту')
    # 13b. Wagon repair & maintenance (direct fleet cost)
    if not inflow and ('ремонт' in p or 'колесных пар' in p or 'вчд' in p or 'деповск' in p or 'тор №' in p):
        return ('COGS','Ремонт и ТО вагонов')
    # 14. TRANSPORT core (revenue vs COGS)
    if knp in ('813','819','811','814','810','812','815','855') or has(p,TRANSPORT_KW):
        if inflow: return ('REVENUE','Выручка: транспортно-логистические услуги')
        return ('COGS','Ж/д тариф, вагоны, услуги перевозчиков')
    # 15. GOODS / office
    if knp=='710' or 'ежедневник' in p or 'канцеляр' in p:
        return ('OPEX','ТМЦ / офисные расходы')
    # 16. residual
    if inflow:
        if 'оплата' in p or 'услуг' in p or 'счет' in p or 'счёт' in p:
            return ('REVENUE','Выручка: транспортно-логистические услуги')
        return ('OTHER_IN','Прочие поступления')
    return ('OPEX','Прочие услуги и расходы')

for x in rows:
    ft,cat=classify(x)
    x['flow']=ft; x['cat']=cat
    x['amount']= x['credit'] if x['credit']>0 else -x['debit']
json.dump(rows, open('/home/user/board-presentation/mgmt-report/ledger_classified.json','w'), ensure_ascii=False, indent=1)

agg=collections.defaultdict(lambda:[0,0.0,0.0])
for x in rows:
    k=(x['flow'],x['cat']); agg[k][0]+=1; agg[k][1]+=x['credit']; agg[k][2]+=x['debit']
order=['REVENUE','CONTRA','COGS','OPEX','TAX','FIN_INCOME','FIN_COST','FINANCING','TREASURY','INTERNAL','OTHER_IN']
print(f"{'FLOW':<11}{'Категория':<46}{'n':>4}{'Поступления':>17}{'Выплаты':>17}")
for fl in order:
    for (f,c),(n,i,o) in sorted(agg.items(), key=lambda kv:-(kv[1][1]+kv[1][2])):
        if f!=fl: continue
        print(f"{f:<11}{c:<46}{n:>4}{i:>17,.0f}{o:>17,.0f}")
print("-"*95)
print(f"{'ИТОГО':<61}{sum(x['credit'] for x in rows):>17,.0f}{sum(x['debit'] for x in rows):>17,.0f}")
print(f"NET = {sum(x['amount'] for x in rows):,.0f}  (свод остатков: 297,933 -> 113,411,502 = +113,113,569)")
