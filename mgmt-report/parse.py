# -*- coding: utf-8 -*-
"""Parse 3 bank statements of TOO WayStar Group into a unified ledger."""
import xlrd, openpyxl, warnings, re, json
warnings.filterwarnings('ignore')
UP='/root/.claude/uploads/8c625b59-905c-5bbf-a364-1efacfc0757a/'
F1=UP+'5ea86348-__________________15.06.202615.07.2026.xls'   # БЦК
F2=UP+'91e90539-____________________15.06.202615.07.2026.xls' # Halyk
F3=UP+'fdba00ea-__________________15.06.202615.07.2026.xlsx'  # RBK

def num(s):
    if s is None: return 0.0
    if isinstance(s,(int,float)): return float(s)
    s=str(s).replace('\xa0','').replace(' ','').replace(',','.').strip()
    if s=='' : return 0.0
    try: return float(s)
    except: return 0.0

def d10(s):
    s=str(s)
    m=re.search(r'(\d{2}\.\d{2}\.\d{4})', s)
    return m.group(1) if m else s

rows=[]

# ---- FILE 1: БЦК ----
wb=xlrd.open_workbook(F1); sh=wb.sheet_by_index(0)
# header row 8, data from row 9
for r in range(9, sh.nrows):
    v=[sh.cell_value(r,c) for c in range(sh.ncols)]
    if not str(v[0]).strip(): continue
    debit=num(v[7]); credit=num(v[8])
    if debit==0 and credit==0: continue
    rows.append(dict(bank='БЦК', date=d10(v[1]), doc=str(v[0]),
        cp=str(v[5]).strip(), bin=str(v[4] if credit>0 else v[6]).strip(),
        cp_bin=str(v[4]).strip(),  # sender bin
        debit=debit, credit=credit, knp=str(v[9]).strip(),
        purpose=str(v[11]).strip()))

# ---- FILE 2: Halyk ----
wb=xlrd.open_workbook(F2); sh=wb.sheet_by_index(0)
# header row 10, data from 11
for r in range(11, sh.nrows):
    v=[sh.cell_value(r,c) for c in range(sh.ncols)]
    if not str(v[0]).strip(): continue
    if re.match(r'(Исходящий|Входящий|Итого)', str(v[0]).strip()): continue
    debit=num(v[7]); credit=num(v[8])
    if debit==0 and credit==0: continue
    rows.append(dict(bank='Halyk', date=d10(v[0]), doc=str(v[1]),
        cp=str(v[4]).strip(), bin=str(v[5]).strip(), cp_bin=str(v[5]).strip(),
        debit=debit, credit=credit, knp=str(v[11]).strip(),
        purpose=str(v[9]).strip()))

# ---- FILE 3: RBK ----
wb=openpyxl.load_workbook(F3, data_only=True); sh=wb.active
def parse_party(cell):
    """Extract name and BIN from multiline party cell."""
    t=str(cell or '')
    name=t.split('\n')[0].strip()
    m=re.search(r'БИН:\s*(\d+)', t)
    return name, (m.group(1) if m else '')
for r in range(20, sh.max_row+1):
    v=[sh.cell(r,c).value for c in range(1, sh.max_column+1)]
    if v[0] is None or str(v[0]).strip()=='': continue
    if not re.match(r'\d', str(v[0])): continue
    debit=num(v[5]); credit=num(v[6])
    if debit==0 and credit==0: continue
    send_name, send_bin = parse_party(v[3])
    recv_name, recv_bin = parse_party(v[4])
    # counterparty = the one that is not WayStar
    if credit>0:   # inflow: sender is counterparty
        cp, cpbin = send_name, send_bin
    else:          # outflow: receiver is counterparty
        cp, cpbin = recv_name, recv_bin
    rows.append(dict(bank='RBK', date=d10(v[1]), doc=str(v[2]),
        cp=cp, bin=cpbin, cp_bin=cpbin,
        debit=debit, credit=credit, knp='',
        purpose=str(v[7] or '').strip()))

json.dump(rows, open('/home/user/board-presentation/mgmt-report/ledger.json','w'), ensure_ascii=False, indent=1)
tot_in=sum(x['credit'] for x in rows); tot_out=sum(x['debit'] for x in rows)
print("Total records:", len(rows))
for b in ['БЦК','Halyk','RBK']:
    rr=[x for x in rows if x['bank']==b]
    print(f"  {b}: {len(rr)} tx | in(credit)={sum(x['credit'] for x in rr):,.0f} | out(debit)={sum(x['debit'] for x in rr):,.0f}")
print(f"TOTAL inflow={tot_in:,.0f}  outflow={tot_out:,.0f}  net={tot_in-tot_out:,.0f}")
