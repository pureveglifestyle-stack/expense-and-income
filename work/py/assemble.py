import json,re,sys,collections,datetime
sys.path.insert(0,'py'); from cats import catm
ext=json.load(open('out_final2.json'))
keep=[]
for o in ext:
    if o['sub'] in ('ENBD card – statement total','Travel & hotels (ENBD card)','ENBD card – travel moved to work travel'): continue
    if o['sub']=='Emirates Islamic card': continue   # now detailed below
    if o['group']=='Credit card purchases (ENBD)': continue
    if o['group']=='Paid to other credit cards': continue   # FAB & RAK cards now itemised below
    keep.append(o)
ext=keep
W='Work travel (Envision allowance)'
def add(acct,date,desc,amt):
    c=catm(desc)
    if c=='Saudi': g,s=W,'Saudi purchases (credit card)'
    elif c=='Travel & hotels': g,s=W,'Travel & hotels'
    else: g,s=('FAB & RAK card purchases' if acct in ('FAB credit card','RAKBANK credit card') else 'Credit card purchases'),c
    ext.append(dict(date=date,acct=acct,group=g,sub=s,amt=round(amt,2),desc=desc))
# ENBD CC
cc=json.load(open('cc/cc_rows2.json'))
def kind(r):
    d=r['desc'].upper()
    if 'PAYMENT RECEIVED' in d or 'PAYMENT RECVD' in d: return 'payment'
    if 'CONVERSION TO INSTALLMENT' in d or 'FULL BALANCE IPP' in d or 'BALANCE CONVERSION' in d: return 'ipp'
    if re.search(r'CASH ADVANCE FEE|VAT ON|FINANCE CHARGE|PROCESSING FEE|FORECLOSURE|LATE PAYMENT|TEMP CREDIT|ANNUAL FEE',d): return 'fee'
    if re.search(r'CASH ADVANCE|BANKNET|^IBN BATUTA MALL BRANCH|^DIB |^MALL OF THE EMIRATES BRAN',d): return 'cash'
    return 'purchase'
n=0
for r in cc:
    if kind(r)!='purchase': continue
    d,m,y=r['tdate'].split('/'); add('ENBD credit card',f'{y}-{m}-{d}',r['desc'],r['amt']); n+=1
for mer,dt,a in json.load(open('cc/nov26.json')):
    if mer=='TOTAL': continue
    add('ENBD credit card',dt,mer.upper(),a); n+=1
# EI
ei=json.load(open('ei/ei_rows.json'))
for r in ei:
    if r['kind']=='purchase': add('Emirates Islamic card',r['date'],r['desc'],r['amt'])
    elif r['kind']=='credit': ext.append(dict(date=r['date'],acct='Emirates Islamic card',group='Credit card purchases',sub='Cashback & offers',amt=r['amt'],desc=r['desc']))
    elif r['kind']=='fee': ext.append(dict(date=r['date'],acct='Emirates Islamic card',group='Bank charges',sub='Emirates Islamic profit & fees',amt=r['amt'],desc=r['desc']))
# EI instalment profit inside EMIs paid so far (19,425 plan: 2 EMIs; 7,082 plan: 1 EMI; 20,513 plan is 0%)
ei_int=round(2*(923.98-19425/24)+(674.44-7082/12),2)
ext.append(dict(date='2026-10-31',acct='Emirates Islamic card',group='Bank charges',sub='Emirates Islamic profit & fees',amt=ei_int,desc='Profit inside instalment EMIs paid so far (24-month 19,425 plan ×2, 12-month 7,082 plan ×1)'))

# ---- FAB credit card (statements Oct & Nov 2026) ----
CC='Credit card purchases'; BC='Bank charges'
for r in json.load(open('fab/fab_rows.json')):
    d=r['desc'].upper(); a=r['amt']
    if re.search(r'PAYMENT RECEIVED|QUICK CASH BOOKING|BALANCE TRANSFER BOOKING|^QC |^BT |BILLED FINANCE|LATE|VAT REVERSAL',d): continue
    if re.search(r'FEE|CREDIT SHIELD|VAT',d):
        ext.append(dict(date=r['date'],acct='FAB credit card',group=BC,sub='FAB card fees & insurance',amt=a,desc=r['desc'])); continue
    add('FAB credit card',r['date'],r['desc'],a)
# interest inside the first FAB instalments (Quick Cash 5,000 / 6 m, Balance Transfer 2,000 / 3 m)
ext.append(dict(date='2026-11-01',acct='FAB credit card',group=BC,sub='FAB card fees & insurance',amt=round((852.83-5000/6)+(673.26-2000/3),2),desc='Interest inside first Quick Cash & Balance Transfer instalments'))
# ---- RAKBANK credit card (statements May-Nov 2026) ----
for r in json.load(open('rakcc/rakcc_rows.json')):
    d=r['desc'].upper(); a=r['amt']
    if re.search(r'PAYMENT RECD|CREDIT ADJUSTMENT|BALANCE TRANSFER|MONTHLY AMOUNT DUE|CREDIT CARD CHEQUE|CONVERT TRANSACTION',d): continue
    if re.search(r'CASHBACK|JOINING BONUS',d):
        ext.append(dict(date=r['date'],acct='RAKBANK credit card',group='FAB & RAK card purchases',sub='Cashback & offers',amt=a,desc=r['desc'])); continue
    if re.search(r'FEE|FINANCE CHARGE',d):
        ext.append(dict(date=r['date'],acct='RAKBANK credit card',group=BC,sub='RAKBANK card interest & fees',amt=a,desc=r['desc'])); continue
    if d.startswith('FAB,'):
        ext.append(dict(date=r['date'],acct='RAKBANK credit card',group='Cash withdrawn',sub='Cash from credit card',amt=a,desc=r['desc'])); continue
    add('RAKBANK credit card',r['date'],r['desc'],a)
# interest inside the first card-cheque instalment (9,000 over 12 m at 0.59%/month)
ext.append(dict(date='2026-11-12',acct='RAKBANK credit card',group=BC,sub='RAKBANK card interest & fees',amt=53.10,desc='Interest inside card-cheque instalment (9,000 over 12 months, 0.59% a month)'))
ext.sort(key=lambda o:o['date'])
json.dump(ext,open('out_final3.json','w'))
g=collections.defaultdict(float); sg=collections.defaultdict(float)
for o in ext: g[o['group']]+=o['amt']; sg[(o['group'],o['sub'])]+=o['amt']
for k in sorted(sg): print(k,round(sg[k],2))
print({k:round(v,2) for k,v in g.items()},round(sum(g.values()),2),len(ext),'cc rows',n,'ei_int',ei_int)
