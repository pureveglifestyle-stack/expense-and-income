import json,re,collections
r=json.load(open('enbd/enbd_rows.json'))
RAK='AE9204000000335'
def cat(x):
    d=re.sub(r'\s+',' ',x['desc']); a=x['amt']
    dn=d.replace(' ','')
    if a>0:
        if 'ENVISION' in d and 'Travel' in d: return 'IN Envision travel'
        if 'ENVISION' in d: return 'IN Envision salary'
        if 'FROM-4413' in d: return 'IN ENBD card cash advance'
        if 'SDM REF' in d: return 'IN cash deposit'
        if 'IPP' in d and 'RAJESH GAJENDRAN' in d: return 'IN own transfer (RAK)'
        if 'IPI TT' in d: return 'IN own transfer (RAK)'
        if 'RAJESWARAN' in d: return 'IN Rajeswaran Subramanian'
        if 'TT REF' in d and 'RAJESH' in d: return 'IN own transfer (other)'
        if 'CARD NO' in d: return 'IN card refund'
        return 'IN other'
    if 'INR' in d and ('RMA REF' in d or 'EPHCOR' in d): return 'OUT India remittance'
    if 'CC NO.-4413' in dn.replace('CCNO.-','CC NO.-') or 'CC NO.-441358' in d: return 'OUT pay ENBD credit card'
    if 'CC NO.-413840' in d: return 'OUT pay EI card 3845'
    if 'RMA' in d and RAK in dn: return 'OUT own transfer to RAK'
    if 'INWARD CLEARING CHQ' in d: return 'OUT cheque'
    if 'Tabb' in d or 'TABBY' in d: return 'OUT Tabby'
    if 'CARD NO' in d: return 'OUT debit card purchase'
    if re.search(r'\d{2}:\d{2}:\d{2}',d): return 'OUT ATM cash'
    if 'SAR' in d and 'RMA TT' in d: return 'OUT own transfer to SNB / charge'
    if 'MOBILE BANKING TRANSFER' in d or 'RMA' in d: return 'OUT transfer to others'
    if re.search(r'CHARGE|FEE|VAT',d,re.I): return 'OUT bank fees'
    return 'OUT other'
s=collections.defaultdict(float);c=collections.Counter();ex=collections.defaultdict(list)
for x in r:
    k=cat(x); x['cat']=k; s[k]+=x['amt']; c[k]+=1
    if len(ex[k])<40: ex[k].append((x['date'],x['amt'],x['desc'][:110]))
for k in sorted(s): print(f'{k:40s} {c[k]:4d} {s[k]:12.2f}')
json.dump(r,open('enbd/enbd_cat.json','w'))
for k in ['OUT other','OUT transfer to others','OUT bank fees','IN other','OUT ATM cash','OUT cheque']:
    print('---',k)
    for e in ex[k]: print(e)
