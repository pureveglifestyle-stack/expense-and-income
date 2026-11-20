import re,glob,json
amt=r'[\d,]+\.\d{2}'
rows=[]
for f in sorted(glob.glob('*.txt'), key=lambda x:x):
    txt=open(f,encoding='utf-8').read()
    blocks=re.split(r'\n\s*\n',txt)
    for b in blocks:
        m=re.search(r'^\s{0,12}(\d{2}/\d{2}/\d{4})\s',b,re.M)
        bal=re.findall(r'('+amt+r')\s+(Cr|Dr)\b',b)
        if not m or not bal: continue
        if 'BROUGHT FORWARD' in b or 'CARRIED FORWARD' in b:
            rows.append(dict(f=f,date=m.group(1),desc='BF',bal=float(bal[-1][0].replace(',',''))*(1 if bal[-1][1]=='Cr' else -1)));continue
        lines=[l.strip() for l in b.split('\n') if l.strip()]
        desc=' '.join(lines)
        desc=re.sub(r'\d{2}/\d{2}/\d{4}','',desc,1)
        desc=re.sub(amt+r'(\s+(Cr|Dr))?','',desc)
        rows.append(dict(f=f,date=m.group(1),desc=re.sub(r'\s+',' ',desc).strip(),bal=float(bal[-1][0].replace(',',''))*(1 if bal[-1][1]=='Cr' else -1)))
# signs
prev=None;out=[]
for r in rows:
    if r['desc']=='BF':
        if prev is not None and abs(prev-r['bal'])>0.01: print('BF mismatch',r,prev)
        prev=r['bal'];continue
    r['amt']=round(r['bal']-prev,2);prev=r['bal'];out.append(r)
json.dump(out,open('enbd_rows.json','w'))
print(len(out),'final bal',prev)
for r in out[:5]+out[-5:]: print(r['date'],r['amt'],r['bal'],r['desc'][:80])
