import re,glob,json,sys
files=sorted(glob.glob(sys.argv[1]+'/rakbankstament__TransactionDetails_*.pdf.txt'))
files=[f for f in files if '(1)' not in f]
amt=r'[\d,]+\.\d{2}'
rows=[]
for f in files:
    lines=open(f,encoding='utf-8').read().split('\n')
    # keep only transaction region lines
    blocks=[];cur=[]
    intx=False
    for l in lines:
        if 'Your Current Account Transactions' in l: intx=True; continue
        if 'The National Bank of Ras Al Khaimah' in l: intx=False
        if not intx: 
            if cur: blocks.append(cur);cur=[]
            continue
        if l.strip()=='' :
            if cur: blocks.append(cur);cur=[]
        else: cur.append(l)
    for b in blocks:
        txt=' '.join(b)
        m=re.search(r'(\d{2}-[A-Za-z]{3}-\d{4})',txt)
        bal=re.findall(r'('+amt+r')\s+(Cr|Dr)\.',txt)
        if not m or not bal: continue
        nums=re.findall(amt,txt)
        a=float(nums[-2].replace(',','')) if len(nums)>=2 else None
        b_=float(bal[-1][0].replace(',',''))*(1 if bal[-1][1]=='Cr' else -1)
        desc=re.sub(r'\s+',' ',txt)
        desc=desc.replace(m.group(1),'',1)
        desc=re.sub(r'Value Date:\s*\d{2}-\d{2}-\d{4}','',desc)
        desc=re.sub(amt+r'(\s+(Cr|Dr)\.)?','',desc).strip()
        rows.append(dict(file=f.split('_')[-1][:6],date=m.group(1),desc=desc,amt=a,bal=b_))
json.dump(rows,open(sys.argv[2],'w'),indent=0)
print(len(rows))
