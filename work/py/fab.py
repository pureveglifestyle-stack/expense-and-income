# Parse FAB credit card statements (pdftotext -layout text) into rows
import re,sys,json,glob,os
R=re.compile(r'^\s*(\d\d-\d\d-\d{4})\s+(\d\d-\d\d-\d{4})\s+(.+?)\s+AED\s+([\d,]+\.\d\d)\s')
rows=[];seen=set()
for f in sorted(glob.glob(sys.argv[1]+'/*.txt')):
    for ln in open(f):
        m=R.match(ln)
        if not m: continue
        td,pd,desc,a=m.groups(); desc=re.sub(r'\s+',' ',desc.strip())
        amt=float(a.replace(',',''))
        if re.search(r'PAYMENT RECEIVED|REVERSAL',desc): amt=-amt
        d,mo,y=td.split('-'); k=(td,desc,amt)
        rows.append(dict(date=f'{y}-{mo}-{d}',desc=desc,amt=amt))
json.dump(rows,open(sys.argv[2],'w'),indent=0); print(len(rows))
