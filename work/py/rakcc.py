# Parse RAKBANK credit card statements (text from pdftotext -layout) into rows
import re,sys,json,glob,os
R=re.compile(r'^\s*(\d\d/\d\d/\d{4})\s{2,}(.+?)\s{2,}([A-Z]{3})\s+([\d,]+\.\d\d)( Cr)?\s+(\S+)\s+([\d,]+\.\d\d)( Cr)?\s*$')
rows=[];seen=set()
for f in sorted(glob.glob(sys.argv[1]+'/*.txt')):
    tag=os.path.basename(f)[:7]
    for ln in open(f):
        m=R.match(ln)
        if not m: continue
        d,desc,cur,oa,ocr,fx,a,cr=m.groups()
        amt=float(a.replace(',',''))*(-1 if cr else 1)
        dd,mm,yy=d.split('/')
        k=(d,desc.strip(),amt)
        rows.append(dict(f=tag,date=f'{yy}-{mm}-{dd}',desc=re.sub(r'\s+',' ',desc.strip()),cur=cur,orig=float(oa.replace(',','')),amt=amt))
json.dump(rows,open(sys.argv[2],'w'),indent=0)
print(len(rows))
