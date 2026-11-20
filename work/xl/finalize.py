# Store calculated values in the workbook so phone viewers show totals.
# Usage: python3 finalize.py in.xlsx out.xlsx   (needs LibreOffice 'soffice')
import sys,os,re,shutil,subprocess,tempfile,zipfile,urllib.parse
src,dst=sys.argv[1],sys.argv[2]
tmp=tempfile.mkdtemp(); base='book.xlsx'; shutil.copy(src,os.path.join(tmp,base))
out=os.path.join(tmp,'out'); os.makedirs(out)
subprocess.run(['soffice','--headless','--calc','--convert-to','xlsx:Calc MS Excel 2007 XML','--outdir',out,os.path.join(tmp,base)],check=True,capture_output=True,timeout=300)
rt=os.path.join(out,base)
zin=zipfile.ZipFile(rt); zout=zipfile.ZipFile(dst,'w',zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    data=zin.read(item.filename)
    if re.match(r'xl/worksheets/sheet\d+\.xml$',item.filename):
        s=data.decode()
        def fix(m):
            h=re.sub(r'\s+r:id="[^"]*"','',m.group(0))
            return re.sub(r'location="([^"]*)"',lambda x:'location="'+urllib.parse.unquote(x.group(1))+'"',h)
        s=re.sub(r'<hyperlink [^>]*location="[^"]*"[^>]*/>',fix,s); data=s.encode()
    elif re.match(r'xl/worksheets/_rels/sheet\d+\.xml\.rels$',item.filename):
        s=data.decode()
        s=re.sub(r'<Relationship [^>]*relationships/hyperlink"[^>]*TargetMode="External"/>','',s); data=s.encode()
    zout.writestr(item,data)
zout.close(); shutil.rmtree(tmp); print('saved',dst)
