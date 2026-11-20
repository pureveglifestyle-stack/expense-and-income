import re,sys
src=open('page2.py').read().rsplit('wb.save(sys.argv[1])',1)[0]
exec(src)
def clean(d,acct):
    d=re.sub(r'\s+',' ',d)
    d=re.sub(r'^CARD NO\.\d+X+\d+ ','',d); d=re.sub(r'^PURCHASE TRXN\.?-?(REV)? ','',d)
    d=re.sub(r'\s(\d{6}XX|\d{6,}).*$','',d)
    d=re.sub(r'(:|\s)(AE|SA|ARE|HK|IE|CY)\b.*$','',d)
    d=re.sub(r' D ?UBAI\b',' DUBAI',d); d=d.replace(' D AE','')
    return d.strip().title()[:45]
# ---------------- INDIA ----------------
wi=wb.create_sheet("3. India transfers"); ws=wi
ws.sheet_view.showGridLines=False; ws.sheet_properties.tabColor="2F6DB5"
for c,w in zip("ABCDEFGH",[2,13,12,16,15,15,16,2]): ws.column_dimensions[c].width=w
ws.merge_cells("B1:G1"); ws["B1"]="MONEY SENT TO INDIA"; ws["B1"].font=Font(bold=True,size=16,color=WHITE); ws["B1"].fill=fill("2F6DB5"); ws["B1"].alignment=Alignment(vertical="center",indent=1); ws.row_dimensions[1].height=34
ws.merge_cells("B2:G2"); ws["B2"]="All transfers to Sornapriya · Sep 2025 – Nov 2026 · AED is what left your account"; ws["B2"].font=Font(italic=True,size=9,color="5B6475")
ind=[o for o in ext if o['group']=='Sent to India']
first=6; last=first+len(ind)-1
ws.merge_cells("B3:C3"); ws["B3"]="TOTAL SENT (AED)"; ws["B3"].font=Font(bold=True,size=9,color="2F6DB5")
ws.merge_cells("B4:C4"); ws["B4"]=f"=SUM(F{first}:F{last})"; ws["B4"].font=Font(bold=True,size=18,color="2F6DB5"); ws["B4"].number_format="#,##0"
ws["D3"]="TRANSFERS"; ws["D3"].font=Font(bold=True,size=9,color="2F6DB5"); ws["D4"]=f"=COUNT(F{first}:F{last})"; ws["D4"].font=Font(bold=True,size=18,color="2F6DB5")
ws.merge_cells("E3:G3"); ws.merge_cells("E4:G4"); ws["E3"]="TOTAL INR (ENBD transfers)"; ws["E3"].font=Font(bold=True,size=9,color="2F6DB5"); ws["E4"]=f"=SUM(G{first}:G{last})"; ws["E4"].font=Font(bold=True,size=18,color="2F6DB5"); ws["E4"].number_format="#,##0"
for c in "BCDEFG":
    for r in (3,4): ws[f"{c}{r}"].fill=fill("E6EDF6")
hdr_cols="BCDEFG"
for c,l in zip(hdr_cols,["Date","From","To","Month","AED","INR"]):
    x=ws[f"{c}5"]; x.value=l; x.font=Font(bold=True,color=WHITE); x.fill=fill(NAVY); x.alignment=Alignment(horizontal="center")
prev=None
for i,o in enumerate(ind):
    r=first+i; y,m,d=map(int,o['date'].split('-'))
    inr=re.search(r'INR\s?([\d,]+)',o['desc']); inr=float(inr.group(1).replace(',','')) if inr else None
    vals=[date(y,m,d),'ENBD' if o['acct']=='ENBD' else 'RAKBANK','Sornapriya',date(y,m,1),o['amt'],inr]
    for c,v in zip(hdr_cols,vals):
        x=ws[f"{c}{r}"]; x.value=v; x.border=Border(bottom=thin)
    ws[f"B{r}"].number_format="dd-mmm-yy"; ws[f"E{r}"].number_format="mmm yyyy"; ws[f"F{r}"].number_format="#,##0.00"; ws[f"G{r}"].number_format="#,##0"
    if (y,m)!=prev and prev is not None:
        for c in hdr_cols: ws[f"{c}{r}"].border=Border(top=Side(style="medium",color="9DB7DB"),bottom=thin)
    prev=(y,m)
    if m%2: 
        for c in hdr_cols: ws[f"{c}{r}"].fill=fill("F4F7FC")
ws.conditional_formatting.add(f"F{first}:F{last}",DataBarRule(start_type="num",start_value=0,end_type="max",color="8DB4E2",showValue=True))
tr=last+1; ws[f"E{tr}"]="TOTAL"; ws[f"F{tr}"]=f"=SUM(F{first}:F{last})"; ws[f"G{tr}"]=f"=SUM(G{first}:G{last})"
ws[f"F{tr}"].number_format="#,##0.00"; ws[f"G{tr}"].number_format="#,##0"
for c in hdr_cols: ws[f"{c}{tr}"].font=Font(bold=True,color=WHITE); ws[f"{c}{tr}"].fill=fill("2F6DB5")
# monthly summary to the right? keep below for phone
ms=tr+3; ws[f"B{ms}"]="PER MONTH"; ws[f"B{ms}"].font=Font(bold=True,size=12,color=NAVY)
for c,l in zip("BCD",["Month","Transfers","AED"]):
    x=ws[f"{c}{ms+1}"]; x.value=l; x.font=Font(bold=True,color=WHITE); x.fill=fill(NAVY)
months=[]; y,m=2025,9
while (y,m)<=(2026,11):
    months.append((y,m)); m+=1
    if m==13: y,m=y+1,1
for i,(y,m) in enumerate(months):
    r=ms+2+i; ws[f"B{r}"]=date(y,m,1); ws[f"B{r}"].number_format="mmm yyyy"
    ws[f"C{r}"]=f'=COUNTIFS($E${first}:$E${last},B{r})'
    ws[f"D{r}"]=f'=SUMIFS($F${first}:$F${last},$E${first}:$E${last},B{r})'; ws[f"D{r}"].number_format='#,##0;-#,##0;"–"'
    for c in "BCD": ws[f"{c}{r}"].border=Border(bottom=thin)
mle=ms+2+len(months)-1
ws.conditional_formatting.add(f"D{ms+2}:D{mle}",DataBarRule(start_type="num",start_value=0,end_type="max",color="2F6DB5",showValue=True))
ch=BarChart(); ch.type="col"; ch.title="Sent to India per month (AED)"
ch.add_data(Reference(ws,min_col=4,min_row=ms+1,max_row=mle),titles_from_data=True)
ch.set_categories(Reference(ws,min_col=2,min_row=ms+2,max_row=mle))
ch.series[0].graphicalProperties.solidFill="2F6DB5"; ch.legend=None; ch.y_axis.majorGridlines=None; ch.y_axis.numFmt="#,##0"; ch.x_axis.number_format="mmm yy"
ch.height=8; ch.width=18; ws.add_chart(ch,f"B{mle+2}")
ws.freeze_panes="A6"
# ---------------- CATEGORY ----------------
wc=wb.create_sheet("4. Spending by category"); ws=wc
ws.sheet_view.showGridLines=False; ws.sheet_properties.tabColor="D9822B"
for c,w in zip("ABCDEFG",[2,13,22,44,14,12,2]): ws.column_dimensions[c].width=w
ws.merge_cells("B1:F1"); ws["B1"]="SPENDING BY CATEGORY"; ws["B1"].font=Font(bold=True,size=16,color=WHITE); ws["B1"].fill=fill("B8650F"); ws["B1"].alignment=Alignment(vertical="center",indent=1); ws.row_dimensions[1].height=34
ws.merge_cells("B2:F2"); ws["B2"]="Every payment grouped by category · tap a category name to jump to its list"; ws["B2"].font=Font(italic=True,size=9,color="5B6475")
CATS=[("Groceries",["Groceries"],0,"Need"),("Dining & food",["Dining & food"],0,"Cut"),("Taxi & transport",["Taxi & transport"],0,"Reduce"),
("Shopping, electronics & home",["Shopping & home"],0,"Cut"),("Education (DBA / courses)",["Education"],0,"Check"),("Telecom, DEWA & bills",["Telecom & bills"],0,"Need"),("Health & pharmacy",["Health & pharmacy"],0,"Need"),
("Personal care & laundry",["Personal care & laundry"],0,"Need"),("Kids & leisure",["Kids & leisure"],0,"Reduce"),("Subscriptions & apps",["Subscriptions & apps"],0,"Reduce"),
("Tabby / BNPL",["Tabby / BNPL"],0,"Stop"),("Courier & services",["Courier & services"],0,"-"),("Other purchases",["Other purchases"],0,"-"),("Cashback & offers (credit)",["Cashback & offers"],0,"-"),
("Rent",["Rent cheques (RAK)","Rent cheque (ENBD)"],0,"Need"),("Transfers to people",["Transfers to other people / accounts"],0,"Check"),
("Paid to FAB & RAK cards",["FAB card","RAK credit card"],0,"Debt"),("Cash withdrawn (UAE ATM)",["ATM – RAK","ATM – ENBD (UAE)"],0,"Reduce"),
("Bank charges & interest",["RAK account fees","ENBD account charges","ENBD transfer charges","SNB fees","ENBD card interest & fees","Emirates Islamic profit & fees"],0,"Stop"),
("WORK TRAVEL – Travel & hotels",["Travel & hotels"],0,"Envision"),("WORK TRAVEL – Saudi purchases",["Saudi purchases (ENBD card)","Saudi purchases (SNB)","Saudi purchases (credit card)"],0,"Envision"),("WORK TRAVEL – Saudi ATM cash",["ATM – ENBD (Saudi)","ATM – SNB (Saudi)"],0,"Envision")]
NC={"Need":"1D6B3A","Cut":RED,"Reduce":"B8650F","Stop":RED,"Envision":"20A39E","-":"5B6475","Check":"5B6475","Debt":RED}
# index
ws["B4"]="CATEGORY"; ws["D4"]="ITEMS"; ws["E4"]="TOTAL (AED)"; ws["F4"]="Need / Cut?"
ws.merge_cells("B4:C4")
for c in "BCDEF": ws[f"{c}4"].font=Font(bold=True,color=WHITE); ws[f"{c}4"].fill=fill(NAVY)
idx0=5; R0=idx0+len(CATS)+3
r=R0; anchors=[]
for k,(cn,subs,ccamt,flag) in enumerate(CATS):
    items=[o for o in ext if o['sub'] in subs]
    items.sort(key=lambda o:o['date'])
    head=r
    ws.merge_cells(f"B{r}:D{r}"); ws[f"B{r}"]=cn
    for c in "BCDEF": ws[f"{c}{r}"].fill=fill("FCEFDF" if not cn.startswith("WORK") else "DDF1E4")
    ws[f"B{r}"].font=Font(bold=True,size=12,color="7A3E00" if not cn.startswith("WORK") else "1D6B3A")
    ws[f"F{r}"]=flag; ws[f"F{r}"].font=Font(bold=True,color=NC[flag]); ws[f"F{r}"].alignment=Alignment(horizontal="right")
    r+=1
    for c,l in zip("BCDEF",["Date","Account","Paid to / details","Amount (AED)",""]):
        x=ws[f"{c}{r}"]; x.value=l; x.font=Font(bold=True,size=9,color="5B6475"); x.border=Border(bottom=thin)
    r+=1; s0=r
    if ccamt:
        ws[f"B{r}"]="Nov 25–Nov 26"; ws[f"C{r}"]="ENBD credit card"; ws[f"D{r}"]="All ENBD card purchases in this category (12 months, from statements)"; ws[f"E{r}"]=ccamt
        ws[f"D{r}"].font=Font(italic=True,size=9); ws[f"E{r}"].number_format="#,##0.00"
        for c in "BCDE": ws[f"{c}{r}"].border=Border(bottom=thin); ws[f"{c}{r}"].fill=fill("FFF8EC")
        r+=1
    for o in items:
        y,m,d=map(int,o['date'].split('-'))
        ws[f"B{r}"]=date(y,m,d); ws[f"B{r}"].number_format="dd-mmm-yy"
        ws[f"C{r}"]=o['acct']
        det=o['desc'] if o['acct'] in ('SNB (KSA)',) or o['sub'].endswith('fees') else clean(o['desc'],o['acct'])
        if o['sub'].startswith('Rent'): det='Rent cheque '+(re.search(r'0000\d+',o['desc']).group(0) if re.search(r'0000\d+',o['desc']) else '')
        if o['sub'] in ('FAB card','RAK credit card'): det='Bill payment – '+o['sub']
        if o['sub'].startswith('ATM'): det=o['sub']
        if o['sub']=='Transfers to other people / accounts' and o['acct']=='ENBD':
            ib=re.search(r'AE\d{2}\s?\d{3}',o['desc'].replace(' ','')); det='Transfer to account '+(re.sub(r'\s','',o['desc'])[re.sub(r'\s','',o['desc']).find('AE'):][:23] if 'AE' in o['desc'] else '')
        ws[f"D{r}"]=det; ws[f"E{r}"]=o['amt']; ws[f"E{r}"].number_format="#,##0.00"
        ws[f"D{r}"].font=Font(size=9)
        for c in "BCDE": ws[f"{c}{r}"].border=Border(bottom=Side(style="hair",color="D5DAE1"))
        r+=1
    s1=r-1
    ws[f"D{r}"]=f"Total {cn}"; ws[f"E{r}"]=f"=SUM(E{s0}:E{s1})" if s1>=s0 else 0
    ws[f"D{r}"].alignment=Alignment(horizontal="right")
    for c in "DE": ws[f"{c}{r}"].font=Font(bold=True); ws[f"{c}{r}"].border=Border(top=Side(style="medium",color="B8650F"))
    ws[f"E{r}"].number_format="#,##0.00"
    anchors.append((cn,head,r,s0,s1,flag))
    r+=3
# fill index
for k,(cn,head,tot,s0,s1,flag) in enumerate(anchors):
    ir=idx0+k
    ws.merge_cells(f"B{ir}:C{ir}"); ws[f"B{ir}"]=cn
    ws[f"B{ir}"].hyperlink=f"#'4. Spending by category'!B{head}"; ws[f"B{ir}"].font=Font(color="1F5FAD",underline="single",bold=cn.startswith("WORK"))
    ws[f"D{ir}"]=f"=COUNT(E{s0}:E{s1})" if s1>=s0 else 0
    ws[f"E{ir}"]=f"=E{tot}"; ws[f"E{ir}"].number_format="#,##0.00"
    ws[f"F{ir}"]=flag; ws[f"F{ir}"].font=Font(bold=True,color=NC[flag])
    for c in "BCDEF": ws[f"{c}{ir}"].border=Border(bottom=thin)
    if cn.startswith("WORK"):
        for c in "BCDEF": ws[f"{c}{ir}"].fill=fill("EEF8F1")
it=idx0+len(anchors)
ws[f"B{it}"]="TOTAL"; ws[f"E{it}"]=f"=SUM(E{idx0}:E{it-1})"; ws[f"E{it}"].number_format="#,##0.00"
for c in "BCDEF": ws[f"{c}{it}"].font=Font(bold=True,color=WHITE); ws[f"{c}{it}"].fill=fill("B8650F")
ws.conditional_formatting.add(f"E{idx0}:E{it-1}",DataBarRule(start_type="num",start_value=0,end_type="max",color="F2B36A",showValue=True))
ws.freeze_panes="A3"
for w_ in (wi,wc):
    w_.page_setup.orientation="portrait"; w_.page_setup.fitToWidth=1; w_.page_setup.fitToHeight=0; w_.sheet_properties.pageSetUpPr.fitToPage=True
print("india",len(ind),"cats",[(a[0],a[4]-a[3]+1) for a in anchors])
wb.save(sys.argv[1])
