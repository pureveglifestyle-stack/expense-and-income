import json,sys,datetime,collections
src=open('build.py').read().rsplit('import sys; wb.save',1)[0]
exec(src)
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.series import DataPoint
ext=json.load(open('../out_final3.json'))
ws2=wb.create_sheet("2. Money out"); ws=ws2
ws.sheet_view.showGridLines=False; ws.sheet_properties.tabColor="B42318"
widths=[2,24,16,16,15,15,15,14,14,14,14,15,2]
for i,w in enumerate(widths): ws.column_dimensions["ABCDEFGHIJKLM"[i]].width=w
RED="B42318"
GC={"Work travel (Envision allowance)":"1D6B3A","Sent to India":"2F6DB5","Housing":"7B4FC9","Credit card purchases":"D9822B","Debit card purchases":"E0B21E","FAB & RAK card purchases":"C2410C","Transfers to people":"20A39E","Cash withdrawn":"6B7280","Bank charges":"B42318"}
GP={"Work travel (Envision allowance)":"DDF1E4","Sent to India":"DCE8F7","Housing":"E9E1F7","Credit card purchases":"FCE9D2","Debit card purchases":"FBF3D0","FAB & RAK card purchases":"FDE3D3","Transfers to people":"D6F2F0","Cash withdrawn":"ECEEF1","Bank charges":"FBE0DD"}
ws.merge_cells("B1:L1"); ws["B1"]="MY DUBAI MONEY JOURNEY  ·  PAGE 2: MONEY OUT"
ws["B1"].font=Font(bold=True,size=16,color=WHITE); ws["B1"].fill=fill("7A1C14"); ws["B1"].alignment=Alignment(vertical="center",indent=1); ws.row_dimensions[1].height=36
ws.merge_cells("B2:L2"); ws["B2"]="30 Aug 2025 – 19 Nov 2026 · AED · RAK, ENBD & SNB accounts + ENBD, Emirates Islamic, FAB & RAKBANK card statements · money moved between your own accounts is NOT counted"
ws["B2"].font=Font(italic=True,size=9,color="5B6475")
# detail location
N=len(ext); DH=120; df=DH+1; dl=df+N-1
X=lambda c:f"${c}${df}:${c}${dl}"   # B date C account D group E sub F amount G desc
def SG(g): return f'SUMIFS({X("F")},{X("D")},"{g}")'
groups=["Sent to India","Housing","Credit card purchases","Debit card purchases","FAB & RAK card purchases","Transfers to people","Cash withdrawn","Bank charges","Work travel (Envision allowance)"]
gdesc={"Sent to India":"38 transfers to Sornapriya (ENBD + RAK)","Housing":"Rent cheques (RAK 7, ENBD 1)","Credit card purchases":"Every purchase on ENBD + Emirates Islamic cards (excl. travel)","Debit card purchases":"RAK & ENBD debit cards in UAE, incl. Tabby","FAB & RAK card purchases":"Every purchase on FAB & RAKBANK cards (statements unlocked)","Transfers to people":"Workers, friends, other accounts","Cash withdrawn":"ATM cash in UAE","Bank charges":"Card interest, card fees & insurance, account fees","Work travel (Envision allowance)":"Travel, hotels, Saudi purchases & Saudi ATM cash"}
# tiles
spent_f="=C10+C11+C12+C13+C14"
tiles=[("B","C","TOTAL MONEY OUT","=C18",RED,"FBE6E4","all money that left you"),
       ("D","E","SENT TO INDIA","=C9","2F6DB5","E6EDF6","=TEXT(C9/C18,\"0%\")&\" of money out\""),
       ("F","H","PERSONAL SPEND","=SUM(C10:C16)","B8650F","FCEFDF","rent, cards, purchases, cash, charges"),
       ("I","L","WORK TRAVEL","=C17","1D6B3A","E3F4E8","=TEXT('1. Income'!F5/C17,\"0%\")&\" covered by Envision allowance\"")]
for a,b,lab,fx,dark,pale,note in tiles:
    for r in (4,5,6): ws.merge_cells(f"{a}{r}:{b}{r}")
    for r in (4,5,6):
        for cc_ in range(ord(a),ord(b)+1): ws[f"{chr(cc_)}{r}"].fill=fill(pale)
    ws[f"{a}4"]=lab; ws[f"{a}4"].font=Font(bold=True,size=9,color=dark)
    ws[f"{a}5"]=fx; ws[f"{a}5"].font=Font(bold=True,size=20,color=dark); ws[f"{a}5"].number_format="#,##0"
    ws[f"{a}6"]=note; ws[f"{a}6"].font=Font(size=8,color="5B6475")
    for r in (4,5,6): ws[f"{a}{r}"].alignment=Alignment(horizontal="center",vertical="center")
ws.row_dimensions[5].height=34
def hdr(r,cols,labels,color=NAVY):
    for c,l in zip(cols,labels):
        x=ws[f"{c}{r}"]; x.value=l; x.font=Font(bold=True,color=WHITE,size=10); x.fill=fill(color); x.alignment=Alignment(horizontal="center",wrap_text=True)
ws["B7"]="WHERE THE MONEY WENT"; ws["B7"].font=Font(bold=True,size=12,color=NAVY)
hdr(8,"BCDE",["Bucket","Amount (AED)","Share","What it is"])
for i,gname in enumerate(groups):
    r=9+i
    ws[f"B{r}"]=gname; ws[f"B{r}"].font=Font(bold=True,color=GC[gname]); ws[f"B{r}"].fill=fill(GP[gname])
    ws[f"C{r}"]="="+SG(gname); ws[f"C{r}"].number_format="#,##0.00"
    ws[f"D{r}"]=f"=C{r}/$C$18"; ws[f"D{r}"].number_format="0.0%"
    ws[f"E{r}"]=gdesc[gname]; ws[f"E{r}"].font=Font(size=8,color="5B6475")
    for c in "BCDE": ws[f"{c}{r}"].border=Border(bottom=thin)
ws["B18"]="TOTAL MONEY OUT"; ws["C18"]="=SUM(C9:C17)"; ws["D18"]="=SUM(D9:D17)"
for c in "BCD": ws[f"{c}18"].font=Font(bold=True); ws[f"{c}18"].border=Border(top=Side(style="medium",color=NAVY))
ws["C18"].number_format="#,##0.00"; ws["D18"].number_format="0%"
ws.conditional_formatting.add("C9:C16",DataBarRule(start_type="num",start_value=0,end_type="max",color="E07A6B",showValue=True)) if False else ws.conditional_formatting.add("C9:C17",DataBarRule(start_type="num",start_value=0,end_type="max",color="E07A6B",showValue=True))
pie=PieChart(); pie.title="Money out by bucket"
pie.add_data(Reference(ws,min_col=3,min_row=8,max_row=17),titles_from_data=True)
pie.set_categories(Reference(ws,min_col=2,min_row=9,max_row=17))
pie.dataLabels=DataLabelList(); pie.dataLabels.showPercent=True; pie.dataLabels.showVal=False; pie.dataLabels.showCatName=False; pie.dataLabels.showSerName=False; pie.dataLabels.showLegendKey=False
for idx,gname in enumerate(groups):
    pt=DataPoint(idx=idx); pt.graphicalProperties.solidFill=GC[gname]; pie.series[0].dPt.append(pt)
pie.height=8.5; pie.width=13; ws.add_chart(pie,"G7")
# income vs out
ws["B20"]="INCOME vs MONEY OUT   |   WORK TRAVEL vs ENVISION ALLOWANCE"; ws["B20"].font=Font(bold=True,size=12,color=NAVY)
box=[("Total income (page 1)","='1. Income'!B5","1D6B3A"),("Total money out","=C18",RED),("Went out MORE than came in","=C21-C22","7A1C14")]
for i,(l,fx,col) in enumerate(box):
    r=21+i; ws[f"B{r}"]=l; ws[f"C{r}"]=fx; ws[f"C{r}"].number_format="#,##0.00;[Red]-#,##0.00"
    ws[f"B{r}"].font=Font(bold=True,color=col); ws[f"C{r}"].font=Font(bold=True,color=col,size=12)
    for c in "BC": ws[f"{c}{r}"].border=Border(bottom=thin)
abox=[("Envision travel allowance","='1. Income'!F5","1D6B3A"),("Work travel spent","=C17","B8650F"),("Extra paid from your own pocket","=F22-F21",RED)]
for i,(l,fx,col) in enumerate(abox):
    r=21+i; ws.merge_cells(f"D{r}:E{r}"); ws[f"D{r}"]=l; ws[f"F{r}"]=fx; ws[f"F{r}"].number_format="#,##0.00;[Red]-#,##0.00"; ws[f"D{r}"].font=Font(bold=True,color=col); ws[f"F{r}"].font=Font(bold=True,color=col,size=12)
    for c in "DEF": ws[f"{c}{r}"].border=Border(bottom=thin)
ws["B24"]="The gap was covered by credit cards (debt today ≈ 117,660 + 8,274 RAKBANK card-cheque loan), card cash advances, Rajeswaran's 6,763 and Envision-period cash deposits (28,450)."
ws["B24"].font=Font(italic=True,size=9,color="5B6475")
# category table
ws["B27"]="WHAT YOU BOUGHT (credit + debit cards)"; ws["B27"].font=Font(bold=True,size=12,color=NAVY)
hdr(28,"BCDEF",["Category","Credit cards","Debit cards","Total","Need / Cut?"])
cats=[("Groceries",["Groceries"],"Need"),("Shopping, electronics & home",["Shopping & home"],"Cut"),("Education (DBA / courses)",["Education"],"Check"),("Telecom, DEWA & bills",["Telecom & bills"],"Need"),("Dining & food",["Dining & food"],"Cut"),("Taxi & transport",["Taxi & transport"],"Reduce"),("Tabby / BNPL",["Tabby / BNPL"],"Stop"),("Health, pharmacy & nutrition",["Health & pharmacy"],"Need"),("Personal care & laundry",["Personal care & laundry"],"Need"),("Kids & leisure",["Kids & leisure"],"Reduce"),("Subscriptions & apps",["Subscriptions & apps"],"Reduce"),("Courier & services",["Courier & services"],"-"),("Other",["Other purchases"],"-"),("Cashback & offers",["Cashback & offers"],"-"),("Travel & hotels (work)",["Travel & hotels"],"Envision"),("Saudi purchases (work)",["Saudi purchases (SNB)","Saudi purchases (ENBD card)","Saudi purchases (credit card)"],"Envision")]
NC={"Need":"1D6B3A","Cut":RED,"Reduce":"B8650F","Stop":RED,"Envision":"20A39E","-":"5B6475","Check":"5B6475"}
CCA=["ENBD credit card","Emirates Islamic card","FAB credit card","RAKBANK credit card"]
for i,(cn,subs,flag) in enumerate(cats):
    r=29+i; ws[f"B{r}"]=cn
    ws[f"C{r}"]="="+"+".join(f'SUMIFS({X("F")},{X("E")},"{s_}",{X("C")},"{a_}")' for s_ in subs for a_ in CCA)
    ws[f"D{r}"]="="+"+".join(f'SUMIFS({X("F")},{X("E")},"{s_}")' for s_ in subs)+f"-C{r}"
    ws[f"E{r}"]=f"=C{r}+D{r}"; ws[f"F{r}"]=flag; ws[f"F{r}"].font=Font(bold=True,color=NC[flag]); ws[f"F{r}"].alignment=Alignment(horizontal="center")
    for c in "CDE": ws[f"{c}{r}"].number_format='#,##0;-#,##0;"–"'
    ws[f"E{r}"].font=Font(bold=True)
    for c in "BCDEF":
        ws[f"{c}{r}"].border=Border(bottom=thin)
        if i%2: ws[f"{c}{r}"].fill=fill("F7F9FC")
cl=29+len(cats)-1; ct=cl+1
ws[f"B{ct}"]="TOTAL BOUGHT"
for c in "CDE": ws[f"{c}{ct}"]=f"=SUM({c}29:{c}{cl})"; ws[f"{c}{ct}"].number_format="#,##0"
for c in "BCDE": ws[f"{c}{ct}"].font=Font(bold=True); ws[f"{c}{ct}"].border=Border(top=Side(style="medium",color=NAVY))
ws.conditional_formatting.add(f"E29:E{cl}",DataBarRule(start_type="num",start_value=0,end_type="max",color="F2B36A",showValue=True))
hb=BarChart(); hb.type="bar"; hb.title="Spending by category (AED)"; hb.style=10
hb.add_data(Reference(ws,min_col=5,min_row=28,max_row=cl),titles_from_data=True)
hb.set_categories(Reference(ws,min_col=2,min_row=29,max_row=cl))
hb.series[0].graphicalProperties.solidFill="D9822B"; hb.legend=None; hb.y_axis.majorGridlines=None
hb.x_axis.scaling.orientation="maxMin"; hb.height=9; hb.width=12
ws.add_chart(hb,"G27")
# monthly table
M0=ct+3
ws[f"B{M0}"]="MONTH BY MONTH"; ws[f"B{M0}"].font=Font(bold=True,size=12,color=NAVY)
mh=["Month","India","Housing","Card buys","Debit buys","Other cards","People","Cash","Charges","Work travel","TOTAL"]
hdr(M0+1,"BCDEFGHIJKL",mh)
months=[]; y,m=2025,8
while (y,m)<=(2026,11):
    months.append((y,m)); m+=1
    if m==13: y,m=y+1,1
for i,(y,m) in enumerate(months):
    r=M0+2+i; ws[f"B{r}"]=date(y,m,1); ws[f"B{r}"].number_format="mmm yyyy"
    for j,gname in enumerate(groups):
        c="CDEFGHIJK"[j]
        ws[f"{c}{r}"]=f'=SUMIFS({X("F")},{X("D")},"{gname}",{X("B")},">="&DATE({y},{m},1),{X("B")},"<="&EOMONTH(DATE({y},{m},1),0))'
        ws[f"{c}{r}"].number_format='#,##0;-#,##0;"–"'
    ws[f"L{r}"]=f"=SUM(C{r}:K{r})"; ws[f"L{r}"].number_format="#,##0"; ws[f"L{r}"].font=Font(bold=True)
    for c in "BCDEFGHIJKL":
        ws[f"{c}{r}"].border=Border(bottom=thin)
        if i%2: ws[f"{c}{r}"].fill=fill("F7F9FC")
ml=M0+2+len(months)-1; mt=ml+1
ws[f"B{mt}"]="TOTAL"
for c in "CDEFGHIJKL": ws[f"{c}{mt}"]=f"=SUM({c}{M0+2}:{c}{ml})"; ws[f"{c}{mt}"].number_format="#,##0"
for c in "BCDEFGHIJKL": ws[f"{c}{mt}"].font=Font(bold=True,color=WHITE); ws[f"{c}{mt}"].fill=fill("7A1C14")
ws.conditional_formatting.add(f"C{M0+2}:C{ml}",ColorScaleRule(start_type="min",start_color="FFFFFF",end_type="max",end_color="8DB4E2"))
ws.conditional_formatting.add(f"L{M0+2}:L{ml}",DataBarRule(start_type="num",start_value=0,end_type="max",color="E07A6B",showValue=True))
sb=BarChart(); sb.type="col"; sb.grouping="stacked"; sb.overlap=100; sb.title="Money out per month (AED)"
sb.add_data(Reference(ws,min_col=3,max_col=11,min_row=M0+1,max_row=ml),titles_from_data=True)
sb.set_categories(Reference(ws,min_col=2,min_row=M0+2,max_row=ml))
for s_,gname in zip(sb.series,groups): s_.graphicalProperties.solidFill=GC[gname]; s_.graphicalProperties.line.noFill=True
sb.y_axis.numFmt="#,##0"; sb.y_axis.majorGridlines=None; sb.x_axis.number_format="mmm yy"; sb.legend.position="b"
sb.height=9; sb.width=26
ws.add_chart(sb,f"B{mt+2}")
# not counted box
NB=mt+21
ws[f"B{NB}"]="NOT COUNTED – your own money moving (no double counting)"; ws[f"B{NB}"].font=Font(bold=True,size=12,color=NAVY)
nc=[("ENBD account → ENBD credit card bill payments",122853.84),("ENBD account → Emirates Islamic card bill payments",5853.00),("RAK → ENBD (instant transfers + TT)",67878.00),("ENBD → RAK (transfers)",35977.00),("RAK → RAK savings account (net)",2410.00),("ENBD → SNB Saudi account (SAR 110)",109.63),("Cash taken from RAK and re-deposited at ENBD (Thiran salary)",23000.00)]
for i,(l,v) in enumerate(nc):
    r=NB+1+i; ws[f"B{r}"]=l; ws[f"F{r}"]=v; ws[f"F{r}"].number_format="#,##0.00"
    ws[f"B{r}"].font=Font(size=9,color="5B6475"); ws[f"F{r}"].font=Font(size=9,color="5B6475")
    for c in "BCDEF": ws[f"{c}{r}"].border=Border(bottom=thin)
# detail
ws[f"B{DH-1}"]="EVERY OUTGOING PAYMENT (use the filter buttons)"; ws[f"B{DH-1}"].font=Font(bold=True,size=12,color=NAVY)
hdr(DH,"BCDEFG",["Date","Account","Bucket","Type","Amount (AED)","Details"])
for i,o in enumerate(ext):
    r=df+i; y,m,d=map(int,o['date'].split('-'))
    vals=[date(y,m,d),o['acct'],o['group'],o['sub'],o['amt'],o['desc']]
    for j,v in enumerate(vals):
        c="BCDEFG"[j]; x=ws[f"{c}{r}"]; x.value=v; x.border=Border(bottom=thin)
    ws[f"B{r}"].number_format="dd-mmm-yy"; ws[f"F{r}"].number_format="#,##0.00"
    ws[f"D{r}"].fill=fill(GP[o['group']]); ws[f"D{r}"].font=Font(bold=True,color=GC[o['group']],size=9)
    ws[f"G{r}"].font=Font(size=8,color="5B6475")
ws.auto_filter.ref=f"B{DH}:G{dl}"
ws[f"E{dl+1}"]="TOTAL"; ws[f"F{dl+1}"]=f"=SUM(F{df}:F{dl})"; ws[f"F{dl+1}"].number_format="#,##0.00"
for c in "BCDEFG": ws[f"{c}{dl+1}"].font=Font(bold=True,color=WHITE); ws[f"{c}{dl+1}"].fill=fill("7A1C14")
ws.freeze_panes="A3"
ws.page_setup.orientation="landscape"; ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0; ws.sheet_properties.pageSetUpPr.fitToPage=True
print("monthly rows",M0,ml,"not counted",NB,"detail",DH,dl)
wb.save(sys.argv[1])
