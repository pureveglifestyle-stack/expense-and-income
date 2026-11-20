from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import DataBarRule, FormulaRule, ColorScaleRule
from datetime import date
import calendar
rows=[
("2025-09-01","Thiran Tech",17938.54,"Aug 2025 salary","RAKBANK","Salary",17938.54,"AED"),
("2025-10-02","Thiran Tech",14786.93,"Sep 2025 salary","RAKBANK","Salary",14786.93,"AED"),
("2025-10-02","Thiran Tech",787.5,"Initial expenses reimbursement","RAKBANK","Salary",787.5,"AED"),
("2025-11-03","Thiran Tech",14786.93,"Oct 2025 salary","RAKBANK","Salary",14786.93,"AED"),
("2025-12-03","Thiran Tech",14786.93,"Nov 2025 salary","RAKBANK","Salary",14786.93,"AED"),
("2025-12-31","Thiran Tech",17170.93,"Dec 2025 salary + flight ticket","RAKBANK","Salary",17170.93,"AED"),
("2026-02-02","Thiran Tech",16500,"Jan 2026 salary","RAKBANK","Salary",16500,"AED"),
("2026-03-02","Thiran Tech",16500,"Feb 2026 salary","RAKBANK","Salary",16500,"AED"),
("2026-03-30","Thiran Tech",3800,"Final settlement","RAKBANK","Salary",3800,"AED"),
("2026-04-16","Envision",5000,"April 2026 salary (advance)","ENBD","Salary",5000,"AED"),
("2026-05-07","Envision",20000,"April 2026 salary (balance)","ENBD","Salary",20000,"AED"),
("2026-06-09","Envision",10000,"May 2026 salary (part 1)","ENBD","Salary",10000,"AED"),
("2026-06-12","Envision",15000,"May 2026 salary (part 2)","ENBD","Salary",15000,"AED"),
("2026-07-09","Envision",25000,"June 2026 salary","ENBD","Salary",25000,"AED"),
("2026-08-02","Envision (KSA)",2175.01,"July 2026 salary portion – Aram voucher 056, cash","Cash (KSA)","Salary",2258,"SAR"),
("2026-08-06","Envision",22825,"July 2026 salary (25,000 less 2,175.01 paid in KSA)","ENBD","Salary",22825,"AED"),
("2026-08-13","Envision",6193.74,"Travel expenses","ENBD","Travel allowance",6193.74,"AED"),
("2026-08-31","Envision (KSA)",4821.35,"August 2026 salary portion – Aram payroll","SNB (KSA)","Salary",5000,"SAR"),
("2026-09-04","Envision",14616.89,"Travel expenses","ENBD","Travel allowance",14616.89,"AED"),
("2026-09-05","Envision",20179,"August 2026 salary (25,000 less 4,821.35 paid in KSA)","ENBD","Salary",20179,"AED"),
("2026-10-04","Envision",25821,"September 2026 salary + full & final settlement","ENBD","Salary",25821,"AED"),
("2026-10-05","Envision (KSA)",4821.35,"September 2026 salary portion – Aram payroll (AED at Aug payslip rate)","SNB (KSA)","Salary",5000,"SAR"),
("2026-10-08","Envision",2087.73,"Travel expenses","ENBD","Travel allowance",2087.73,"AED"),
("2026-10-27","GenMahi AI",38000,"First salary, cash (37,500 deposited to ENBD)","Cash","Salary",38000,"AED"),
]
NAVY="12355B"; WHITE="FFFFFF"
COL={"Thiran Tech":"2F6DB5","Envision":"2E8B57","Envision (KSA)":"20A39E","GenMahi AI":"7B4FC9","Travel allowance":"E08A1E"}
PALE={"Thiran Tech":"DCE8F7","Envision":"DDF1E4","Envision (KSA)":"D6F2F0","GenMahi AI":"E9E1F7","Travel allowance":"FCE9D2"}
thin=Side(style="thin",color="C9CFD8")
def fill(h): return PatternFill("solid",fgColor=h)
wb=Workbook(); ws=wb.active; ws.title="1. Income"
ws.sheet_view.showGridLines=False
for col,w in zip("ABCDEFGHIJ",[2,16,22,16,46,13,16,14,9,2]): ws.column_dimensions[col].width=w
# banner
ws.merge_cells("B1:I1"); ws["B1"]="MY DUBAI MONEY JOURNEY  ·  PAGE 1: INCOME"
ws["B1"].font=Font(bold=True,size=16,color=WHITE); ws["B1"].fill=fill(NAVY); ws["B1"].alignment=Alignment(vertical="center",indent=1)
ws.row_dimensions[1].height=36
ws.merge_cells("B2:I2"); ws["B2"]="1 Sep 2025 – 19 Nov 2026  ·  all amounts in AED  ·  source: RAK, ENBD & SNB statements, Envision payslips, Aram voucher"
ws["B2"].font=Font(italic=True,size=9,color="5B6475")
# detail table location
D0=56; first=D0+1; last=first+len(rows)-1
R=lambda c:f"${c}${first}:${c}${last}"
# KPI tiles rows 4-6
tiles=[("B","C","TOTAL INCOME",f"=SUM({R('D')})","1D6B3A","E3F4E8","everything earned"),
       ("D","E","SALARY",f'=SUMIFS({R("D")},{R("G")},"Salary")',NAVY,"E6EDF6","Thiran + Envision + GenMahi"),
       ("F","G","TRAVEL ALLOWANCE",f'=SUMIFS({R("D")},{R("G")},"Travel allowance")',"B8650F","FCEFDF","Envision, for Saudi costs"),
       ("H","I","AVERAGE / MONTH","=B5/14.6","5A3BA8","EEE8FA","over 14.6 months")]
for a,b,lab,fx,dark,pale,note in tiles:
    for r in (4,5,6): ws.merge_cells(f"{a}{r}:{b}{r}")
    for r in (4,5,6):
        for c in (a,b): ws[f"{c}{r}"].fill=fill(pale)
    ws[f"{a}4"]=lab; ws[f"{a}4"].font=Font(bold=True,size=9,color=dark)
    ws[f"{a}5"]=fx; ws[f"{a}5"].font=Font(bold=True,size=20,color=dark); ws[f"{a}5"].number_format="#,##0"
    ws[f"{a}6"]=note; ws[f"{a}6"].font=Font(size=8,color="5B6475")
    for r in (4,5,6): ws[f"{a}{r}"].alignment=Alignment(horizontal="center",vertical="center")
ws.row_dimensions[5].height=34
# fix tile anchors (merged cell text lives in first col)
ws["B5"].value=tiles[0][3]
# company table
def hdr(r,cols,labels):
    for c,l in zip(cols,labels):
        x=ws[f"{c}{r}"]; x.value=l; x.font=Font(bold=True,color=WHITE,size=10); x.fill=fill(NAVY); x.alignment=Alignment(horizontal="center")
ws["B8"]="WHO PAID ME"; ws["B8"].font=Font(bold=True,size=12,color=NAVY)
hdr(9,"BCDE",["Company","Amount (AED)","Share","What it is"])
comp=[("Thiran Tech","Thiran Tech","Salary Aug 2025 – Feb 2026 + settlement"),
      ("Envision","Envision","Salary Apr – Sep 2026 (paid to ENBD)"),
      ("Envision (KSA)","Envision (KSA)","Salary part paid in Saudi via Aram"),
      ("GenMahi AI","GenMahi AI","First salary, cash (Oct 2026)"),
      ("Travel allowance","Travel allowance","Envision travel money for Saudi")]
for i,(lab,key,desc) in enumerate(comp):
    r=10+i
    ws[f"B{r}"]=lab; ws[f"B{r}"].font=Font(bold=True,color=COL[key]); ws[f"B{r}"].fill=fill(PALE[key])
    if key=="Travel allowance": fx=f'=SUMIFS({R("D")},{R("G")},"Travel allowance")'
    elif key=="Envision": fx=f'=SUMIFS({R("D")},{R("C")},"Envision",{R("G")},"Salary")'
    else: fx=f'=SUMIFS({R("D")},{R("C")},"{key}")'
    ws[f"C{r}"]=fx; ws[f"C{r}"].number_format="#,##0.00"
    ws[f"D{r}"]=f"=C{r}/$C$15"; ws[f"D{r}"].number_format="0.0%"
    ws[f"E{r}"]=desc; ws[f"E{r}"].font=Font(size=9,color="5B6475")
    for c in "BCDE": ws[f"{c}{r}"].border=Border(bottom=thin)
ws["B15"]="TOTAL"; ws["C15"]="=SUM(C10:C14)"; ws["D15"]="=SUM(D10:D14)"
for c in "BCD": ws[f"{c}15"].font=Font(bold=True); ws[f"{c}15"].border=Border(top=Side(style="medium",color=NAVY))
ws["C15"].number_format="#,##0.00"; ws["D15"].number_format="0%"
ws.conditional_formatting.add("C10:C14",DataBarRule(start_type="num",start_value=0,end_type="max",color="5B9BD5",showValue=True))
# pie chart
pie=PieChart(); pie.title="Income by source"
pie.add_data(Reference(ws,min_col=3,min_row=9,max_row=14),titles_from_data=True)
pie.set_categories(Reference(ws,min_col=2,min_row=10,max_row=14))
pie.dataLabels=DataLabelList(); pie.dataLabels.showPercent=True; pie.dataLabels.showVal=False; pie.dataLabels.showCatName=False; pie.dataLabels.showSerName=False; pie.dataLabels.showLegendKey=False
from openpyxl.chart.series import DataPoint
for idx,(lab,key,_) in enumerate(comp):
    pt=DataPoint(idx=idx); pt.graphicalProperties.solidFill=COL[key]; pie.series[0].dPt.append(pt)
pie.height=7.5; pie.width=11
ws.add_chart(pie,"F8")
# monthly table
ws["B18"]="MONTH BY MONTH"; ws["B18"].font=Font(bold=True,size=12,color=NAVY)
mcols=["Thiran Tech","Envision","Envision (KSA)","GenMahi AI","Travel allowance"]
hdr(19,"BCDEFGH",["Month","Thiran","Envision","Envision KSA","GenMahi","Travel","Total"])
months=[]; y,m=2025,9
while (y,m)<=(2026,11):
    months.append((y,m)); m+=1
    if m==13: y,m=y+1,1
for i,(y,m) in enumerate(months):
    r=20+i
    ws[f"B{r}"]=date(y,m,1); ws[f"B{r}"].number_format="mmm yyyy"
    for j,key in enumerate(mcols):
        c="CDEFG"[j]
        start=f"DATE({y},{m},1)"; end=f"EOMONTH(DATE({y},{m},1),0)"
        if key=="Travel allowance": cond=f'{R("G")},"Travel allowance"'
        elif key=="Envision": cond=f'{R("C")},"Envision",{R("G")},"Salary"'
        else: cond=f'{R("C")},"{key}"'
        ws[f"{c}{r}"]=f'=SUMIFS({R("D")},{cond},{R("B")},">="&{start},{R("B")},"<="&{end})'
        ws[f"{c}{r}"].number_format='#,##0;-#,##0;"–"'
    ws[f"H{r}"]=f"=SUM(C{r}:G{r})"; ws[f"H{r}"].number_format='#,##0;-#,##0;"–"'; ws[f"H{r}"].font=Font(bold=True)
    for c in "BCDEFGH":
        ws[f"{c}{r}"].border=Border(bottom=thin)
        if i%2: ws[f"{c}{r}"].fill=fill("F7F9FC")
lm=20+len(months)-1
ws.conditional_formatting.add(f"H20:H{lm}",DataBarRule(start_type="num",start_value=0,end_type="max",color="63BE7B",showValue=True))
bar=BarChart(); bar.type="col"; bar.grouping="stacked"; bar.overlap=100; bar.title="Income per month (AED)"
bar.add_data(Reference(ws,min_col=3,max_col=7,min_row=19,max_row=lm),titles_from_data=True)
bar.set_categories(Reference(ws,min_col=2,min_row=20,max_row=lm))
for s,key in zip(bar.series,mcols): s.graphicalProperties.solidFill=COL[key]; s.graphicalProperties.line.noFill=True
bar.y_axis.numFmt="#,##0"; bar.y_axis.majorGridlines=None; bar.x_axis.number_format="mmm yy"
bar.height=8; bar.width=24; bar.legend.position="b"
ws.add_chart(bar,f"B{lm+2}")
# detail table
ws[f"B{D0-1}"]="EVERY PAYMENT RECEIVED"; ws[f"B{D0-1}"].font=Font(bold=True,size=12,color=NAVY)
hdr(D0,"BCDEFGHI",["Date","Company","Amount (AED)","Description","Paid into","Category","Original","Currency"])
for i,rw in enumerate(rows):
    r=first+i; key="Travel allowance" if rw[5]=="Travel allowance" else rw[1]
    y,m,d=map(int,rw[0].split("-"))
    vals=[date(y,m,d)]+list(rw[1:])
    for j,v in enumerate(vals):
        c="BCDEFGHI"[j]; x=ws[f"{c}{r}"]; x.value=v; x.border=Border(bottom=thin)
        if c=="B": x.number_format="dd-mmm-yy"
        if c in "DH": x.number_format="#,##0.00"
        if c=="E": x.alignment=Alignment(wrap_text=True,vertical="top")
    ws[f"C{r}"].fill=fill(PALE[key]); ws[f"C{r}"].font=Font(bold=True,color=COL[key])
    ws[f"G{r}"].font=Font(color=COL[key])
ws.conditional_formatting.add(f"D{first}:D{last}",DataBarRule(start_type="num",start_value=0,end_type="max",color="9BC2E6",showValue=True))
tr=last+1
ws[f"C{tr}"]="TOTAL"; ws[f"D{tr}"]=f"=SUM(D{first}:D{last})"; ws[f"D{tr}"].number_format="#,##0.00"
for c in "BCDEFGHI": ws[f"{c}{tr}"].font=Font(bold=True,color=WHITE); ws[f"{c}{tr}"].fill=fill("1D6B3A")
ws.freeze_panes="A3"
ws.page_setup.orientation="landscape"; ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0; ws.sheet_properties.pageSetUpPr.fitToPage=True
ws.sheet_properties.tabColor="1D6B3A"
wb.calculation.fullCalcOnLoad=True
import sys; wb.save(sys.argv[1]); print("ok rows",first,last,"month rows",lm)
