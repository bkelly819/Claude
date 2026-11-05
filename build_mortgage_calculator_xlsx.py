import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter as L
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.comments import Comment

OUT = "/home/user/Claude/Mortgage_Paydown_vs_Invest_Calculator.xlsx"
NROWS = 480  # 40 years of months

F = "Arial"
BLUE = Font(name=F, color="0000FF")
BLACK = Font(name=F)
BOLD = Font(name=F, bold=True)
GREEN = Font(name=F, color="008000")
TITLE = Font(name=F, bold=True, size=16)
H2 = Font(name=F, bold=True, size=12)
HDR = Font(name=F, bold=True, color="FFFFFF")
MUTED = Font(name=F, italic=True, color="595959", size=9)
YELLOW = PatternFill("solid", fgColor="FFFF00")
HDRFILL = PatternFill("solid", fgColor="1D5C55")
SECFILL = PatternFill("solid", fgColor="E3EFEC")
AFILL = PatternFill("solid", fgColor="DCE8F8")
BFILL = PatternFill("solid", fgColor="FBE3D9")
thin = Side(style="thin", color="BFBFBF")
BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
USD = '$#,##0;($#,##0);"-"'
USD2 = '$#,##0.00;($#,##0.00);"-"'
PCT = '0.00%;(0.00%);"-"'
PCT1 = '0.0%;(0.0%);"-"'
MON = 'mmm yyyy'

wb = Workbook()

def name(n, ref):
    wb.defined_names[n] = DefinedName(n, attr_text=ref)

def style_range(ws, rng, font=None, fill=None, fmt=None, border=None, align=None):
    for row in ws[rng]:
        for c in row:
            if font: c.font = font
            if fill: c.fill = fill
            if fmt: c.number_format = fmt
            if border: c.border = border
            if align: c.alignment = align

# ======================= INPUTS =======================
wsI = wb.active
wsI.title = "Inputs"
wsI.sheet_view.showGridLines = False
wsI.column_dimensions["A"].width = 2
wsI.column_dimensions["B"].width = 44
wsI.column_dimensions["C"].width = 18
wsI.column_dimensions["D"].width = 16
wsI.column_dimensions["E"].width = 70
wsI["B1"] = "Mortgage Paydown vs. Investing Calculator: Inputs"; wsI["B1"].font = TITLE
wsI["B2"] = "Prepared for John-Henry Scherck · Analysis date Nov 5, 2026"; wsI["B2"].font = MUTED
wsI["B3"] = "Legend: blue text = input you can change · yellow fill = needed / confirm (not in the email) · black = calculated. Blank tax rates count as 0%."
wsI["B3"].font = MUTED

dv_yes = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False)
dv_filing = DataValidation(type="list", formula1='"Married filing jointly,Single,Head of household,Married filing separately"', allow_blank=True)
dv_item = DataValidation(type="list", formula1='"Compare to standard,Yes,No"')
dv_fed = DataValidation(type="list", formula1='"10%,12%,22%,24%,32%,35%,37%"', allow_blank=True)
dv_ltcg = DataValidation(type="list", formula1='"0%,15%,20%"', allow_blank=True)
dv_freq = DataValidation(type="list", formula1='"Monthly,Annually"')
dv_stop = DataValidation(type="list", formula1='"Until payoff,For a set period"')
for d in (dv_yes, dv_filing, dv_item, dv_fed, dv_ltcg, dv_freq, dv_stop):
    wsI.add_data_validation(d)

row = 5
def section(title):
    global row
    row += 1
    c = wsI.cell(row=row, column=2, value=title); c.font = H2
    for col in range(2, 6): wsI.cell(row=row, column=col).fill = SECFILL
    wsI.cell(row=row, column=3, value="Value").font = BOLD
    wsI.cell(row=row, column=4, value="Source").font = BOLD
    wsI.cell(row=row, column=5, value="Notes").font = BOLD
    row += 1

def inp(label, nm, value, fmt, source, note, kind="input", dv=None):
    """kind: input (blue), need (blue+yellow), calc (black formula)"""
    global row
    wsI.cell(row=row, column=2, value=label).font = BLACK
    c = wsI.cell(row=row, column=3, value=value)
    c.number_format = fmt
    c.border = BOX
    c.alignment = Alignment(horizontal="right")
    if kind == "calc":
        c.font = BLACK
    else:
        c.font = BLUE
    if kind == "need":
        c.fill = YELLOW
    wsI.cell(row=row, column=4, value=source).font = Font(name=F, size=9, color="1D5C55" if source.startswith("Email") else "595959")
    n = wsI.cell(row=row, column=5, value=note); n.font = Font(name=F, size=9); n.alignment = Alignment(wrap_text=True, vertical="top")
    if dv: dv.add(c)
    if nm: name(nm, f"Inputs!$C${row}")
    row += 1

section("1. Mortgage")
inp("Current mortgage balance", "Loan", 1070000, USD, "Email", '"About $1,070,000", may come down slightly after negotiations. Confirm at closing.')
inp("Interest rate", "Rate", 0.07, PCT, "Email", "30-year fixed.")
inp("Remaining term (months)", "TermM", 360, "0", "Email", "New 30-year loan = 360 months.")
inp("Monthly payment override (optional)", "PmtOverride", None, USD2, "Optional", "Leave blank to use the calculated payment below.")
inp("Monthly payment used (P&I)", "Pmt", "=IF(AND(ISNUMBER(PmtOverride),PmtOverride>Loan*Rate/12),PmtOverride,PMT(Rate/12,TermM,-Loan))", USD2, "Calculated", "Principal and interest only; taxes and insurance are the same in both strategies.", kind="calc")
inp("First payment month", "FirstPay", dt.date(2027, 1, 1), MON, "Confirm", "Closing date not set yet. Placeholder; drives payoff dates.", kind="need")

section("2. Extra principal payments (Option A)")
inp("Extra payment per month", "Extra", 2500, USD, "Email: $2–3K", 'Email: "an extra 2-3K a month". Try 2,000 / 2,500 / 3,000.')
inp("Extra payments start month", "ExStart", dt.date(2027, 1, 1), MON, "Assumption", "Defaults to the first payment month.")
inp("Raise extra payment each year by", "Grow", 0, PCT1, "Assumption", "0% = flat.")
inp("Continue extra payments", "StopMode", "Until payoff", "@", "Assumption", "Until payoff, or For a set period (uses the years below).", dv=dv_stop)
inp("Years of extra payments (if set period)", "ExYears", 10, "0", "Assumption", "")
inp("Extra start offset (months)", "ExStartIdx", "=MAX(0,(YEAR(ExStart)-YEAR(FirstPay))*12+MONTH(ExStart)-MONTH(FirstPay))", "0", "Calculated", "", kind="calc")
lump_hdr = row
wsI.cell(row=row, column=2, value="One-time principal payments (month, amount)").font = BOLD
wsI.cell(row=row, column=5, value="None proposed in the email. Under Option B each lump sum is invested on the same date instead.").font = Font(name=F, size=9)
row += 1
wsI.cell(row=row, column=2, value="Month (first of month)").font = MUTED
wsI.cell(row=row, column=3, value="Amount").font = MUTED
wsI.cell(row=row, column=4, value="Month # (calc)").font = MUTED
row += 1
lump_first = row
for i in range(10):
    r = row + i
    wsI.cell(row=r, column=2).number_format = MON
    wsI.cell(row=r, column=2).font = BLUE
    wsI.cell(row=r, column=2).border = BOX
    wsI.cell(row=r, column=3).number_format = USD
    wsI.cell(row=r, column=3).font = BLUE
    wsI.cell(row=r, column=3).border = BOX
    c = wsI.cell(row=r, column=4, value=f'=IF(OR(B{r}="",C{r}=""),-1,(YEAR(B{r})-YEAR(FirstPay))*12+MONTH(B{r})-MONTH(FirstPay))')
    c.font = BLACK
row += 10
name("LumpIdx", f"Inputs!$D${lump_first}:$D${lump_first+9}")
name("LumpAmt", f"Inputs!$C${lump_first}:$C${lump_first+9}")

section("3. Investing and comparison period (Option B)")
inp("Expected annual return", "Ret", 0.07, PCT, "Assumption", "Planner assumption; not in the email.", kind="need")
inp("Taxable yield (dividends/interest)", "Yld", 0.015, PCT, "Assumption", "Part of the total return, taxed every year.")
inp("Contribution frequency", "Freq", "Monthly", "@", "Assumption", "Monthly, or Annually (deposited at each year end).", dv=dv_freq)
inp("Comparison period (years, 1–40)", "Horizon", 30, "0", "Assumption", "Both strategies are measured at the same end date.")
inp("Invest full payment once a loan is paid off?", "InvestFreed", "Yes", "@", "Assumption", "Yes keeps the cash outlay equal for the whole period. No holds it as cash.", dv=dv_yes)
inp("Invest mortgage-interest tax savings each year?", "ReinvestTax", "Yes", "@", "Assumption", "No holds the savings as cash.", dv=dv_yes)
inp("Monthly growth rate", "GM", "=(1+Ret)^(1/12)-1", "0.0000%", "Calculated", "", kind="calc")

section("4. Taxes")
inp("Filing status", "Filing", None, "@", "Needed", "Not in the email. Sets the standard deduction if you leave the override blank.", kind="need", dv=dv_filing)
inp("Itemizes deductions?", "Itemize", "Compare to standard", "@", "Assumption", "Compare to standard: counts only what itemizing adds above the standard deduction.", dv=dv_item)
inp("Federal marginal bracket", "FedRate", None, PCT1, "Needed", "Not in the email.", kind="need", dv=dv_fed)
inp("State marginal bracket", "StateRate", None, PCT, "Needed", "Not in the email. Client signs from Los Angeles, CA.", kind="need")
inp("Federal standard deduction override", "StdFedIn", None, USD, "Optional", "Blank = table value for the filing status.")
inp("State standard deduction override", "StdStIn", None, USD, "Optional", "Blank = table value for the filing status.")
inp("Other itemized deductions, federal", "OtherFed", None, USD, "Optional", "SALT (capped), charitable gifts, etc.")
inp("Other itemized deductions, state", "OtherSt", None, USD, "Optional", "")
inp("Federal mortgage debt limit", "LimFed", 750000, USD, "Tax law", "$750K ($375K if married filing separately).")
inp("State mortgage debt limit", "LimSt", 1000000, USD, "Email: CA", "California $1M ($500K if MFS).")
inp("Cap deduction value at 35% for 37% bracket?", "Cap35", "Yes", "@", "Tax law", "2026+ limit on itemized deduction value for the top bracket.", dv=dv_yes)
inp("Federal LTCG / qualified dividend rate", "LTCG", None, PCT1, "Needed", "Not in the email.", kind="need", dv=dv_ltcg)
inp("Add 3.8% NIIT?", "NIIT", "No", "@", "Needed", "Depends on income.", kind="need", dv=dv_yes)
inp("State taxes investment gains as income?", "StateGains", "Yes", "@", "Tax law", "California taxes capital gains as ordinary income.", dv=dv_yes)
inp("Tax unrealized gains at end (as if sold)?", "Liquidate", "Yes", "@", "Assumption", "", dv=dv_yes)
std_tbl = row + 1
inp("Federal standard deduction used", "StdFed", None, USD, "Calculated", "", kind="calc")
inp("State standard deduction used", "StdSt", None, USD, "Calculated", "", kind="calc")
inp("Federal rate applied to deductions", "FedDedRate", '=IF(AND(Cap35="Yes",FedRate>=0.37),0.35,FedRate)', PCT1, "Calculated", "", kind="calc")
inp("Total tax rate on investment income", "InvRate", '=LTCG+IF(NIIT="Yes",0.038,0)+IF(StateGains="Yes",StateRate,0)', PCT, "Calculated", "", kind="calc")

section("5. Refinance (optional)")
inp("Refinance after year", "RefiYr", None, "0", "Optional", "Blank = no refinance. Balance re-amortized to the original end date; no closing costs.")
inp("New rate", "RefiRate", None, PCT, "Optional", "")

# standard deduction reference table
row += 1
wsI.cell(row=row, column=2, value="Standard deduction reference (verify each year)").font = H2
row += 1
for j, h in enumerate(["Filing status", "Federal 2026", "California 2025"]):
    c = wsI.cell(row=row, column=2 + j, value=h); c.font = BOLD
row += 1
tbl_first = row
for fs, fed, st in [("Married filing jointly", 32200, 11412), ("Single", 16100, 5706), ("Head of household", 24150, 11412), ("Married filing separately", 16100, 5706)]:
    wsI.cell(row=row, column=2, value=fs).font = BLACK
    for j, v in enumerate([fed, st]):
        c = wsI.cell(row=row, column=3 + j, value=v); c.font = BLUE; c.number_format = USD
    row += 1
tbl = f"Inputs!$B${tbl_first}:$D${row-1}"
wsI.cell(row=row, column=2, value="Sources: IRS 2026 standard deduction (One Big Beautiful Bill Act); California FTB 2025 standard deduction.").font = MUTED
# fill std used formulas
for r in range(1, wsI.max_row + 1):
    if wsI.cell(row=r, column=2).value == "Federal standard deduction used":
        wsI.cell(row=r, column=3).value = f'=IF(ISNUMBER(StdFedIn),StdFedIn,IFERROR(VLOOKUP(Filing,{tbl},2,FALSE),0))'
    if wsI.cell(row=r, column=2).value == "State standard deduction used":
        wsI.cell(row=r, column=3).value = f'=IF(ISNUMBER(StdStIn),StdStIn,IFERROR(VLOOKUP(Filing,{tbl},3,FALSE),0))'
wsI.freeze_panes = "A5"

# ======================= MONTHLY MODEL =======================
wsM = wb.create_sheet("Monthly Model")
cols = [
 ("m", "Month #", "0"), ("date", "Date", MON), ("yr", "Year", "0"), ("inper", "In period (1/0)", "0"), ("rate", "Rate", PCT),
 ("exrec", "Recurring extra", USD), ("lump", "Lump sum", USD), ("extra", "Extra this month", USD),
 ("bBeg", "B begin bal", USD), ("bLvl", "B level P&I", USD2), ("bInt", "B interest", USD), ("bPay", "B payment", USD), ("bEnd", "B end bal", USD),
 ("budget", "Shared monthly budget", USD),
 ("aBeg", "A begin bal", USD), ("aInt", "A interest", USD), ("aPay", "A payment", USD), ("aEnd", "A end bal", USD), ("aXL", "A extra to principal", USD),
 ("aInv", "A to invest", USD), ("aCsh", "A to cash", USD), ("bInv", "B to invest", USD), ("bCsh", "B to cash", USD),
 ("aYI", "A yr interest", USD), ("aYA", "A yr avg bal", USD), ("aTF", "A fed tax saving", USD), ("aTS", "A state tax saving", USD), ("aT", "A tax saving", USD),
 ("bYI", "B yr interest", USD), ("bYA", "B yr avg bal", USD), ("bTF", "B fed tax saving", USD), ("bTS", "B state tax saving", USD), ("bT", "B tax saving", USD),
 ("aDA", "A div accrual", USD), ("aYD", "A yr dividends", USD), ("aPd", "A pending", USD), ("aDp", "A deposit", USD), ("aDT", "A div tax", USD), ("aTR", "A tax savings reinvested", USD),
 ("aPt", "A portfolio", USD), ("aBs", "A cost basis", USD), ("aCB", "A cash balance", USD), ("aNW", "A after-tax net worth", USD),
 ("bDA", "B div accrual", USD), ("bYD", "B yr dividends", USD), ("bPd", "B pending", USD), ("bDp", "B deposit", USD), ("bDT", "B div tax", USD), ("bTR", "B tax savings reinvested", USD),
 ("bPt", "B portfolio", USD), ("bBs", "B cost basis", USD), ("bCB", "B cash balance", USD), ("bNW", "B after-tax net worth", USD),
]
C = {k: L(i + 1) for i, (k, _, _) in enumerate(cols)}
FIRST, LAST = 2, NROWS + 1
for i, (k, h, fmt) in enumerate(cols):
    c = wsM.cell(row=1, column=i + 1, value=h)
    c.font = HDR; c.fill = HDRFILL; c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    wsM.column_dimensions[L(i + 1)].width = 13
wsM.row_dimensions[1].height = 42
wsM.freeze_panes = "C2"

def taxf(prefix, which):
    yi, ya = C[prefix + "YI"], C[prefix + "YA"]
    lim, std, oth, rt = ("LimFed", "StdFed", "OtherFed", "FedDedRate") if which == "F" else ("LimSt", "StdSt", "OtherSt", "StateRate")
    return lambda r: (f'=IF(OR(Itemize="No",{yi}{r}<=0),0,IF(Itemize="Yes",{yi}{r}*IF({ya}{r}>0,MIN(1,{lim}/{ya}{r}),1),'
                      f'MAX({std},{oth}+{yi}{r}*IF({ya}{r}>0,MIN(1,{lim}/{ya}{r}),1))-MAX({std},{oth}))*{rt})')

for r in range(FIRST, LAST + 1):
    p = r - 1
    first = r == FIRST
    s = max(FIRST, r - 11)
    m = f"{C['m']}{r}"
    ye = f"MOD({m},12)=0"
    f = {}
    f["m"] = r - 1
    f["date"] = f"=EDATE(FirstPay,{m}-1)"
    f["yr"] = f"=INT(({m}-1)/12)+1"
    f["inper"] = f"=IF({m}<=Horizon*12,1,0)"
    f["rate"] = f"=IF(AND(RefiYr>0,ISNUMBER(RefiRate),{m}>RefiYr*12),RefiRate,Rate)"
    f["exrec"] = f'=IF(AND({m}-1>=ExStartIdx,OR(StopMode="Until payoff",{m}-1-ExStartIdx<ExYears*12)),Extra*(1+Grow)^INT(({m}-1-ExStartIdx)/12),0)'
    f["lump"] = f"=SUMIF(LumpIdx,{m}-1,LumpAmt)"
    f["extra"] = f"={C['exrec']}{r}+{C['lump']}{r}"
    f["bBeg"] = "=Loan" if first else f"={C['bEnd']}{p}"
    f["bLvl"] = "=Pmt" if first else (f"=IF(AND(RefiYr>0,ISNUMBER(RefiRate),{m}=RefiYr*12+1),IF({C['bBeg']}{r}>0.005,PMT({C['rate']}{r}/12,MAX(1,TermM-({m}-1)),-{C['bBeg']}{r}),0),{C['bLvl']}{p})")
    f["bInt"] = f"=IF({C['bBeg']}{r}>0.005,{C['bBeg']}{r}*{C['rate']}{r}/12,0)"
    f["bPay"] = f"=IF({C['bBeg']}{r}>0.005,MIN({C['bLvl']}{r},{C['bBeg']}{r}+{C['bInt']}{r}),0)"
    f["bEnd"] = f"=IF({C['bBeg']}{r}+{C['bInt']}{r}-{C['bPay']}{r}<=0.005,0,{C['bBeg']}{r}+{C['bInt']}{r}-{C['bPay']}{r})"
    f["budget"] = f"={C['bPay']}{r}+{C['extra']}{r}+IF({C['bBeg']}{r}>0.005,0,{C['bLvl']}{r})"
    f["aBeg"] = "=Loan" if first else f"={C['aEnd']}{p}"
    f["aInt"] = f"=IF({C['aBeg']}{r}>0.005,{C['aBeg']}{r}*{C['rate']}{r}/12,0)"
    f["aPay"] = f"=IF({C['aBeg']}{r}>0.005,MIN({C['budget']}{r},{C['aBeg']}{r}+{C['aInt']}{r}),0)"
    f["aEnd"] = f"=IF({C['aBeg']}{r}+{C['aInt']}{r}-{C['aPay']}{r}<=0.005,0,{C['aBeg']}{r}+{C['aInt']}{r}-{C['aPay']}{r})"
    f["aXL"] = f"=IF({C['aBeg']}{r}>0.005,MAX(0,{C['aPay']}{r}-MIN({C['bPay']}{r},{C['aBeg']}{r}+{C['aInt']}{r})),0)"
    f["aInv"] = f'=IF(InvestFreed="Yes",MAX(0,{C["budget"]}{r}-{C["aPay"]}{r}),0)'
    f["aCsh"] = f'=IF(InvestFreed="Yes",0,MAX(0,{C["budget"]}{r}-{C["aPay"]}{r}))'
    f["bInv"] = f'={C["extra"]}{r}+IF(InvestFreed="Yes",MAX(0,{C["budget"]}{r}-{C["bPay"]}{r}-{C["extra"]}{r}),0)'
    f["bCsh"] = f'=IF(InvestFreed="Yes",0,MAX(0,{C["budget"]}{r}-{C["bPay"]}{r}-{C["extra"]}{r}))'
    for x in "ab":
        beg, it = C[x + "Beg"], C[x + "Int"]
        f[x + "YI"] = f"=IF({ye},SUM({it}{s}:{it}{r}),0)"
        f[x + "YA"] = f'=IF({ye},SUMIF({beg}{s}:{beg}{r},">0.005")/MAX(1,COUNTIF({beg}{s}:{beg}{r},">0.005")),0)'
        f[x + "TF"] = taxf(x, "F")(r)
        f[x + "TS"] = taxf(x, "S")(r)
        f[x + "T"] = f"={C[x+'TF']}{r}+{C[x+'TS']}{r}"
        pt, bs, pd, cb = C[x + "Pt"], C[x + "Bs"], C[x + "Pd"], C[x + "CB"]
        inv, csh = C[x + "Inv"], C[x + "Csh"]
        prevpt = "0" if first else f"{pt}{p}"
        prevpd = "0" if first else f"{pd}{p}"
        prevbs = "0" if first else f"{bs}{p}"
        prevcb = "0" if first else f"{cb}{p}"
        f[x + "DA"] = f"={prevpt}*Yld/12"
        f[x + "YD"] = f"=IF({ye},SUM({C[x+'DA']}{s}:{C[x+'DA']}{r}),0)"
        f[x + "Pd"] = f'=IF(Freq="Annually",IF({ye},0,{prevpd}+{inv}{r}),0)'
        f[x + "Dp"] = f'=IF(Freq="Annually",IF({ye},{prevpd}+{inv}{r},0),{inv}{r})'
        f[x + "DT"] = f"={C[x+'YD']}{r}*InvRate"
        f[x + "TR"] = f'=IF(ReinvestTax="Yes",{C[x+"T"]}{r},0)'
        f[x + "Pt"] = f"={prevpt}*(1+GM)+{C[x+'Dp']}{r}-{C[x+'DT']}{r}+{C[x+'TR']}{r}"
        f[x + "Bs"] = f"={prevbs}+{C[x+'Dp']}{r}+{C[x+'YD']}{r}+{C[x+'TR']}{r}"
        f[x + "CB"] = f'={prevcb}+{csh}{r}+IF(ReinvestTax="Yes",0,{C[x+"T"]}{r})'
        f[x + "NW"] = f'={pt}{r}-IF(Liquidate="Yes",MAX(0,{pt}{r}-{bs}{r})*InvRate,0)+{cb}{r}-{C[x+"End"]}{r}'
    for i, (k, _, fmt) in enumerate(cols):
        c = wsM.cell(row=r, column=i + 1, value=f[k])
        c.number_format = fmt
        c.font = BLACK

def rng(k):
    return f"'Monthly Model'!${C[k]}${FIRST}:${C[k]}${LAST}"
def at_h(k):
    return f"INDEX({rng(k)},Horizon*12)"
def tot(k):
    return f"SUMIFS({rng(k)},{rng('inper')},1)"

# ======================= ANNUAL =======================
wsA = wb.create_sheet("Annual Schedule")
ah = ["Year", "Year ending", "In period", "A mortgage balance", "B mortgage balance", "A investments", "B investments",
      "A net worth (after tax)", "B net worth (after tax)", "B minus A", "A interest", "B interest", "A tax savings", "B tax savings"]
for i, h in enumerate(ah):
    c = wsA.cell(row=1, column=i + 1, value=h)
    c.font = HDR; c.fill = HDRFILL; c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    wsA.column_dimensions[L(i + 1)].width = 15
wsA.row_dimensions[1].height = 32
wsA.freeze_panes = "B2"
amap = ["aEnd", "bEnd", "aPt", "bPt", "aNW", "bNW"]
for y in range(1, 41):
    r = y + 1
    mr = 12 * y + 1  # monthly row of year-end
    wsA.cell(row=r, column=1, value=y).font = BLACK
    c = wsA.cell(row=r, column=2, value=f"='Monthly Model'!{C['date']}{mr}"); c.number_format = MON; c.font = GREEN
    wsA.cell(row=r, column=3, value=f'=IF(A{r}<=Horizon,"Yes","")').font = BLACK
    for j, k in enumerate(amap):
        c = wsA.cell(row=r, column=4 + j, value=f"='Monthly Model'!{C[k]}{mr}"); c.number_format = USD; c.font = GREEN
    c = wsA.cell(row=r, column=10, value=f"=I{r}-H{r}"); c.number_format = USD; c.font = BLACK
    for j, k in enumerate(["aYI", "bYI", "aT", "bT"]):
        c = wsA.cell(row=r, column=11 + j, value=f"='Monthly Model'!{C[k]}{mr}"); c.number_format = USD; c.font = GREEN
for r in range(2, 42):
    for col in (4, 6, 8, 11, 13):
        wsA.cell(row=r, column=col).fill = AFILL
    for col in (5, 7, 9, 12, 14):
        wsA.cell(row=r, column=col).fill = BFILL

# ======================= COMPARISON (dashboard) =======================
ws = wb.create_sheet("Comparison", 0)
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 2
ws.column_dimensions["B"].width = 46
for col in "CDE": ws.column_dimensions[col].width = 20
ws.column_dimensions["F"].width = 3
ws["B1"] = "Pay down the 7% mortgage, or invest the difference?"; ws["B1"].font = TITLE
ws["B2"] = "Prepared for John-Henry Scherck · Mortgage paydown vs. investing · Analysis date Nov 5, 2026"; ws["B2"].font = MUTED
ws["B3"] = '=IF(OR(Filing="",FedRate="",StateRate="",LTCG=""),"Tax inputs needed on the Inputs sheet (filing status, federal bracket, state bracket, capital gains rate). Blank rates count as 0%, so those taxes are left out until entered.","All tax inputs entered.")'
ws["B3"].font = Font(name=F, bold=True, color="7F6000"); ws["B3"].fill = PatternFill("solid", fgColor="FDF3DC")
ws.merge_cells("B3:E3"); ws["B3"].alignment = Alignment(wrap_text=True, vertical="center"); ws.row_dimensions[3].height = 32

ws["B5"] = "FINAL COMPARISON · AFTER-TAX NET WORTH AT END OF PERIOD"; ws["B5"].font = Font(name=F, bold=True, size=9, color="595959")
ws["B6"] = ('=IF(ABS(D9)<1,"Both strategies end in the same place.",IF(D9>0,"Investing "&TEXT(Extra,"$#,##0")&"/mo instead of prepaying leaves an estimated "&TEXT(D9,"$#,##0")&" more after taxes over "&Horizon&" years.",'
            '"Paying down the mortgage with "&TEXT(Extra,"$#,##0")&"/mo leaves an estimated "&TEXT(-D9,"$#,##0")&" more after taxes over "&Horizon&" years."))')
ws["B6"].font = Font(name=F, bold=True, size=14); ws.merge_cells("B6:E6"); ws["B6"].alignment = Alignment(wrap_text=True, vertical="center"); ws.row_dimensions[6].height = 40
for j, h in enumerate(["Option A · Pay down", "Option B · Invest", "Difference (B − A)", "Higher strategy"]):
    c = ws.cell(row=8, column=2 + j, value=h); c.font = HDR; c.fill = HDRFILL; c.alignment = Alignment(horizontal="center")
ws["B9"] = "=C25"; ws["C9"] = "=D25"; ws["D9"] = "=C9-B9"
ws["E9"] = '=IF(ABS(D9)<1,"Even",IF(D9>0,"B · Invest","A · Pay down"))'
for c in ("B9", "C9", "D9"):
    ws[c].number_format = USD; ws[c].font = Font(name=F, bold=True, size=14); ws[c].alignment = Alignment(horizontal="center")
ws["E9"].font = Font(name=F, bold=True, size=12); ws["E9"].alignment = Alignment(horizontal="center")
ws["B9"].fill = AFILL; ws["C9"].fill = BFILL
ws.row_dimensions[9].height = 26

ws["B11"] = "Side-by-side"; ws["B11"].font = H2
ws["B12"] = '="Both options start "&TEXT(FirstPay,"mmm yyyy")&" and are measured at "&TEXT(EDATE(FirstPay,Horizon*12-1),"mmm yyyy")&"."'
ws["B12"].font = MUTED
for j, h in enumerate(["Measure", "A · Pay down", "B · Invest", "A minus B"]):
    c = ws.cell(row=13, column=2 + j, value=h); c.font = HDR; c.fill = HDRFILL; c.alignment = Alignment(horizontal="center" if j else "left")

aPayN = f'COUNTIF({rng("aEnd")},">0")+1'
bPayN = f'COUNTIF({rng("bEnd")},">0")+1'
rows = [
 ("Total mortgage interest paid (within period)", f"={tot('aInt')}", f"={tot('bInt')}", USD),
 ("Interest saved by prepaying", "=D14-C14", None, USD),
 ("Mortgage payoff date", f'=IF({aPayN}>{NROWS},"Beyond 40 yrs",EDATE(FirstPay,{aPayN}-1))', f'=IF({bPayN}>{NROWS},"Beyond 40 yrs",EDATE(FirstPay,{bPayN}-1))', MON),
 ("Mortgage balance at end of period", f"={at_h('aEnd')}", f"={at_h('bEnd')}", USD),
 ("Extra principal paid", f"={tot('aXL')}", "=0", USD),
 ("Total investment contributions (incl. reinvested tax savings)", f"={tot('aInv')}+{tot('aTR')}", f"={tot('bInv')}+{tot('bTR')}", USD),
 ("Investment gains (after annual tax on yield)", "=C21-C19", "=D21-D19", USD),
 ("Projected investment value (before tax if sold)", f"={at_h('aPt')}", f"={at_h('bPt')}", USD),
 ("Estimated mortgage-interest tax savings", f"={tot('aT')}", f"={tot('bT')}", USD),
 ("Estimated investment taxes (yield tax + tax if sold)", f'={tot("aDT")}+IF(Liquidate="Yes",MAX(0,{at_h("aPt")}-{at_h("aBs")})*InvRate,0)', f'={tot("bDT")}+IF(Liquidate="Yes",MAX(0,{at_h("bPt")}-{at_h("bBs")})*InvRate,0)', USD),
 ("Uninvested cash", f"={at_h('aCB')}", f"={at_h('bCB')}", USD),
 ("After-tax net worth (excluding home)", f"={at_h('aNW')}", f"={at_h('bNW')}", USD),
]
for i, (lab, fa, fb, fmt) in enumerate(rows):
    r = 14 + i
    ws.cell(row=r, column=2, value=lab).font = BOLD if r == 25 else BLACK
    ca = ws.cell(row=r, column=3, value=fa); ca.number_format = fmt
    cb = ws.cell(row=r, column=4, value=fb if fb else "–"); cb.number_format = fmt
    ca.font = cb.font = BOLD if r == 25 else BLACK
    if r == 16:
        d = ws.cell(row=r, column=5, value=f'=IF(OR(ISTEXT(C16),ISTEXT(D16)),"–",ROUND(({bPayN})-({aPayN}),0)&" months sooner")')
    elif fb:
        d = ws.cell(row=r, column=5, value=f"=C{r}-D{r}"); d.number_format = USD
    else:
        d = ws.cell(row=r, column=5, value="–")
    d.font = BOLD if r == 25 else BLACK
    for col in range(2, 6):
        ws.cell(row=r, column=col).border = Border(bottom=thin)
        if col > 2: ws.cell(row=r, column=col).alignment = Alignment(horizontal="right")
for col in range(2, 6):
    ws.cell(row=25, column=col).border = Border(top=Side(style="medium"), bottom=Side(style="medium"))

ws["B27"] = "What prepaying does to interest and tax savings"; ws["B27"].font = H2
tax_rows = [
 ("Interest avoided", "=D14-C14", USD),
 ("Tax savings given up", "=D22-C22", USD),
 ("Net after-tax interest saved", "=C28-C29", USD),
 ("Effective after-tax mortgage rate, year 1", f"=IFERROR(Rate*(1-'Monthly Model'!{C['bT']}13/'Monthly Model'!{C['bYI']}13),Rate)", PCT),
 ("Share of year-1 interest federally deductible", f"=IFERROR(MIN(1,LimFed/'Monthly Model'!{C['bYA']}13),1)", PCT1),
]
for i, (lab, fml, fmt) in enumerate(tax_rows):
    r = 28 + i
    ws.cell(row=r, column=2, value=lab).font = BLACK
    c = ws.cell(row=r, column=3, value=fml); c.number_format = fmt; c.font = BLACK
    for col in (2, 3): ws.cell(row=r, column=col).border = Border(bottom=thin)

ws["B34"] = "Key facts"; ws["B34"].font = H2
facts = [("Scheduled P&I payment", "=Pmt", USD2), ("Extra principal per month", "=Extra", USD), ("Assumed investment return", "=Ret", PCT),
         ("Comparison period (years)", "=Horizon", "0"),
         ("Break-even return (no taxes)", "=(1+Rate/12)^12-1", PCT)]
for i, (lab, fml, fmt) in enumerate(facts):
    r = 35 + i
    ws.cell(row=r, column=2, value=lab).font = BLACK
    c = ws.cell(row=r, column=3, value=fml); c.number_format = fmt; c.font = GREEN
ws["D39"] = "With taxes, find it with Data > What-If Analysis > Goal Seek: set D9 to 0 by changing Inputs return."
ws["D39"].font = MUTED

ws["B41"] = "Not modeled: the planned $30–40K/yr renovations, refinance closing costs, PMI, inflation, market volatility, and the liquidity of invested assets vs. home equity. Illustration only, not a guarantee of future results."
ws["B41"].font = MUTED; ws.merge_cells("B41:E41"); ws["B41"].alignment = Alignment(wrap_text=True); ws.row_dimensions[41].height = 30

# charts (from Annual Schedule)
def line(title, ycols, labels, anchor, ytitle):
    ch = LineChart()
    ch.title = title; ch.height = 8.5; ch.width = 17
    ch.y_axis.number_format = '$#,##0,,"M"' if "Balance" in title or "net worth" in title.lower() else '$#,##0,"K"'
    ch.x_axis.title = "Year"
    for col in ycols:
        ch.add_data(Reference(wsA, min_col=col, min_row=1, max_row=41), titles_from_data=True)
    ch.set_categories(Reference(wsA, min_col=1, min_row=2, max_row=41))
    colors = {4: "2A78D6", 5: "EB6834", 6: "2A78D6", 7: "EB6834", 8: "2A78D6", 9: "EB6834"}
    for s, col in zip(ch.series, ycols):
        s.graphicalProperties.line.solidFill = colors[col]
        s.graphicalProperties.line.width = 22000
        s.smooth = False
        if col in (6, 7): s.graphicalProperties.line.dashStyle = "dash"
    ch.y_axis.delete = False; ch.x_axis.delete = False
    ch.legend.position = "b"
    ws.add_chart(ch, anchor)

line("Balance vs. growth", [4, 5, 6, 7], None, "G2", "$")
line("After-tax net worth", [8, 9], None, "G20", "$")
bc = BarChart(); bc.type = "col"; bc.title = "Annual mortgage interest"; bc.height = 8.5; bc.width = 17
for col in (11, 12):
    bc.add_data(Reference(wsA, min_col=col, min_row=1, max_row=41), titles_from_data=True)
bc.set_categories(Reference(wsA, min_col=1, min_row=2, max_row=41))
bc.series[0].graphicalProperties.solidFill = "2A78D6"; bc.series[1].graphicalProperties.solidFill = "EB6834"
bc.y_axis.number_format = '$#,##0,"K"'; bc.y_axis.delete = False; bc.x_axis.delete = False; bc.legend.position = "b"; bc.gapWidth = 40
ws.add_chart(bc, "G38")
ws["G1"] = "Charts show all 40 years (mortgage solid, investments dashed). The comparison period is set on the Inputs sheet."; ws["G1"].font = MUTED

# ======================= NOTES =======================
wsN = wb.create_sheet("Notes")
wsN.sheet_view.showGridLines = False
wsN.column_dimensions["B"].width = 110
notes = [
 ("From the client's email", True),
 ('"It\'s about $1,070,000 at 7% for a 30 year fixed. I am planning to pay an extra 2-3K a month to pay the mortgage off early." (Nov 4, 2026)', False),
 ('"I\'m also planning to put about 30-40K a year into the home for renovations and improvements." Not part of this comparison, but it competes for the same cash flow.', False),
 ("", False),
 ("How the model works", True),
 ("Same dollars, same dates. Each month both options spend the same total: the regular-loan scheduled payment plus the extra amount. Option A sends the extra to principal; Option B invests it in a taxable account.", False),
 ("When either loan is paid off, the freed-up payment is invested (Inputs: 'Invest full payment once a loan is paid off'), so neither option gets more cash than the other.", False),
 ("Net worth = investment account minus tax owed if sold, plus uninvested cash, minus remaining mortgage. The home is worth the same in both options and is left out.", False),
 ("Mortgage interest deduction: each year's interest × (limit ÷ average balance, max 100%). With 'Compare to standard', only the amount itemizing adds above the standard deduction counts, at the marginal federal and state rates. SALT on the federal return is not modeled; standard deductions are held flat.", False),
 ("Investment taxes: the taxable yield is taxed each year at the federal qualified rate (+ NIIT and state if selected). Remaining gains are taxed at the end if 'as if sold' is Yes. Returns are steady, not variable.", False),
 ("Refinance (optional): from the refinance month the rate changes and the regular loan is re-amortized to the original end date. Closing costs are not included.", False),
 ("", False),
 ("Sheets", True),
 ("Comparison: verdict, side-by-side table, tax effects and charts. Inputs: every assumption (blue = editable, yellow = needed/confirm). Monthly Model: the full month-by-month calculation for both options (480 months). Annual Schedule: year-end values feeding the charts.", False),
 ("All results are live formulas. Change any blue cell on the Inputs sheet and everything recalculates.", False),
 ("", False),
 ("Illustration based on stated assumptions; not a guarantee of future results.", False),
]
for i, (t, b) in enumerate(notes):
    c = wsN.cell(row=2 + i, column=2, value=t)
    c.font = H2 if b else BLACK
    c.alignment = Alignment(wrap_text=True, vertical="top")

ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
wsI.page_setup.fitToWidth = 1; wsI.page_setup.fitToHeight = 0; wsI.sheet_properties.pageSetUpPr.fitToPage = True
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print("saved", OUT)
