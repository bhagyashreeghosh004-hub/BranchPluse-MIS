"""Builds the Excel MIS pack. Every KPI is a LIVE formula (SUMIFS, COUNTIFS, AVERAGEIFS, VLOOKUP, INDEX/MATCH),
so the report keeps working when you paste new data into the Data_ sheets."""
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.properties import PageSetupProperties
from src import config as C

NAVY, TEAL, ORANGE, LIGHT, GREY = "0B2545", "13A89E", "F28F3B", "EAF2F8", "F4F6F8"
F_TITLE = Font(name="Calibri", size=22, bold=True, color="FFFFFF")
F_HDR = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
F_B = Font(name="Calibri", size=11, bold=True)
FILL_HDR = PatternFill("solid", fgColor=NAVY)
FILL_LIGHT = PatternFill("solid", fgColor=LIGHT)
FILL_INPUT = PatternFill("solid", fgColor="FFF4CC")
thin = Side(style="thin", color="D0D7DE")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

INR = '"₹"#,##0'
PCT = "0.0%"


def hdr(ws, row, col, labels, height=32):
    for i, t in enumerate(labels):
        c = ws.cell(row, col + i, t)
        c.font, c.fill, c.alignment, c.border = F_HDR, FILL_HDR, CENTER, BORDER
    ws.row_dimensions[row].height = height


def dump(ws, df, widths, formats=None):
    hdr(ws, 1, 1, list(df.columns))
    for r in df.itertuples(index=False):
        ws.append(list(r))
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[L(i)].width = w
    for col, fmt in (formats or {}).items():
        for cell in ws[col][1:]:
            cell.number_format = fmt
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_properties.tabColor = "8D99AE"


def build(path):
    d = C.CLEAN_DIR
    master = pd.read_csv(d / "branch_master.csv")
    pol = pd.read_csv(d / "policies.csv", parse_dates=["Policy_Date"])
    claims = pd.read_csv(d / "claims.csv", parse_dates=["Claim_Date"])
    ren = pd.read_csv(d / "renewals.csv", parse_dates=["Due_Month"])
    dq = pd.read_csv(C.OUT_DIR / "data_quality_log.csv")
    nb = len(master)
    npol, ncl, nren = len(pol), len(claims), len(ren)

    wb = Workbook()
    ws_dash = wb.active; ws_dash.title = "Dashboard"
    ws_sum = wb.create_sheet("Branch_Summary")
    ws_trend = wb.create_sheet("Monthly_Trend")
    ws_dq = wb.create_sheet("DQ_Report")
    ws_set = wb.create_sheet("Settings")
    ws_m = wb.create_sheet("Branch_Master")
    ws_p = wb.create_sheet("Data_Policies")
    ws_c = wb.create_sheet("Data_Claims")
    ws_r = wb.create_sheet("Data_Renewals")

    # ------------------------------------------------------------------ DATA SHEETS
    dump(ws_m, master, [11, 18, 14, 20, 16], {"E": INR})

    p = pol.copy(); p["Policy_Date"] = p["Policy_Date"].dt.date
    p = p[["Policy_ID", "Policy_Date", "Branch_ID", "LOB", "Premium", "Policy_Type", "Channel"]]
    p.insert(2, "Month", None)
    dump(ws_p, p, [12, 13, 11, 11, 20, 13, 13, 15], {"B": "dd-mmm-yyyy", "C": "mmm-yy", "F": INR})
    ws_p["C1"].value = "Month (formula)"
    for r in range(2, npol + 2):
        ws_p[f"C{r}"] = f"=EOMONTH(B{r},-1)+1"
    # Policies: A id, B date, C month, D branch, E LOB, F premium, G type, H channel

    c = claims.copy(); c["Claim_Date"] = c["Claim_Date"].dt.date
    c = c[["Claim_ID", "Policy_ID", "Claim_Date", "Claim_Amount", "Status", "TAT_Days"]]
    c["TAT_Days"] = c["TAT_Days"].astype("Float64").astype(object).where(c["TAT_Days"].notna(), None)
    c.insert(6, "Branch_ID_Lookup", None)
    dump(ws_c, c, [11, 12, 13, 15, 11, 10, 20], {"C": "dd-mmm-yyyy", "D": INR})
    ws_c["G1"].value = "Branch_ID (VLOOKUP)"
    for r in range(2, ncl + 2):
        ws_c[f"G{r}"] = f'=IFERROR(VLOOKUP(B{r},Data_Policies!$A$2:$D${npol+1},4,FALSE),"NOT FOUND")'

    rr = ren.copy(); rr["Due_Month"] = rr["Due_Month"].dt.date
    dump(ws_r, rr[["Due_ID", "Branch_ID", "Due_Month", "Renewed"]], [11, 11, 12, 10], {"C": "mmm-yy"})

    # ranges
    P = lambda col: f"Data_Policies!${col}$2:${col}${npol+1}"
    CL = lambda col: f"Data_Claims!${col}$2:${col}${ncl+1}"
    RN = lambda col: f"Data_Renewals!${col}$2:${col}${nren+1}"

    # ------------------------------------------------------------------ SETTINGS
    ws_set.sheet_properties.tabColor = ORANGE
    ws_set["A1"] = "Settings: change the yellow cells and the whole report updates"
    ws_set["A1"].font = Font(name="Calibri", size=14, bold=True, color=NAVY)
    items = [
        (3, "Weight: Target achievement", C.WEIGHTS["achievement"], "0%"),
        (4, "Weight: Claim settlement ratio", C.WEIGHTS["settlement"], "0%"),
        (5, "Weight: Renewal rate", C.WEIGHTS["renewal"], "0%"),
        (6, "Weight: Claim TAT", C.WEIGHTS["tat"], "0%"),
        (7, "Weights total (must be 100%)", "=SUM(B3:B6)", "0%"),
        (9, "Green if score is at least", C.GREEN_MIN, "0"),
        (10, "Amber if score is at least", C.AMBER_MIN, "0"),
        (12, "Red flag: achievement below", C.FLAG_ACHIEVEMENT_BELOW, "0%"),
        (13, "Red flag: settlement ratio below", C.FLAG_SETTLEMENT_BELOW, "0%"),
        (14, "Red flag: average TAT above (days)", C.FLAG_TAT_ABOVE, "0"),
        (15, "Red flag: renewal rate below", C.FLAG_RENEWAL_BELOW, "0%"),
        (17, "TAT: full marks up to (days)", C.TAT_GOOD, "0"),
        (18, "TAT: points lost per extra day", C.TAT_PENALTY, "0"),
    ]
    for r, label, val, fmt in items:
        ws_set.cell(r, 1, label)
        cell = ws_set.cell(r, 2, val); cell.number_format = fmt; cell.border = BORDER
        cell.alignment = Alignment(horizontal="center")
        if r != 7: cell.fill = FILL_INPUT
    ws_set["C7"] = '=IF(ROUND(B7,4)=1,"OK","Weights must total 100%")'
    ws_set["A20"] = ("Health Score = 40% x min(achievement,100%) + 25% x settlement ratio + 20% x renewal rate "
                     "+ 15% x TAT score (weights above). TAT score = 100 - (avg TAT - full-marks days) x penalty, kept between 0 and 100.")
    ws_set["A20"].alignment = Alignment(wrap_text=True, vertical="top")
    ws_set.merge_cells("A20:C23")
    ws_set.column_dimensions["A"].width = 42; ws_set.column_dimensions["B"].width = 14; ws_set.column_dimensions["C"].width = 28

    # ------------------------------------------------------------------ BRANCH SUMMARY
    cols = ["Branch ID", "Branch", "Region", "Manager", "Annual Target", "Premium Achieved", "Achievement %",
            "Policies Issued", "Avg Premium / Policy", "Claims Received", "Claims Settled", "Settlement Ratio",
            "Avg Claim TAT (days)", "Renewals Due", "Renewed", "Renewal Rate", "Health Score", "Status", "Red Flags", "Rank"]
    hdr(ws_sum, 1, 1, cols, 44)
    R1, R2 = 2, nb + 1
    for i, bid in enumerate(master["Branch_ID"], start=2):
        ws_sum[f"A{i}"] = bid
        ws_sum[f"B{i}"] = f"=VLOOKUP($A{i},Branch_Master!$A$2:$E${nb+1},2,FALSE)"
        ws_sum[f"C{i}"] = f"=VLOOKUP($A{i},Branch_Master!$A$2:$E${nb+1},3,FALSE)"
        ws_sum[f"D{i}"] = f"=VLOOKUP($A{i},Branch_Master!$A$2:$E${nb+1},4,FALSE)"
        ws_sum[f"E{i}"] = f"=VLOOKUP($A{i},Branch_Master!$A$2:$E${nb+1},5,FALSE)"
        ws_sum[f"F{i}"] = f"=SUMIFS({P('F')},{P('D')},$A{i})"
        ws_sum[f"G{i}"] = f"=IF(E{i}=0,0,F{i}/E{i})"
        ws_sum[f"H{i}"] = f"=COUNTIFS({P('D')},$A{i})"
        ws_sum[f"I{i}"] = f"=IF(H{i}=0,0,F{i}/H{i})"
        ws_sum[f"J{i}"] = f"=COUNTIFS({CL('G')},$A{i})"
        ws_sum[f"K{i}"] = f'=COUNTIFS({CL("G")},$A{i},{CL("E")},"Settled")'
        ws_sum[f"L{i}"] = f"=IF(J{i}=0,0,K{i}/J{i})"
        ws_sum[f"M{i}"] = f'=IFERROR(AVERAGEIFS({CL("F")},{CL("G")},$A{i},{CL("E")},"Settled"),0)'
        ws_sum[f"N{i}"] = f"=COUNTIFS({RN('B')},$A{i})"
        ws_sum[f"O{i}"] = f'=COUNTIFS({RN("B")},$A{i},{RN("D")},"Yes")'
        ws_sum[f"P{i}"] = f"=IF(N{i}=0,0,O{i}/N{i})"
        ws_sum[f"Q{i}"] = (f"=Settings!$B$3*MIN(G{i},1)*100+Settings!$B$4*L{i}*100+Settings!$B$5*P{i}*100"
                           f"+Settings!$B$6*MAX(0,MIN(100,100-(M{i}-Settings!$B$17)*Settings!$B$18))")
        ws_sum[f"R{i}"] = f'=IF(Q{i}>=Settings!$B$9,"Green",IF(Q{i}>=Settings!$B$10,"Amber","Red"))'
        parts = (f'IF(G{i}<Settings!$B$12,"Below target; ","")&IF(L{i}<Settings!$B$13,"Low claim settlement; ","")'
                 f'&IF(M{i}>Settings!$B$14,"Slow claim TAT; ","")&IF(P{i}<Settings!$B$15,"Low renewal rate; ","")')
        ws_sum[f"S{i}"] = (f'=IF(AND(G{i}>=Settings!$B$12,L{i}>=Settings!$B$13,M{i}<=Settings!$B$14,P{i}>=Settings!$B$15),'
                           f'"None",LEFT({parts},LEN({parts})-2))')
        ws_sum[f"T{i}"] = f"=RANK(Q{i},$Q${R1}:$Q${R2},0)"
    t = nb + 2  # totals row
    ws_sum[f"A{t}"] = "TOTAL"; ws_sum[f"B{t}"] = "All branches"
    for col in "EFHJKNO":
        ws_sum[f"{col}{t}"] = f"=SUM({col}{R1}:{col}{R2})"
    ws_sum[f"G{t}"] = f"=F{t}/E{t}"
    ws_sum[f"I{t}"] = f"=F{t}/H{t}"
    ws_sum[f"L{t}"] = f"=K{t}/J{t}"
    ws_sum[f"M{t}"] = f'=IFERROR(AVERAGEIFS({CL("F")},{CL("E")},"Settled"),0)'
    ws_sum[f"P{t}"] = f"=O{t}/N{t}"
    ws_sum[f"Q{t}"] = f"=AVERAGE(Q{R1}:Q{R2})"
    for col in range(1, 21):
        cell = ws_sum.cell(t, col); cell.font = F_B; cell.fill = FILL_LIGHT
    fmts = {"E": INR, "F": INR, "G": PCT, "H": "#,##0", "I": INR, "J": "#,##0", "K": "#,##0", "L": PCT, "M": "0.0",
            "N": "#,##0", "O": "#,##0", "P": PCT, "Q": "0.0", "T": "0"}
    for col, f in fmts.items():
        for r in range(2, t + 1):
            ws_sum[f"{col}{r}"].number_format = f
    for r in range(2, t + 1):
        for col in range(1, 21):
            cell = ws_sum.cell(r, col); cell.border = BORDER
            if col not in (2, 3, 4, 19): cell.alignment = Alignment(horizontal="center")
    for i, w in enumerate([9, 15, 12, 17, 15, 17, 13, 11, 14, 11, 11, 12, 13, 11, 10, 11, 10, 9, 44, 7], 1):
        ws_sum.column_dimensions[L(i)].width = w
    ws_sum.freeze_panes = "C2"
    ws_sum.sheet_properties.tabColor = TEAL
    green = PatternFill("solid", bgColor="C6EFCE", fgColor="C6EFCE")
    amber = PatternFill("solid", bgColor="FFEB9C", fgColor="FFEB9C")
    red = PatternFill("solid", bgColor="FFC7CE", fgColor="FFC7CE")

    def rag_cf(ws, rng):
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Green"'], fill=green, font=Font(name="Calibri", bold=True, color="006100")))
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Amber"'], fill=amber, font=Font(name="Calibri", bold=True, color="9C5700")))
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Red"'], fill=red, font=Font(name="Calibri", bold=True, color="9C0006")))
    rag_cf(ws_sum, f"R{R1}:R{R2}")
    ws_sum.conditional_formatting.add(f"G{R1}:G{R2}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1.3, color=TEAL))
    ws_sum.conditional_formatting.add(f"Q{R1}:Q{R2}", ColorScaleRule(start_type="min", start_color="F8696B", mid_type="percentile", mid_value=50, mid_color="FFEB84", end_type="max", end_color="63BE7B"))
    ws_sum.conditional_formatting.add(f"S{R1}:S{R2}", FormulaRule(formula=[f'$S{R1}<>"None"'], font=Font(name="Calibri", color="9C0006", bold=True)))

    # ------------------------------------------------------------------ MONTHLY TREND
    ws_trend.sheet_properties.tabColor = TEAL
    lobs = ["Motor", "Health", "Travel", "Home", "Personal Accident"]
    hdr(ws_trend, 1, 1, ["Month"] + lobs + ["Total Premium", "Monthly Target", "Achievement %", "Policies Issued"], 36)
    months = pd.date_range(C.FY_START, C.FY_END, freq="MS")
    for i, m in enumerate(months, start=2):
        ws_trend[f"A{i}"] = m.date(); ws_trend[f"A{i}"].number_format = "mmm-yy"
        for j, lob in enumerate(lobs):
            col = L(2 + j)
            ws_trend[f"{col}{i}"] = f"=SUMIFS({P('F')},{P('C')},$A{i},{P('E')},{col}$1)"
            ws_trend[f"{col}{i}"].number_format = INR
        ws_trend[f"G{i}"] = f"=SUM(B{i}:F{i})"
        ws_trend[f"H{i}"] = f"=Branch_Summary!$E${nb+2}/12"
        ws_trend[f"I{i}"] = f"=G{i}/H{i}"
        ws_trend[f"J{i}"] = f"=COUNTIFS({P('C')},$A{i})"
    tr = len(months) + 2
    ws_trend[f"A{tr}"] = "TOTAL"
    for col in "BCDEFGHJ":
        ws_trend[f"{col}{tr}"] = f"=SUM({col}2:{col}{tr-1})"
    ws_trend[f"I{tr}"] = f"=G{tr}/H{tr}"
    for col in range(1, 11):
        ws_trend.cell(tr, col).font = F_B; ws_trend.cell(tr, col).fill = FILL_LIGHT
    for r in range(2, tr + 1):
        for col in "GH": ws_trend[f"{col}{r}"].number_format = INR
        ws_trend[f"I{r}"].number_format = PCT; ws_trend[f"J{r}"].number_format = "#,##0"
        for col in "BCDEF": ws_trend[f"{col}{r}"].number_format = INR
        for col in range(1, 11): ws_trend.cell(r, col).border = BORDER
    for i, w in enumerate([11, 14, 14, 14, 14, 17, 15, 15, 13, 13], 1):
        ws_trend.column_dimensions[L(i)].width = w
    ws_trend.conditional_formatting.add(f"I2:I{tr-1}", ColorScaleRule(start_type="num", start_value=0.7, start_color="F8696B", mid_type="num", mid_value=1, mid_color="FFEB84", end_type="num", end_value=1.3, end_color="63BE7B"))

    # LOB mix + region blocks
    lb = tr + 3
    ws_trend[f"A{lb-1}"] = "Line of business mix"; ws_trend[f"A{lb-1}"].font = Font(name="Calibri", bold=True, color=NAVY, size=12)
    hdr(ws_trend, lb, 1, ["Line of Business", "Premium", "Share %"], 24)
    for k, lob in enumerate(lobs, start=1):
        r = lb + k
        ws_trend[f"A{r}"] = lob
        ws_trend[f"B{r}"] = f"=SUMIFS({P('F')},{P('E')},A{r})"; ws_trend[f"B{r}"].number_format = INR
        ws_trend[f"C{r}"] = f"=B{r}/SUM($B${lb+1}:$B${lb+len(lobs)})"; ws_trend[f"C{r}"].number_format = PCT
        for col in "ABC": ws_trend[f"{col}{r}"].border = BORDER
    LOB_FIRST, LOB_LAST = lb + 1, lb + len(lobs)
    rb = LOB_LAST + 3
    ws_trend[f"A{rb-1}"] = "Region performance"; ws_trend[f"A{rb-1}"].font = Font(name="Calibri", bold=True, color=NAVY, size=12)
    hdr(ws_trend, rb, 1, ["Region", "Premium", "Target", "Achievement %"], 24)
    regions = list(dict.fromkeys(master["Region"]))
    for k, reg in enumerate(regions, start=1):
        r = rb + k
        ws_trend[f"A{r}"] = reg
        ws_trend[f"B{r}"] = f"=SUMIFS(Branch_Summary!$F${R1}:$F${R2},Branch_Summary!$C${R1}:$C${R2},A{r})"
        ws_trend[f"C{r}"] = f"=SUMIFS(Branch_Summary!$E${R1}:$E${R2},Branch_Summary!$C${R1}:$C${R2},A{r})"
        ws_trend[f"D{r}"] = f"=B{r}/C{r}"
        ws_trend[f"B{r}"].number_format = INR; ws_trend[f"C{r}"].number_format = INR; ws_trend[f"D{r}"].number_format = PCT
        for col in "ABCD": ws_trend[f"{col}{r}"].border = BORDER

    # ------------------------------------------------------------------ DQ REPORT
    ws_dq.sheet_properties.tabColor = ORANGE
    ws_dq["A1"] = "Data Quality & Reconciliation Report"; ws_dq["A1"].font = Font(name="Calibri", size=16, bold=True, color=NAVY)
    ws_dq["A3"] = "1. Reconciliation checks (live formulas)"; ws_dq["A3"].font = Font(name="Calibri", bold=True, size=12, color=NAVY)
    hdr(ws_dq, 4, 1, ["Check", "Source value", "Report value", "Difference", "Result"], 24)
    checks = [
        ("Total premium: data sheet vs Branch_Summary", f"=SUM({P('F')})", f"=Branch_Summary!F{t}"),
        ("Policy count: data sheet vs Branch_Summary", f"=COUNTA({P('A')})", f"=Branch_Summary!H{t}"),
        ("Claim count: data sheet vs Branch_Summary", f"=COUNTA({CL('A')})", f"=Branch_Summary!J{t}"),
        ("Renewals due: data sheet vs Branch_Summary", f"=COUNTA({RN('A')})", f"=Branch_Summary!N{t}"),
        ("Total premium: data sheet vs Monthly_Trend", f"=SUM({P('F')})", f"=Monthly_Trend!G{tr}"),
        ("Claims that failed the VLOOKUP (should be 0)", f'=COUNTIF({CL("G")},"NOT FOUND")', "=0"),
    ]
    for k, (name, a, b) in enumerate(checks, start=5):
        ws_dq[f"A{k}"] = name; ws_dq[f"B{k}"] = a; ws_dq[f"C{k}"] = b
        ws_dq[f"D{k}"] = f"=ROUND(B{k}-C{k},2)"
        ws_dq[f"E{k}"] = f'=IF(D{k}=0,"PASS","FAIL")'
        for col in "ABCDE": ws_dq[f"{col}{k}"].border = BORDER
        ws_dq[f"B{k}"].number_format = "#,##0"; ws_dq[f"C{k}"].number_format = "#,##0"
    ws_dq.conditional_formatting.add(f"E5:E{4+len(checks)}", CellIsRule(operator="equal", formula=['"PASS"'], fill=green, font=Font(name="Calibri", bold=True, color="006100")))
    ws_dq.conditional_formatting.add(f"E5:E{4+len(checks)}", CellIsRule(operator="equal", formula=['"FAIL"'], fill=red, font=Font(name="Calibri", bold=True, color="9C0006")))
    q0 = 5 + len(checks) + 2
    ws_dq[f"A{q0}"] = "2. Issues found and fixed before reporting (from the Python Data Quality Gate)"
    ws_dq[f"A{q0}"].font = Font(name="Calibri", bold=True, size=12, color=NAVY)
    hdr(ws_dq, q0 + 1, 1, list(dq.columns), 24)
    for k, row in enumerate(dq.itertuples(index=False), start=q0 + 2):
        for j, v in enumerate(row, start=1):
            cell = ws_dq.cell(k, j, v); cell.border = BORDER; cell.alignment = Alignment(wrap_text=True, vertical="top")
    for col, w in zip("ABCDE", [58, 44, 16, 52, 10]):
        ws_dq.column_dimensions[col].width = w

    # ------------------------------------------------------------------ DASHBOARD
    ws = ws_dash
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = NAVY
    for col in range(1, 15):
        ws.column_dimensions[L(col)].width = 13.5
    ws.column_dimensions["A"].width = 2; ws.column_dimensions["N"].width = 2
    ws.merge_cells("B1:M2")
    ws["B1"] = f"BranchPulse | Pune Branch MIS Dashboard | {C.FY_LABEL}"
    ws["B1"].font = F_TITLE; ws["B1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for r in (1, 2):
        for col in range(2, 14): ws.cell(r, col).fill = FILL_HDR
    ws.merge_cells("B3:M3")
    ws["B3"] = "Synthetic demo data. All figures are live Excel formulas driven by the Data_ sheets and the Settings sheet."
    ws["B3"].font = Font(name="Calibri", italic=True, color="5A6B7B", size=10); ws["B3"].alignment = Alignment(indent=1)

    cards = [
        ("TOTAL PREMIUM", f"=Branch_Summary!F{t}/10000000", '"₹"0.00" Cr"'),
        ("TARGET ACHIEVEMENT", f"=Branch_Summary!G{t}", PCT),
        ("POLICIES ISSUED", f"=Branch_Summary!H{t}", "#,##0"),
        ("CLAIM SETTLEMENT", f"=Branch_Summary!L{t}", PCT),
        ("AVG CLAIM TAT", f"=Branch_Summary!M{t}", '0.0" days"'),
        ("RENEWAL RATE", f"=Branch_Summary!P{t}", PCT),
    ]
    for k, (lab, f, fmt) in enumerate(cards):
        c0 = 2 + k * 2
        ws.merge_cells(start_row=5, start_column=c0, end_row=5, end_column=c0 + 1)
        ws.merge_cells(start_row=6, start_column=c0, end_row=7, end_column=c0 + 1)
        a, b = ws.cell(5, c0, lab), ws.cell(6, c0, f)
        a.font = Font(name="Calibri", size=9, bold=True, color="FFFFFF"); a.alignment = CENTER
        b.font = Font(name="Calibri", size=20, bold=True, color=NAVY); b.alignment = CENTER; b.number_format = fmt
        for rr_ in (5, 6, 7):
            for cc in (c0, c0 + 1):
                ws.cell(rr_, cc).fill = PatternFill("solid", fgColor=TEAL if rr_ == 5 else LIGHT)
    ws.row_dimensions[5].height = 20
    ws.merge_cells("B9:E9"); ws["B9"] = "Branch status"
    ws["B9"].font = Font(name="Calibri", bold=True, color=NAVY, size=12)
    sm = f"Branch_Summary!$R${R1}:$R${R2}"
    for k, (lab, key) in enumerate([("Green", "Green"), ("Amber", "Amber"), ("Red", "Red")]):
        c0 = 2 + k * 2
        ws.merge_cells(start_row=10, start_column=c0, end_row=10, end_column=c0 + 1)
        cell = ws.cell(10, c0, f'="{lab}: "&COUNTIF({sm},"{key}")&" branches"')
        cell.alignment = CENTER; cell.font = F_B
        fill = {"Green": "C6EFCE", "Amber": "FFEB9C", "Red": "FFC7CE"}[key]
        for cc in (c0, c0 + 1): ws.cell(10, cc).fill = PatternFill("solid", fgColor=fill)
    ws.merge_cells("H10:M10")
    ws["H10"] = f'="Network health score: "&TEXT(Branch_Summary!Q{t},"0.0")&" / 100"'
    ws["H10"].font = F_B; ws["H10"].alignment = CENTER
    for cc in range(8, 14): ws.cell(10, cc).fill = FILL_LIGHT
    ws.row_dimensions[10].height = 24

    def style_chart(ch, title, w=15.4, h=7.6):
        ch.title = title; ch.width, ch.height = w, h
        ch.legend.position = "b"

    # chart 1: premium vs target
    c1 = BarChart(); c1.type = "col"; c1.grouping = "clustered"
    c1.add_data(Reference(ws_sum, min_col=5, min_row=1, max_row=R2), titles_from_data=True)
    c1.add_data(Reference(ws_sum, min_col=6, min_row=1, max_row=R2), titles_from_data=True)
    c1.set_categories(Reference(ws_sum, min_col=2, min_row=2, max_row=R2))
    c1.series[0].graphicalProperties.solidFill = "B8C4D0"; c1.series[1].graphicalProperties.solidFill = TEAL
    style_chart(c1, "Premium achieved vs annual target (₹)"); c1.y_axis.numFmt = '0,,"M"'
    c1.y_axis.delete = False; c1.x_axis.delete = False
    ws.add_chart(c1, "B12")
    # chart 2: monthly premium trend
    c2 = LineChart()
    c2.add_data(Reference(ws_trend, min_col=7, min_row=1, max_row=tr - 1), titles_from_data=True)
    c2.add_data(Reference(ws_trend, min_col=8, min_row=1, max_row=tr - 1), titles_from_data=True)
    c2.set_categories(Reference(ws_trend, min_col=1, min_row=2, max_row=tr - 1))
    c2.series[0].graphicalProperties.line.solidFill = TEAL; c2.series[0].graphicalProperties.line.width = 32000
    c2.series[1].graphicalProperties.line.solidFill = ORANGE; c2.series[1].graphicalProperties.line.dashStyle = "dash"
    c2.series[0].smooth = False; c2.series[1].smooth = False
    style_chart(c2, "Monthly premium vs monthly target"); c2.y_axis.numFmt = '0.0,,"M"'; c2.x_axis.number_format = "mmm-yy"
    c2.y_axis.delete = False; c2.x_axis.delete = False
    ws.add_chart(c2, "H12")
    # chart 3: LOB mix
    c3 = PieChart()
    c3.add_data(Reference(ws_trend, min_col=2, min_row=lb, max_row=LOB_LAST), titles_from_data=True)
    c3.set_categories(Reference(ws_trend, min_col=1, min_row=LOB_FIRST, max_row=LOB_LAST))
    c3.dataLabels = DataLabelList(); c3.dataLabels.showPercent = True
    c3.dataLabels.showVal = False; c3.dataLabels.showCatName = False; c3.dataLabels.showSerName = False; c3.dataLabels.showLeaderLines = False
    for idx, colr in enumerate([NAVY, TEAL, ORANGE, "8D99AE", "E4572E"]):
        pt = DataPoint(idx=idx); pt.graphicalProperties.solidFill = colr; c3.series[0].dPt.append(pt)
    style_chart(c3, "Premium mix by line of business")
    ws.add_chart(c3, "B28")
    # chart 4: health score
    c4 = BarChart(); c4.type = "bar"
    c4.add_data(Reference(ws_sum, min_col=17, min_row=1, max_row=R2), titles_from_data=True)
    c4.set_categories(Reference(ws_sum, min_col=2, min_row=2, max_row=R2))
    c4.series[0].graphicalProperties.solidFill = NAVY
    style_chart(c4, "Branch Health Score (0 to 100)", h=8.6); c4.legend = None
    c4.y_axis.scaling.min = 0; c4.y_axis.scaling.max = 100
    c4.y_axis.delete = False; c4.x_axis.delete = False; c4.x_axis.scaling.orientation = "maxMin"
    ws.add_chart(c4, "H28")

    # Leaderboard + red-flag radar
    s0 = 44
    ws.merge_cells(f"B{s0}:M{s0}")
    ws[f"B{s0}"] = "Branch leaderboard & Red-Flag Radar (auto-sorted by Health Score)"
    ws[f"B{s0}"].font = Font(name="Calibri", bold=True, color=NAVY, size=13)
    heads = [("B", "B"), ("C", "D"), ("E", "E"), ("F", "F"), ("G", "G"), ("H", "M")]
    labels = ["Rank", "Branch", "Region", "Score", "Status", "Red Flags"]
    for (a, b), lab in zip(heads, labels):
        if a != b: ws.merge_cells(f"{a}{s0+1}:{b}{s0+1}")
        c_ = ws[f"{a}{s0+1}"]; c_.value = lab; c_.font, c_.fill, c_.alignment = F_HDR, FILL_HDR, CENTER
        for cc in range(ord(a), ord(b) + 1): ws[f"{chr(cc)}{s0+1}"].fill = FILL_HDR
    QS, BS, CS_, FS, RS = (f"Branch_Summary!${x}${R1}:${x}${R2}" for x in "QBCSR")
    for k in range(1, nb + 1):
        r = s0 + 1 + k
        ws[f"B{r}"] = k
        ws[f"F{r}"] = f"=LARGE({QS},B{r})"; ws[f"F{r}"].number_format = "0.0"
        ws[f"C{r}"] = f"=INDEX({BS},MATCH(F{r},{QS},0))"
        ws[f"E{r}"] = f"=INDEX({CS_},MATCH(F{r},{QS},0))"
        ws[f"G{r}"] = f"=INDEX({RS},MATCH(F{r},{QS},0))"
        ws[f"H{r}"] = f"=INDEX({FS},MATCH(F{r},{QS},0))"
        ws.merge_cells(f"C{r}:D{r}"); ws.merge_cells(f"H{r}:M{r}")
        for col in "BCDEFGHIJKLM":
            cell = ws[f"{col}{r}"]; cell.border = BORDER
            cell.alignment = Alignment(horizontal="left" if col in "CH" else "center", vertical="center")
        ws.row_dimensions[r].height = 19
    rag_cf(ws, f"G{s0+2}:G{s0+1+nb}")
    ws.conditional_formatting.add(f"H{s0+2}:H{s0+1+nb}", FormulaRule(formula=[f'$H{s0+2}<>"None"'], font=Font(name="Calibri", color="9C0006", bold=True)))
    ws.conditional_formatting.add(f"F{s0+2}:F{s0+1+nb}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=100, color=TEAL))
    ws.page_setup.orientation = "portrait"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_area = f"A1:N{s0+nb+3}"

    wb.calculation.fullCalcOnLoad = True
    wb.save(path)
    print(f"[workbook] saved {path}")


if __name__ == "__main__":
    build(C.OUT_DIR / "BranchPulse_MIS_Report.xlsx")
