"""Build the Excel sales performance tracker from the monthly input CSVs.

Usage:
    python build_workbook.py <data_dir> <output.xlsx>

Every number on the Dashboard, Targets and Checks tabs is an Excel formula that reads the
input tables, so the file keeps working when the next month is pasted in.
"""
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.text import RichText, Text
from openpyxl.chart.title import Title
from openpyxl.comments import Comment
from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties, RegularTextRun
from openpyxl.drawing.text import Font as DrawingFont
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.pagebreak import Break, RowBreak
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.table import Table, TableStyleInfo

FONT = "Arial"
NAVY = "1F3864"
MID = "2E75B6"
LIGHT = "F2F4F7"
GRID = "D9DEE7"
WHITE = "FFFFFF"
NOTE = "595959"

MARKETS = ["United Kingdom", "EU", "Rest of world"]
PRODUCT_LINES = [
    "Christmas & seasonal", "Gifts & accessories", "Candles & lighting", "Party & gift wrap",
    "Kitchen & dining", "Bags & storage", "Toys, crafts & stationery", "Home decor", "Other",
]
CUSTOMER_TYPES = ["New", "Returning", "Guest (no ID)"]
FY_START = pd.Timestamp("2010-12-01")
DEFAULT_MONTH = pd.Timestamp("2011-11-01")
GROWTH = 0.10

GBP = '"£"#,##0;[Red]("£"#,##0);"-"'
GBP2 = '"£"#,##0.00;[Red]("£"#,##0.00);"-"'
PCT = '0.0%;[Red]-0.0%;"-"'
INT = '#,##0;[Red]-#,##0;"-"'
MONTH_FMT = "mmm yyyy"

thin = Side(style="thin", color=GRID)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)


def font(size=10, bold=False, color="000000", italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def fill(color):
    return PatternFill("solid", start_color=color, end_color=color)


def note(ws, ref, text):
    ws[ref] = text
    ws[ref].font = font(9, italic=True, color=NOTE)


def header_row(ws, row, col, labels, color=NAVY, height=None):
    for i, text in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=text)
        c.font = font(9, True, WHITE)
        c.fill = fill(color)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
    if height:
        ws.row_dimensions[row].height = height


def style_body(ws, rows, cols, fmt_by_col, total_row=None):
    for r in rows:
        for c in cols:
            cell = ws.cell(row=r, column=c)
            cell.font = font(10, bold=(r == total_row))
            cell.border = BOX
            if c in fmt_by_col:
                cell.number_format = fmt_by_col[c]
            if r == total_row:
                cell.fill = fill(LIGHT)


def add_achievement_colors(ws, cell_range):
    ws.conditional_formatting.add(cell_range, CellIsRule(
        operator="greaterThanOrEqual", formula=["1"], fill=fill("C6EFCE"),
        font=Font(name=FONT, color="006100", bold=True)))
    ws.conditional_formatting.add(cell_range, CellIsRule(
        operator="between", formula=["0.95", "0.99999"], fill=fill("FFEB9C"),
        font=Font(name=FONT, color="9C5700", bold=True)))
    ws.conditional_formatting.add(cell_range, CellIsRule(
        operator="lessThan", formula=["0.95"], fill=fill("FFC7CE"),
        font=Font(name=FONT, color="9C0006", bold=True)))


def text_props(size=900, bold=False, color="404040"):
    cp = CharacterProperties(sz=size, b=bold, solidFill=color, latin=DrawingFont(typeface=FONT))
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])


def chart_title(text, size=1100):
    cp = CharacterProperties(sz=size, b=True, solidFill=NAVY, latin=DrawingFont(typeface=FONT))
    para = Paragraph(pPr=ParagraphProperties(defRPr=cp), r=[RegularTextRun(rPr=cp, t=text)])
    return Title(tx=Text(rich=RichText(p=[para])), overlay=False)


def tidy_chart(chart):
    chart.roundedCorners = False
    for axis in (chart.x_axis, chart.y_axis):
        axis.txPr = text_props(900)
    if chart.legend is not None:
        chart.legend.txPr = text_props(900)
    if chart.title is not None:
        chart.title.overlay = False
    if chart.legend is not None:
        chart.legend.overlay = False
    for axis in (chart.x_axis, chart.y_axis):
        axis.delete = False
        if axis.title is not None:
            axis.title.overlay = False


def value_labels(fmt):
    labels = DataLabelList()
    labels.showVal = True
    labels.showSerName = False
    labels.showCatName = False
    labels.showLegendKey = False
    labels.showPercent = False
    labels.numFmt = fmt
    labels.dLblPos = "outEnd"
    labels.txPr = text_props(900, True, NAVY)
    return labels


def build(data_dir, out_path):
    data = Path(data_dir)
    sales = pd.read_csv(data / "sales_input.csv", parse_dates=["Month"])
    cust = pd.read_csv(data / "customers_input.csv", parse_dates=["Month"])
    control = pd.read_csv(data / "control_totals.csv", parse_dates=["Month"])
    cleaning = pd.read_csv(data / "cleaning_log.csv").fillna("")
    country_map = pd.read_csv(data / "country_market.csv")
    rules = pd.read_csv(data / "product_rules.csv")

    wb = Workbook()
    dash = wb.active
    dash.title = "Dashboard"
    guide = wb.create_sheet("How to update")
    s_in = wb.create_sheet("Sales_Input")
    c_in = wb.create_sheet("Customers_Input")
    tgt = wb.create_sheet("Targets")
    chk = wb.create_sheet("Checks")
    st = wb.create_sheet("Settings")
    mp = wb.create_sheet("Mapping")
    dash.sheet_properties.tabColor = NAVY
    for ws in (s_in, c_in, tgt):
        ws.sheet_properties.tabColor = "70AD47"

    # ---------------- Settings ----------------
    st["A1"] = "Settings"
    st["A1"].font = font(14, True, NAVY)
    rows = [
        (3, "Target growth vs same month last year", GROWTH, "0.0%",
         "Input. An assumption to set the targets, not a figure from the data. Change it to test other targets."),
        (4, "First month of the fiscal year", FY_START.to_pydatetime(), MONTH_FMT,
         "Input. The fiscal year runs Dec - Nov. Change it to the new first month when a new year starts."),
    ]
    for r, label, value, fmt, comment in rows:
        st.cell(row=r, column=1, value=label).font = font(10)
        c = st.cell(row=r, column=2, value=value)
        c.number_format = fmt
        c.font = font(10, True, "0000FF")
        c.fill = fill("FFFF00")
        c.border = BOX
        note(st, f"C{r}", comment)
    st["A5"] = "Currency"
    st["A5"].font = font(10)
    st["B5"] = "GBP (£)"
    st["B5"].font = font(10)
    st["A7"] = "Lists used by the month drop-down, the dashboard tables and the checks"
    st["A7"].font = font(10, True, NAVY)
    for ref, text in (("B8", "Months in the fiscal year"), ("D8", "Markets"), ("F8", "Product lines"),
                      ("H8", "Customer types")):
        st[ref] = text
        st[ref].font = font(9, True, WHITE)
        st[ref].fill = fill(NAVY)
    for k in range(12):
        c = st.cell(row=9 + k, column=2, value=f"=EDATE($B$4,{k})")
        c.number_format = MONTH_FMT
        c.font = font(10)
    for i, m in enumerate(MARKETS):
        st.cell(row=9 + i, column=4, value=m).font = font(10)
    for i, p in enumerate(PRODUCT_LINES):
        st.cell(row=9 + i, column=6, value=p).font = font(10)
    for i, t in enumerate(CUSTOMER_TYPES):
        st.cell(row=9 + i, column=8, value=t).font = font(10)
    note(st, "A23", "Legend: yellow cells with blue text are inputs you can change. Black text is formulas. "
                    "Blue text on Checks = control totals from the source data.")
    for col, w in zip("ABCDEFGH", (40, 26, 4, 18, 4, 26, 4, 18)):
        st.column_dimensions[col].width = w

    # ---------------- Sales_Input ----------------
    s_in.append(["Month", "Market", "Product line", "Customer type", "Net sales (£)", "Units"])
    for row in sales.itertuples(index=False):
        s_in.append([row.Month.to_pydatetime(), row.Market, row.ProductLine, row.CustomerType,
                     float(row.NetSales), int(row.Units)])
    n_s = len(sales) + 1
    tab = Table(displayName="tblSales", ref=f"A1:F{n_s}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    s_in.add_table(tab)
    for r in range(1, n_s + 1):
        for c in range(1, 7):
            s_in.cell(row=r, column=c).font = font(10, bold=(r == 1), color=WHITE if r == 1 else "000000")
        if r > 1:
            s_in.cell(row=r, column=1).number_format = MONTH_FMT
            s_in.cell(row=r, column=5).number_format = GBP2
            s_in.cell(row=r, column=6).number_format = INT
    for col, w in zip("ABCDEF", (12, 18, 26, 16, 16, 12)):
        s_in.column_dimensions[col].width = w
    s_in.freeze_panes = "A2"
    note(s_in, "H1", "Input table: one row per month, market, product line and customer type.")
    note(s_in, "H2", "To add a month, paste its rows right under the last row (the table grows by itself).")
    note(s_in, "H3", "Dec 2009 - Nov 2010 rows are last year's actuals; the targets are built from them.")
    note(s_in, "H4", "Example row: Nov 2011 | United Kingdom | Home decor | Returning | 112,345.67 | 45,210")
    s_in["A1"].comment = Comment("First day of the month (1 Nov 2011 = November 2011).", "Tracker")

    # ---------------- Customers_Input ----------------
    c_in.append(["Month", "Market", "Customer type", "Customers", "Orders"])
    for row in cust.itertuples(index=False):
        c_in.append([row.Month.to_pydatetime(), row.Market, row.CustomerType, int(row.Customers), int(row.Orders)])
    n_c = len(cust) + 1
    tab2 = Table(displayName="tblCustomers", ref=f"A1:E{n_c}")
    tab2.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    c_in.add_table(tab2)
    for r in range(1, n_c + 1):
        for c in range(1, 6):
            c_in.cell(row=r, column=c).font = font(10, bold=(r == 1), color=WHITE if r == 1 else "000000")
        if r > 1:
            c_in.cell(row=r, column=1).number_format = MONTH_FMT
            c_in.cell(row=r, column=4).number_format = INT
            c_in.cell(row=r, column=5).number_format = INT
    for col, w in zip("ABCDE", (12, 18, 16, 12, 12)):
        c_in.column_dimensions[col].width = w
    c_in.freeze_panes = "A2"
    note(c_in, "G1", "Customers = distinct customers who bought that month. Guests have no customer ID, so they count as 0.")
    note(c_in, "G2", "Orders = distinct invoices that month (returns not counted).")

    # ---------------- Targets ----------------
    tgt.append(["Month", "Market", "Product line", "Same month last year (£)", "Target (£)"])
    r = 2
    for k in range(12):
        for m in MARKETS:
            for p in PRODUCT_LINES:
                tgt.cell(row=r, column=1, value=f"=EDATE(Settings!$B$4,{k})").number_format = MONTH_FMT
                tgt.cell(row=r, column=2, value=m)
                tgt.cell(row=r, column=3, value=p)
                tgt.cell(row=r, column=4, value=(
                    f"=SUMIFS(Sales_Input!$E:$E,Sales_Input!$A:$A,EDATE($A{r},-12),"
                    f"Sales_Input!$B:$B,$B{r},Sales_Input!$C:$C,$C{r})")).number_format = GBP
                tgt.cell(row=r, column=5, value=f"=ROUND(D{r}*(1+Settings!$B$3),0)").number_format = GBP
                r += 1
    n_t = r - 1
    tab3 = Table(displayName="tblTargets", ref=f"A1:E{n_t}")
    tab3.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    tgt.add_table(tab3)
    for rr in range(1, n_t + 1):
        for c in range(1, 6):
            tgt.cell(row=rr, column=c).font = font(10, bold=(rr == 1), color=WHITE if rr == 1 else "000000")
    for col, w in zip("ABCDE", (12, 18, 26, 24, 16)):
        tgt.column_dimensions[col].width = w
    tgt.freeze_panes = "A2"
    note(tgt, "G1", "Target = same month last year x (1 + growth in Settings B3), rounded to the pound.")
    note(tgt, "G2", "Have your own targets? Type them over column E. The dashboard updates by itself.")

    # ---------------- Mapping ----------------
    mp["A1"] = "How countries become markets"
    mp["A1"].font = font(12, True, NAVY)
    header_row(mp, 3, 1, ["Country (as in the source)", "Market"])
    for i, row in enumerate(country_map.itertuples(index=False)):
        mp.cell(row=4 + i, column=1, value=row.Country).font = font(10)
        mp.cell(row=4 + i, column=2, value=row.Market).font = font(10)
    mp["D1"] = "How products become product lines (the first rule that matches the description wins)"
    mp["D1"].font = font(12, True, NAVY)
    header_row(mp, 3, 4, ["Order", "Product line", "Keywords in the description (regular expression)"])
    for i, row in enumerate(rules.itertuples(index=False)):
        mp.cell(row=4 + i, column=4, value=int(row.Order)).font = font(10)
        mp.cell(row=4 + i, column=5, value=row.ProductLine).font = font(10)
        cell = mp.cell(row=4 + i, column=6, value=row.Keywords)
        cell.font = font(9)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    for col, w in zip("ABCDEF", (26, 18, 3, 7, 26, 100)):
        mp.column_dimensions[col].width = w

    # ---------------- Checks ----------------
    chk["A1"] = "Checks"
    chk["A1"].font = font(14, True, NAVY)
    note(chk, "A2", "Control totals (blue) come from prepare_data.py and the cleaned transactions. "
                    "The other columns re-add the input tables with formulas.")
    header_row(chk, 4, 1, ["Check", "Result", "Status"])
    fy_sum = (f'SUMIFS(Sales_Input!$E:$E,Sales_Input!$A:$A,">="&Settings!$B$4,'
              f'Sales_Input!$A:$A,"<="&Settings!$B$20')
    checks = [
        ("Months where the input tables do not match the control totals",
         '=COUNTIF($E$14:$E$200,"CHECK")+COUNTIF($H$14:$H$200,"CHECK")', INT),
        ("Sales_Input rows with an unknown market",
         '=SUMPRODUCT((Sales_Input!$B$2:$B$6000<>"")*(COUNTIF(Settings!$D$9:$D$11,Sales_Input!$B$2:$B$6000)=0))', INT),
        ("Sales_Input rows with an unknown product line",
         '=SUMPRODUCT((Sales_Input!$C$2:$C$6000<>"")*(COUNTIF(Settings!$F$9:$F$17,Sales_Input!$C$2:$C$6000)=0))', INT),
        ("Sales_Input rows with an unknown customer type",
         '=SUMPRODUCT((Sales_Input!$D$2:$D$6000<>"")*(COUNTIF(Settings!$H$9:$H$11,Sales_Input!$D$2:$D$6000)=0))', INT),
        (f"Targets filled (12 months x 3 markets x 9 product lines = {n_t - 1})",
         f"=COUNT(Targets!$E$2:$E${n_t})", INT),
        ("Share of fiscal-year net sales in product line 'Other' (keep it under 15%)",
         f'=IF({fy_sum})=0,0,{fy_sum},Sales_Input!$C:$C,"Other")/{fy_sum}))', PCT),
    ]
    status = ['=IF(B5=0,"OK","CHECK")', '=IF(B6=0,"OK","CHECK")', '=IF(B7=0,"OK","CHECK")',
              '=IF(B8=0,"OK","CHECK")', f'=IF(B9={n_t - 1},"OK","CHECK")', '=IF(B10<0.15,"OK","REVIEW KEYWORDS")']
    for i, ((label, formula, fmt), stat) in enumerate(zip(checks, status)):
        rr = 5 + i
        chk.cell(row=rr, column=1, value=label).font = font(10)
        c = chk.cell(row=rr, column=2, value=formula)
        c.font = font(10)
        c.number_format = fmt
        chk.cell(row=rr, column=3, value=stat).font = font(10, True)
        for col in range(1, 4):
            chk.cell(row=rr, column=col).border = BOX
    chk["A11"] = "All checks"
    chk["A11"].font = font(10, True)
    chk["C11"] = '=IF(COUNTIF($C$5:$C$10,"OK")=6,"All checks OK","Review the checks")'
    chk["C11"].font = font(10, True)
    for ref in ("C5:C10", "C11"):
        chk.conditional_formatting.add(ref, CellIsRule(operator="equal", formula=['"OK"'], font=Font(name=FONT, color="006100", bold=True)))
    chk.conditional_formatting.add("C5:C11", CellIsRule(
        operator="notEqual", formula=['"OK"'], font=Font(name=FONT, color="9C0006", bold=True)))
    chk.conditional_formatting.add("C11", CellIsRule(
        operator="equal", formula=['"All checks OK"'], font=Font(name=FONT, color="006100", bold=True)))
    chk["A13"] = "Monthly control totals"
    chk["A13"].font = font(10, True, NAVY)
    header_row(chk, 13, 2, ["Month", "Net sales in source (£)", "Net sales in Sales_Input (£)", "Status",
                            "Units in source", "Units in Sales_Input", "Status"], height=30)
    for i, row in enumerate(control.itertuples(index=False)):
        rr = 14 + i
        chk.cell(row=rr, column=2, value=row.Month.to_pydatetime()).number_format = MONTH_FMT
        chk.cell(row=rr, column=3, value=float(row.NetSales)).number_format = GBP2
        chk.cell(row=rr, column=4, value=f"=SUMIFS(Sales_Input!$E:$E,Sales_Input!$A:$A,$B{rr})").number_format = GBP2
        chk.cell(row=rr, column=5, value=f'=IF(ABS(D{rr}-C{rr})<1,"OK","CHECK")')
        chk.cell(row=rr, column=6, value=int(row.Units)).number_format = INT
        chk.cell(row=rr, column=7, value=f"=SUMIFS(Sales_Input!$F:$F,Sales_Input!$A:$A,$B{rr})").number_format = INT
        chk.cell(row=rr, column=8, value=f'=IF(G{rr}=F{rr},"OK","CHECK")')
        for col in range(2, 9):
            cell = chk.cell(row=rr, column=col)
            cell.font = font(10, color="0000FF" if col in (3, 6) else "000000")
            cell.border = BOX
    chk["C14"].comment = Comment("Blue = control values written by prepare_data.py from the cleaned source data.", "Tracker")
    last_ctrl = 13 + len(control)
    note(chk, f"B{last_ctrl + 2}", "Tolerance: £1 for net sales (rounding), exact for units. "
                                   "When you add a month, add its control totals here too.")
    chk.column_dimensions["A"].width = 64
    for col, w in zip("BCDEFGH", (12, 20, 22, 10, 14, 16, 10)):
        chk.column_dimensions[col].width = w

    # ---------------- How to update ----------------
    g = guide
    lines = [
        ("Monthly update (about 5 minutes)", True),
        ("1. Export last month's transactions from the sales system: invoice, product, quantity, unit price, customer, country, date.", False),
        ("2. Run  python prepare_data.py <export file> <folder>. It cleans the lines and writes the month's rows and control totals.", False),
        ("3. Paste the new rows right under the last row of Sales_Input and Customers_Input (the tables grow by themselves), "
         "and the control totals under the last row of Checks.", False),
        ("4. Pick the month in Dashboard cell C4. KPIs, tables, charts and targets follow it, with no formula changes.", False),
        ("5. Open Checks: it should say 'All checks OK'.", False),
        ("New fiscal year: change Settings B4 to the first month of the new year. The targets recalculate from last year's actuals.", False),
        ("Your own targets: type them over Targets column E.", False),
        ("Legend: yellow cells with blue text are inputs. Black text is formulas. Green tabs hold the input tables.", False),
        ("", False),
        ("Definitions", True),
        ("Net sales: quantity x unit price, minus returns (invoices starting with C). Postage, fees, manual adjustments "
         "and gift vouchers are excluded.", False),
        ("Target: same month last year x (1 + growth in Settings B3). The 10% growth is an assumption to show the "
         "tracker at work, not a figure from the data.", False),
        ("Achievement: net sales / target. Green is 100% or more, amber 95% to 99.9%, red below 95%. "
         "Gap: net sales - target.", False),
        ("New customer: their first purchase in the data is in that month. Returning: they bought in an earlier month. "
         "Guest: order with no customer ID.", False),
        ("Note: the data starts in Dec 2009, so a customer who last bought before then counts as new when they come back.", False),
        ("Active customers: distinct customers who bought that month. Net sales per order: net sales / orders.", False),
        ("Market: United Kingdom, EU (member states in 2010-2011) or Rest of world. Product line: keyword rules on "
         "the product description (see Mapping).", False),
        ("", False),
        ("Data source", True),
        ("Online Retail II, by Daqing Chen. UCI Machine Learning Repository, 2019. DOI 10.24432/C5CG6D. "
         "https://archive.ics.uci.edu/dataset/502/online+retail+ii", False),
        ("License: Creative Commons Attribution 4.0 (CC BY 4.0). Real transactions of a UK-based online gift retailer, "
         "1 Dec 2009 - 9 Dec 2011. Many of its customers are wholesalers.", False),
        ("This tracker uses the 24 full months Dec 2009 - Nov 2011: fiscal year Dec 2010 - Nov 2011, "
         "with Dec 2009 - Nov 2010 as last year.", False),
        ("", False),
        ("Cleaning log", True),
    ]
    g["B1"] = "How this tracker works and how to update it"
    g["B1"].font = font(14, True, NAVY)
    rr = 3
    for text, bold in lines:
        cell = g.cell(row=rr, column=2, value=text)
        cell.font = font(11 if bold else 10, bold, NAVY if bold else "000000")
        rr += 1
    header_row(g, rr, 2, ["Step", "Rows", "Note"])
    for row in cleaning.itertuples(index=False):
        rr += 1
        g.cell(row=rr, column=2, value=row.Step).font = font(10)
        cell = g.cell(row=rr, column=3, value=int(row.Rows))
        cell.font = font(10)
        cell.number_format = "#,##0"
        g.cell(row=rr, column=4, value=row.Note).font = font(10)
        for col in range(2, 5):
            g.cell(row=rr, column=col).border = BOX
    for col, w in zip("ABCD", (2, 52, 12, 86)):
        g.column_dimensions[col].width = w
    g.sheet_view.showGridLines = False

    # ---------------- Dashboard ----------------
    d = dash
    d.sheet_view.showGridLines = False
    for col, w in {"A": 2, "B": 26, "C": 15, "D": 15, "E": 15, "F": 15, "G": 15, "H": 15, "I": 15,
                   "J": 15, "K": 2}.items():
        d.column_dimensions[col].width = w
    d["B1"] = "Sales Performance Tracker"
    d["B1"].font = font(20, True, NAVY)
    d.row_dimensions[1].height = 30
    note(d, "B2", "UK online gift retailer · real transactions (UCI Online Retail II, CC BY 4.0) · "
                  "fiscal year Dec 2010 - Nov 2011 · GBP")
    d["B4"] = "Month"
    d["B4"].font = font(11, True, NAVY)
    d["B4"].alignment = Alignment(horizontal="right")
    d["C4"] = DEFAULT_MONTH.to_pydatetime()
    d["C4"].number_format = MONTH_FMT
    d["C4"].font = font(12, True, "0000FF")
    d["C4"].fill = fill("FFFF00")
    d["C4"].border = BOX
    d["C4"].alignment = Alignment(horizontal="center")
    d["C4"].comment = Comment("Pick a month from the list. Everything on this tab follows it.", "Tracker")
    dv = DataValidation(type="list", formula1="Settings!$B$9:$B$20", allow_blank=False,
                        showErrorMessage=True, errorTitle="Month", error="Pick a month from the list.")
    d.add_data_validation(dv)
    dv.add("C4")
    note(d, "D4", "← pick a month")
    d["F4"] = "Target rule:"
    d["F4"].font = font(9, True, NAVY)
    d["F4"].alignment = Alignment(horizontal="right")
    d["G4"] = '="Same month last year +"&TEXT(Settings!$B$3,"0%")'
    d["G4"].font = font(9)
    d["I4"] = "Data checks:"
    d["I4"].font = font(9, True, NAVY)
    d["I4"].alignment = Alignment(horizontal="right")
    d["J4"] = "=Checks!$C$11"
    d["J4"].font = font(9, True)
    d.conditional_formatting.add("J4", CellIsRule(operator="equal", formula=['"All checks OK"'],
                                                  font=Font(name=FONT, color="006100", bold=True)))
    d.conditional_formatting.add("J4", CellIsRule(operator="notEqual", formula=['"All checks OK"'],
                                                  font=Font(name=FONT, color="9C0006", bold=True)))

    S = "Sales_Input!"
    T = "Targets!"
    C = "Customers_Input!"
    sel = "$C$4"
    fy = "Settings!$B$4"

    def s_month(extra=""):
        return f"SUMIFS({S}$E:$E,{S}$A:$A,{sel}{extra})"

    def t_month(extra=""):
        return f"SUMIFS({T}$E:$E,{T}$A:$A,{sel}{extra})"

    def s_ytd(extra=""):
        return f'SUMIFS({S}$E:$E,{S}$A:$A,">="&{fy},{S}$A:$A,"<="&{sel}{extra})'

    def t_ytd(extra=""):
        return f'SUMIFS({T}$E:$E,{T}$A:$A,">="&{fy},{T}$A:$A,"<="&{sel}{extra})'

    def by_type(t):
        return f',{S}$D:$D,"{t}"'

    cards1 = [
        ("NET SALES\nMONTH", f"={s_month()}", GBP),
        ("TARGET\nMONTH", f"={t_month()}", GBP),
        ("ACHIEVEMENT\nMONTH", '=IF(C7=0,"",B7/C7)', PCT),
        ("GAP TO TARGET\nMONTH", "=B7-C7", GBP),
        ("NET SALES\nYEAR TO DATE", f"={s_ytd()}", GBP),
        ("TARGET\nYEAR TO DATE", f"={t_ytd()}", GBP),
        ("ACHIEVEMENT\nYEAR TO DATE", '=IF(G7=0,"",F7/G7)', PCT),
        ("GAP TO TARGET\nYEAR TO DATE", "=F7-G7", GBP),
        ("SAME MONTH\nLAST YEAR", f"=SUMIFS({S}$E:$E,{S}$A:$A,EDATE({sel},-12))", GBP),
    ]
    cards2 = [
        ("NEW CUSTOMERS\nMONTH", f'=SUMIFS({C}$D:$D,{C}$A:$A,{sel},{C}$C:$C,"New")', INT),
        ("ACTIVE\nCUSTOMERS", f"=SUMIFS({C}$D:$D,{C}$A:$A,{sel})", INT),
        ("ORDERS\nMONTH", f"=SUMIFS({C}$E:$E,{C}$A:$A,{sel})", INT),
        ("NET SALES\nPER ORDER", '=IF(D10=0,"",B7/D10)', GBP2),
        ("RETURNING\nSHARE OF SALES", f'=IF(B7=0,"",{s_month(by_type("Returning"))}/B7)', PCT),
        ("NEW CUSTOMERS\nSHARE OF SALES", f'=IF(B7=0,"",{s_month(by_type("New"))}/B7)', PCT),
        ("GUEST\nSHARE OF SALES", f'=IF(B7=0,"",{s_month(by_type("Guest (no ID)"))}/B7)', PCT),
        ("NEW CUSTOMERS\nYEAR TO DATE", f'=SUMIFS({C}$D:$D,{C}$A:$A,">="&{fy},{C}$A:$A,"<="&{sel},{C}$C:$C,"New")', INT),
        ("GROWTH VS\nLAST YEAR", '=IF(J7=0,"",B7/J7-1)', PCT),
    ]
    for row_label, row_value, cards in ((6, 7, cards1), (9, 10, cards2)):
        for i, (label, formula, fmt) in enumerate(cards):
            col = 2 + i
            lc = d.cell(row=row_label, column=col, value=label)
            lc.font = font(8, True, NOTE)
            lc.fill = fill(LIGHT)
            lc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            vc = d.cell(row=row_value, column=col, value=formula)
            vc.font = font(13, True, NAVY)
            vc.fill = fill(LIGHT)
            vc.number_format = fmt
            vc.alignment = Alignment(horizontal="center", vertical="center")
            lc.border = Border(left=thin, right=thin, top=thin)
            vc.border = Border(left=thin, right=thin, bottom=thin)
        d.row_dimensions[row_label].height = 26
        d.row_dimensions[row_value].height = 28
    d.row_dimensions[8].height = 8
    add_achievement_colors(d, "D7")
    add_achievement_colors(d, "H7")

    # Row positions: KPIs (6-10), charts (12-43), tables from row 46.
    CHART_ROW1, CHART_ROW2 = 12, 28
    r0 = 46  # header of the market table
    mk_last = r0 + 1 + len(MARKETS)
    p0 = mk_last + 3
    pl_last = p0 + 1 + len(PRODUCT_LINES)
    c0 = pl_last + 3
    ct_last = c0 + 1 + len(CUSTOMER_TYPES)
    m0 = ct_last + 3
    tr_last = m0 + 12

    seg_cols = ["Net sales\nmonth", "Target\nmonth", "Achievement\nmonth", "Net sales\nyear to date",
                "Target\nyear to date", "Achievement\nyear to date", "Gap to target\nyear to date",
                "Last year\nyear to date"]
    seg_fmt = {3: GBP, 4: GBP, 5: PCT, 6: GBP, 7: GBP, 8: PCT, 9: GBP, 10: GBP}

    def segment_table(top, title, names, s_col, t_col):
        d.cell(row=top - 1, column=2, value=title).font = font(11, True, NAVY)
        header_row(d, top, 2, [title.replace("By ", "").capitalize()] + seg_cols, height=30)
        for i, name in enumerate(names + ["Total"]):
            rr = top + 1 + i
            d.cell(row=rr, column=2, value=name)
            if name == "Total":
                for col_letter in "CDFGIJ":
                    d[f"{col_letter}{rr}"] = f"=SUM({col_letter}{top + 1}:{col_letter}{rr - 1})"
            else:
                crit = f",{S}{s_col}:{s_col},$B{rr}"
                tcrit = f",{T}{t_col}:{t_col},$B{rr}"
                d[f"C{rr}"] = f"={s_month(crit)}"
                d[f"D{rr}"] = f"={t_month(tcrit)}"
                d[f"F{rr}"] = f"={s_ytd(crit)}"
                d[f"G{rr}"] = f"={t_ytd(tcrit)}"
                d[f"I{rr}"] = f"=F{rr}-G{rr}"
                d[f"J{rr}"] = (f'=SUMIFS({S}$E:$E,{S}$A:$A,">="&EDATE({fy},-12),'
                               f'{S}$A:$A,"<="&EDATE({sel},-12){crit})')
            d[f"E{rr}"] = f'=IF(D{rr}=0,"",C{rr}/D{rr})'
            d[f"H{rr}"] = f'=IF(G{rr}=0,"",F{rr}/G{rr})'
        last = top + 1 + len(names)
        style_body(d, range(top + 1, last + 1), range(2, 11), seg_fmt, total_row=last)
        add_achievement_colors(d, f"E{top + 1}:E{last}")
        add_achievement_colors(d, f"H{top + 1}:H{last}")

    segment_table(r0, "By market", MARKETS, "$B", "$B")
    segment_table(p0, "By product line", PRODUCT_LINES, "$C", "$C")

    # Customer type table
    d.cell(row=c0 - 1, column=2, value="New vs returning business").font = font(11, True, NAVY)
    header_row(d, c0, 2, ["Customer type", "Net sales\nmonth", "Share\nmonth", "Customers\nmonth",
                          "Orders\nmonth", "Net sales\nyear to date", "Share\nyear to date", "Orders\nyear to date",
                          "Sales per order\nyear to date"], height=30)
    for i, t in enumerate(CUSTOMER_TYPES + ["Total"]):
        rr = c0 + 1 + i
        d.cell(row=rr, column=2, value=t)
        if t == "Total":
            for col_letter in "CEFGI":
                d[f"{col_letter}{rr}"] = f"=SUM({col_letter}{c0 + 1}:{col_letter}{rr - 1})"
        else:
            crit = f",{S}$D:$D,$B{rr}"
            d[f"C{rr}"] = f"={s_month(crit)}"
            d[f"E{rr}"] = f"=SUMIFS({C}$D:$D,{C}$A:$A,{sel},{C}$C:$C,$B{rr})"
            d[f"F{rr}"] = f"=SUMIFS({C}$E:$E,{C}$A:$A,{sel},{C}$C:$C,$B{rr})"
            d[f"G{rr}"] = f"={s_ytd(crit)}"
            d[f"I{rr}"] = f'=SUMIFS({C}$E:$E,{C}$A:$A,">="&{fy},{C}$A:$A,"<="&{sel},{C}$C:$C,$B{rr})'
        d[f"D{rr}"] = f'=IF(C${ct_last}=0,"",C{rr}/C${ct_last})'
        d[f"H{rr}"] = f'=IF(G${ct_last}=0,"",G{rr}/G${ct_last})'
        d[f"J{rr}"] = f'=IF(I{rr}=0,"",G{rr}/I{rr})'
    style_body(d, range(c0 + 1, ct_last + 1), range(2, 11),
               {3: GBP, 4: PCT, 5: INT, 6: INT, 7: GBP, 8: PCT, 9: INT, 10: GBP2}, total_row=ct_last)
    d[f"E{c0 + 3}"].comment = Comment("Guests have no customer ID, so they count as 0 customers.", "Tracker")

    # Month by month table (feeds the trend charts)
    d.cell(row=m0 - 1, column=2, value="Month by month (fiscal year)").font = font(11, True, NAVY)
    header_row(d, m0, 2, ["Month", "Net sales", "Target", "Achievement", "Same month\nlast year",
                          "Growth vs\nlast year", "New", "Returning", "Guest (no ID)"], height=30)
    for k in range(12):
        rr = m0 + 1 + k
        d[f"B{rr}"] = f"=EDATE({fy},{k})"
        d[f"C{rr}"] = f"=SUMIFS({S}$E:$E,{S}$A:$A,$B{rr})"
        d[f"D{rr}"] = f"=SUMIFS({T}$E:$E,{T}$A:$A,$B{rr})"
        d[f"E{rr}"] = f'=IF(D{rr}=0,"",C{rr}/D{rr})'
        d[f"F{rr}"] = f"=SUMIFS({S}$E:$E,{S}$A:$A,EDATE($B{rr},-12))"
        d[f"G{rr}"] = f'=IF(F{rr}=0,"",C{rr}/F{rr}-1)'
        for j, t in enumerate(CUSTOMER_TYPES):
            d.cell(row=rr, column=8 + j, value=f'=SUMIFS({S}$E:$E,{S}$A:$A,$B{rr},{S}$D:$D,"{t}")')
    style_body(d, range(m0 + 1, tr_last + 1), range(2, 11),
               {2: MONTH_FMT, 3: GBP, 4: GBP, 5: PCT, 6: GBP, 7: PCT, 8: GBP, 9: GBP, 10: GBP})
    add_achievement_colors(d, f"E{m0 + 1}:E{tr_last}")
    d.cell(row=m0, column=8).comment = Comment("Columns H-J: net sales by customer type.", "Tracker")
    note(d, f"B{tr_last + 2}", "Targets are an assumption: same month last year +10% (Settings B3). Net sales exclude "
                               "postage, fees and manual adjustments. Definitions and source: see 'How to update'.")

    # ---------------- Charts ----------------
    months = Reference(d, min_col=2, min_row=m0 + 1, max_row=tr_last)

    trend = BarChart()
    trend.type = "col"
    trend.title = chart_title("Net sales vs target by month")
    trend.add_data(Reference(d, min_col=3, min_row=m0, max_row=tr_last), titles_from_data=True)
    trend.set_categories(months)
    trend.y_axis.number_format = '"£"#,##0,"k"'
    trend.x_axis.number_format = "mmm yy"
    trend.series[0].graphicalProperties.solidFill = MID
    trend.series[0].graphicalProperties.line.solidFill = MID
    trend.gapWidth = 60
    target_line = LineChart()
    target_line.add_data(Reference(d, min_col=4, min_row=m0, max_row=tr_last), titles_from_data=True)
    target_line.series[0].graphicalProperties.line.solidFill = "C00000"
    target_line.series[0].graphicalProperties.line.width = 28000
    target_line.series[0].smooth = False
    target_line.series[0].marker.symbol = "none"
    trend += target_line
    trend.legend.position = "b"
    trend.width, trend.height = 16.4, 7.6
    tidy_chart(trend)
    d.add_chart(trend, f"B{CHART_ROW1}")

    by_line = BarChart()
    by_line.type = "bar"
    by_line.title = chart_title("Achievement year to date by product line")
    by_line.add_data(Reference(d, min_col=8, min_row=p0, max_row=pl_last - 1), titles_from_data=True)
    by_line.set_categories(Reference(d, min_col=2, min_row=p0 + 1, max_row=pl_last - 1))
    by_line.x_axis.scaling.orientation = "maxMin"
    by_line.y_axis.number_format = "0%"
    by_line.y_axis.scaling.min = 0
    by_line.y_axis.majorGridlines = None
    by_line.y_axis.crosses = "max"
    by_line.series[0].graphicalProperties.solidFill = MID
    by_line.series[0].graphicalProperties.line.solidFill = MID
    by_line.dataLabels = value_labels("0%")
    by_line.gapWidth = 50
    by_line.legend = None
    by_line.width, by_line.height = 11.6, 7.6
    tidy_chart(by_line)
    d.add_chart(by_line, f"G{CHART_ROW1}")

    mix = BarChart()
    mix.type = "col"
    mix.grouping = "stacked"
    mix.overlap = 100
    mix.title = chart_title("Net sales by customer type")
    mix.add_data(Reference(d, min_col=8, max_col=10, min_row=m0, max_row=tr_last), titles_from_data=True)
    mix.set_categories(months)
    mix.y_axis.number_format = '"£"#,##0,"k"'
    mix.x_axis.number_format = "mmm yy"
    for s, color in zip(mix.series, ("70AD47", MID, "A5A5A5")):
        s.graphicalProperties.solidFill = color
        s.graphicalProperties.line.solidFill = color
    mix.gapWidth = 60
    mix.legend.position = "b"
    mix.width, mix.height = 16.4, 7.6
    tidy_chart(mix)
    d.add_chart(mix, f"B{CHART_ROW2}")

    by_market = BarChart()
    by_market.type = "bar"
    by_market.title = chart_title("Achievement year to date by market")
    by_market.add_data(Reference(d, min_col=8, min_row=r0, max_row=mk_last - 1), titles_from_data=True)
    by_market.set_categories(Reference(d, min_col=2, min_row=r0 + 1, max_row=mk_last - 1))
    by_market.x_axis.scaling.orientation = "maxMin"
    by_market.y_axis.number_format = "0%"
    by_market.y_axis.scaling.min = 0
    by_market.y_axis.majorGridlines = None
    by_market.y_axis.crosses = "max"
    by_market.series[0].graphicalProperties.solidFill = MID
    by_market.series[0].graphicalProperties.line.solidFill = MID
    by_market.dataLabels = value_labels("0%")
    by_market.gapWidth = 80
    by_market.legend = None
    by_market.width, by_market.height = 11.6, 7.6
    tidy_chart(by_market)
    d.add_chart(by_market, f"G{CHART_ROW2}")

    # ---------------- Print setup ----------------
    d.freeze_panes = "A5"
    d.print_area = f"A1:K{tr_last + 2}"
    d.page_setup.orientation = "landscape"
    d.page_setup.paperSize = d.PAPERSIZE_A4
    d.page_setup.scale = 72
    d.page_margins = PageMargins(left=0.4, right=0.4, top=0.4, bottom=0.4, header=0.2, footer=0.2)
    d.row_breaks = RowBreak()
    d.row_breaks.append(Break(id=r0 - 3))
    d.print_options.horizontalCentered = True

    for ws in (guide, s_in, c_in, tgt, chk, st, mp):
        ws.page_setup.orientation = "portrait" if ws in (s_in, c_in, tgt) else "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5, header=0.2, footer=0.2)
    for ws in (s_in, c_in, tgt):
        ws.print_title_rows = "1:1"

    wb.calculation.fullCalcOnLoad = True
    wb.save(out_path)
    print("saved", out_path)
    print(f"market {r0}-{mk_last} | product {p0}-{pl_last} | customers {c0}-{ct_last} | trend {m0}-{tr_last}")


if __name__ == "__main__":
    build(sys.argv[1], sys.argv[2])
