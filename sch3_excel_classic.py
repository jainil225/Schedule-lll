# -*- coding: utf-8 -*-
"""
sch3_excel_classic.py  — Final Version
=======================================
Classic Schedule III Excel workbook generator.

Sheets
------
1. How to Link   — Step-by-step CA guide with Tally group → Schedule III mapping
2. TB Data       — Raw Tally values (green cells = link targets for CA)
3. Balance Sheet — Schedule III BS with Excel formulas + unit conversion
4. P&L Statement — Schedule III P&L with Excel formulas + unit conversion
"""

import re, sys, os

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

try:
    import tally_core as tc
except ImportError:
    tc = None

try:
    import sch3_classic as sc
    bs_figures        = sc.bs_figures
    pnl_figures       = sc.pnl_figures
    fmt_date          = sc.fmt_date
    _py_is_same_as_cy = sc._py_is_same_as_cy
except ImportError:
    sc = bs_figures = pnl_figures = fmt_date = _py_is_same_as_cy = None

# ── Colours ───────────────────────────────────────────────────────────────────
C_DARK   = "1F3864"
C_MID    = "2E5496"
C_LIGHT  = "D6E4F0"
C_ALT    = "EEF4FB"
C_WHITE  = "FFFFFF"
C_GREEN  = "375623"
C_LGREEN = "E8F5E9"
C_GOLD   = "7F6000"
C_LGOLD  = "FFF2CC"
C_RED    = "C00000"
C_ORANGE = "C55A11"
C_GREY   = "666666"

thin  = Side(style="thin",   color="8EA9C1")
thick = Side(style="medium", color=C_DARK)
b_all = Border(left=thin,  right=thin,  top=thin,  bottom=thin)
b_top = Border(left=thin,  right=thin,  top=thick, bottom=thin)
b_hdr = Border(left=thick, right=thick, top=thick, bottom=thick)

def F(bold=False, sz=9, color=C_DARK, italic=False):
    return Font(name="Arial", bold=bold, size=sz, color=color, italic=italic)
def FILL(c): return PatternFill("solid", fgColor=c)
def AL(h="left", v="center", wrap=False, indent=0):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap, indent=indent)

def _fmt_date(d):
    if fmt_date: return fmt_date(d)
    s = str(d or "").strip()
    if len(s)==10 and s[4]=="-": return f"{s[8:10]}.{s[5:7]}.{s[:4]}"
    return s

def _f(v):
    try: return float(v or 0)
    except: return 0.0

# ── Tally Group mapping for each Schedule III line ────────────────────────────
# (Description, Tally Groups to sum, Sign note)
TALLY_MAP = {
    # BS Liabilities
    "eq_share":     ("Share Capital",                   "Capital Account / Equity Share Capital",           "CR balance = positive"),
    "reserves":     ("Reserves & Surplus",              "Reserves & Surplus / P&L A/c (CR balance)",       "Net of P&L + Reserves"),
    "lt_borrow":    ("Long-Term Borrowings",            "Loans (Liability) — Term Loans, Secured/Unsecured","CR balance = positive"),
    "dtl":          ("Deferred Tax Liabilities (Net)",  "Deferred Tax Liability",                           "CR balance = positive"),
    "st_borrow":    ("Short-Term Borrowings",           "Bank OD / CC / Short-term loans (Liability)",      "CR balance = positive"),
    "trade_pay":    ("Trade Payables",                  "Sundry Creditors",                                 "CR balance = positive"),
    "other_cl":     ("Other Current Liabilities",       "Duties & Taxes / Other Liabilities / Advances",   "CR balance only"),
    "short_prov":   ("Short-Term Provisions",           "Provisions / Advance Tax / TDS Payable",           "CR balance = positive"),
    # BS Assets
    "fa_net":       ("Tangible Assets (Net)",           "Fixed Assets (Net of Depreciation / WDV)",         "DR balance → positive display"),
    "inv_total":    ("Non-Current Investments",         "Investments (FDs, Shares, Bonds)",                 "DR balance → positive display"),
    "lt_loans_adv": ("Long-Term Loans & Advances",      "Security Deposits / Capital Advances",             "DR balance → positive display"),
    "other_nca":    ("Other Non-Current Assets",        "Misc. Expenses (Asset) / Unadjusted Forex",        "DR balance → positive display"),
    "inventories":  ("Inventories",                     "Stock-in-Hand / Closing Stock",                    "DR balance → positive display"),
    "trade_recv":   ("Trade Receivables",               "Sundry Debtors (DR balance only)",                 "DR balance → positive display"),
    "cash_bank":    ("Cash & Cash Equivalents",         "Cash-in-Hand + Bank Accounts",                     "DR balance → positive display"),
    "st_loans_adv": ("Short-Term Loans & Advances",     "Loans & Advances (Asset) / TDS Rcvbl / Input GST","DR balance → positive display"),
    "suspense_val": ("Other Current Assets (Suspense)", "Suspense A/c (DR balance only)",                   "DR balance → positive display"),
    # P&L Revenue
    "pnl_revenue":  ("Revenue from Operations",         "Sales Accounts / Revenue from Operations",         "CR balance → positive display"),
    "pnl_other_inc":("Other Income",                    "Indirect Incomes (Interest, Rent etc.)",           "CR balance → positive display"),
    # P&L Expenses
    "pnl_cost_mat": ("Cost of Material Consumed",       "Purchase Accounts + Direct Expenses",              "DR balance → positive display"),
    "pnl_inv_chg":  ("Changes in Inventories",          "Opening Stock − Closing Stock",                    "DR balance → positive display"),
    "pnl_emp_exp":  ("Employee Benefit Expense",        "Salaries, Wages, ESIC, PF (Indirect Expenses)",    "DR balance → positive display"),
    "pnl_fin_cost": ("Financial Costs",                 "Interest + Bank Charges (Indirect Expenses)",      "DR balance → positive display"),
    "pnl_dep":      ("Depreciation & Amortisation",     "Depreciation (from FA schedule / P&L)",            "DR balance → positive display"),
    "pnl_other_exp":("Other Expenses",                  "Remaining Indirect Expenses (Admin, Selling etc.)","DR balance → positive display"),
    "pnl_tax":      ("Current Tax",                     "Income Tax / Provision for Tax",                   "DR balance → positive display"),
}

UNIT_OPTIONS      = ["Rs. (Actual)", "Thousands", "Lakhs", "Crores", "Millions"]
UNIT_DIVISORS     = [1, 1_000, 100_000, 10_000_000, 1_000_000]

def _div_fml(cell):
    opts = ",".join(f'"{u}"' for u in UNIT_OPTIONS)
    divs = ",".join(str(d) for d in UNIT_DIVISORS)
    return f'CHOOSE(MATCH({cell},{{{opts}}},0),{divs})'

def _lbl_fml(cell):
    opts = ",".join(f'"{u}"' for u in UNIT_OPTIONS)
    lbls = ",".join(f'"{u}"' for u in UNIT_OPTIONS)
    return f'CHOOSE(MATCH({cell},{{{opts}}},0),{lbls})'


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 1 — HOW TO LINK (CA Guide)
# ══════════════════════════════════════════════════════════════════════════════
def _build_guide_sheet(ws, company, fy_e, fy_p):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 32
    ws.column_dimensions["C"].width = 48
    ws.column_dimensions["D"].width = 32
    ws.column_dimensions["E"].width = 28

    ri = [1]
    def nr(): r=ri[0]; ri[0]+=1; return r

    # Title
    r = nr()
    ws.merge_cells(f"A{r}:E{r}")
    c = ws.cell(r,1, f"SCHEDULE III EXCEL — TALLY MAPPING GUIDE")
    c.font = F(bold=True, sz=14, color=C_WHITE)
    c.fill = FILL(C_DARK); c.alignment = AL("center"); ws.row_dimensions[r].height=28

    r = nr()
    ws.merge_cells(f"A{r}:E{r}")
    c = ws.cell(r,1, f"{company.upper()}   |   FY: {_fmt_date(fy_e)}")
    c.font = F(bold=False, sz=9, color="AAAAAA", italic=True)
    c.fill = FILL(C_DARK); c.alignment = AL("center"); ws.row_dimensions[r].height=16

    # Gap
    r = nr(); ws.row_dimensions[r].height=8

    # Steps
    steps = [
        ("STEP 1", "Open Both Files",
         "Keep this Excel open. Also open your Trial Balance (TB) file exported from Tally or your own TB sheet."),
        ("STEP 2", "Go to TB Data Sheet",
         "Click the green 'TB Data' tab. All green cells in Column C (Current Year) and Column D (Previous Year) are the link targets."),
        ("STEP 3", "Link Each Green Cell",
         "Click a green cell → type = → switch to your TB file → click the matching group total → press Enter.\n"
         "Example:  Share Capital cell → = → [TB.xlsx]Sheet1!$C$5"),
        ("STEP 4", "Balance Sheet Updates Automatically",
         "Once you link the TB Data cells, the Balance Sheet and P&L Statement sheets recalculate instantly. No manual entry needed."),
        ("STEP 5", "Change Display Unit",
         "On the Balance Sheet or P&L sheet, click cell F2 and select:\n"
         "Rs. (Actual)  |  Thousands  |  Lakhs  |  Crores  |  Millions\n"
         "All amounts convert automatically. The heading also updates."),
        ("STEP 6", "Verify Balance",
         "Check the 'Balance Check' row at the bottom of the Balance Sheet.\n"
         "It must show 0.00. If not, re-check your TB group totals and sign conventions (see table below)."),
    ]

    for num, title, detail in steps:
        r = nr()
        ws.merge_cells(f"A{r}:E{r}")
        c = ws.cell(r,1, f"  {num}  —  {title}")
        c.font = F(bold=True, sz=10, color=C_WHITE)
        c.fill = FILL(C_MID); c.alignment = AL("left"); ws.row_dimensions[r].height=18

        for line in detail.split("\n"):
            r = nr()
            ws.merge_cells(f"B{r}:E{r}")
            c = ws.cell(r,2, line)
            c.font = F(bold=False, sz=9, color=C_DARK)
            c.alignment = AL("left", wrap=True)
            ws.row_dimensions[r].height = 14 if "\n" not in detail else 13
        nr_gap = nr(); ws.row_dimensions[nr_gap].height = 5

    # Sign convention note
    r = nr()
    ws.merge_cells(f"A{r}:E{r}")
    c = ws.cell(r,1, "  IMPORTANT — SIGN CONVENTION")
    c.font = F(bold=True, sz=10, color=C_WHITE)
    c.fill = FILL(C_ORANGE); c.alignment = AL("left"); ws.row_dimensions[r].height=18

    notes = [
        "Tally Trial Balance shows CR balances as POSITIVE and DR balances as NEGATIVE.",
        "Liabilities (Share Capital, Loans, Creditors) → CR balance → enter as positive number.",
        "Assets (Fixed Assets, Debtors, Cash, Stock) → DR balance in Tally → TallySync already converts these to positive display values.",
        "If you link directly from Tally TB export, Assets may appear negative — multiply by -1 or use ABS() in your link formula.",
        "Reserves & Surplus = Capital Account Total − Share Capital + P&L Balance (signed — can be negative for loss companies).",
    ]
    for note in notes:
        r = nr()
        ws.merge_cells(f"B{r}:E{r}")
        c = ws.cell(r,2, f"▸  {note}")
        c.font = F(bold=False, sz=9, color="8B2500")
        c.fill = FILL("FFF0E6"); c.alignment = AL("left", wrap=True); ws.row_dimensions[r].height=14
    nr_gap = nr(); ws.row_dimensions[nr_gap].height=8

    # Full mapping table
    r = nr()
    ws.merge_cells(f"A{r}:E{r}")
    c = ws.cell(r,1, "  TALLY GROUP → SCHEDULE III MAPPING TABLE")
    c.font = F(bold=True, sz=10, color=C_WHITE)
    c.fill = FILL(C_DARK); c.alignment = AL("left"); ws.row_dimensions[r].height=18

    # Table header
    r = nr(); ws.row_dimensions[r].height=22
    for col, (txt, w) in enumerate([
        ("Schedule III Line", 32), ("Tally Groups to Sum", 48),
        ("TB Data Cell", 20), ("Sign Note", 28)], 2):
        c = ws.cell(r, col, txt)
        c.font = F(bold=True, sz=9, color=C_WHITE)
        c.fill = FILL(C_MID); c.alignment = AL("center", wrap=True); c.border = b_all

    sections = [
        ("BS — EQUITY & LIABILITIES", ["eq_share","reserves","lt_borrow","dtl","st_borrow","trade_pay","other_cl","short_prov"]),
        ("BS — ASSETS",               ["fa_net","inv_total","lt_loans_adv","other_nca","inventories","trade_recv","cash_bank","st_loans_adv","suspense_val"]),
        ("P&L — REVENUE",             ["pnl_revenue","pnl_other_inc"]),
        ("P&L — EXPENSES",            ["pnl_cost_mat","pnl_inv_chg","pnl_emp_exp","pnl_fin_cost","pnl_dep","pnl_other_exp","pnl_tax"]),
    ]

    row_num = 0
    for sec_label, keys in sections:
        r = nr()
        ws.merge_cells(f"B{r}:E{r}")
        c = ws.cell(r,2, sec_label)
        c.font = F(bold=True, sz=8.5, color=C_WHITE)
        c.fill = FILL(C_MID); c.alignment = AL("left"); ws.row_dimensions[r].height=14
        c.border = b_all

        for key in keys:
            if key not in TALLY_MAP: continue
            desc, tally_grp, sign = TALLY_MAP[key]
            r = nr(); ws.row_dimensions[r].height=14
            bg = FILL(C_ALT) if row_num%2==0 else FILL(C_WHITE)
            row_num += 1

            c = ws.cell(r,2, desc)
            c.font = F(sz=8.5); c.fill=bg; c.alignment=AL("left"); c.border=b_all

            c = ws.cell(r,3, tally_grp)
            c.font = F(sz=8.5, color=C_GREEN); c.fill=bg; c.alignment=AL("left",wrap=True); c.border=b_all

            c = ws.cell(r,4, f"TB Data → Col C/D")
            c.font = F(sz=8, color=C_GOLD, italic=True); c.fill=bg; c.alignment=AL("center"); c.border=b_all

            c = ws.cell(r,5, sign)
            c.font = F(sz=8, color=C_GREY, italic=True); c.fill=bg; c.alignment=AL("left",wrap=True); c.border=b_all

    # Footer
    r = nr(); ws.row_dimensions[r].height=8
    r = nr()
    ws.merge_cells(f"A{r}:E{r}")
    c = ws.cell(r,1, "Generated by TallySync Pro  |  Schedule III Financial Statements")
    c.font = F(sz=7.5, color=C_GREY, italic=True); c.alignment=AL("center")

    ws.sheet_properties.tabColor = "E8F5E9"


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 2 — TB DATA
# ══════════════════════════════════════════════════════════════════════════════
class TBDataSheet:
    HFILL = FILL(C_DARK);  HFONT = F(bold=True, sz=9, color=C_WHITE)
    SFILL = FILL(C_MID);   SFONT = F(bold=True, sz=8.5, color=C_WHITE)
    VFONT = F(sz=8.5, color=C_GREEN)
    LFONT = F(sz=8.5, color=C_DARK)
    IFILL = FILL(C_LGREEN)

    def __init__(self, ws):
        self.ws = ws; self.ri = 1; self.map = {}
        ws.column_dimensions["A"].width = 20   # Section
        ws.column_dimensions["B"].width = 34   # Description
        ws.column_dimensions["C"].width = 38   # Tally Groups
        ws.column_dimensions["D"].width = 20   # CY
        ws.column_dimensions["E"].width = 20   # PY
        ws.sheet_view.showGridLines = True

    def _nr(self):
        r=self.ri; self.ri+=1; return r

    def write_header(self, company, fy_e, fy_p):
        ws=self.ws
        r=self._nr()
        ws.merge_cells(f"A{r}:E{r}")
        c=ws.cell(r,1,f"TB DATA — {company.upper()}")
        c.font=F(bold=True,sz=11,color=C_DARK); c.alignment=AL("center"); ws.row_dimensions[r].height=18

        r=self._nr()
        ws.merge_cells(f"A{r}:E{r}")
        c=ws.cell(r,1,
            "▶  GREEN CELLS = Link targets. Replace with = references from your Trial Balance. "
            "Column C = Current Year  |  Column D = Previous Year")
        c.font=F(sz=8,color=C_GOLD,italic=True); c.alignment=AL("left"); ws.row_dimensions[r].height=13

        r=self._nr(); ws.row_dimensions[r].height=24
        for col,txt in enumerate(["SECTION","SCHEDULE III LINE",
                                   "TALLY GROUPS TO LINK",
                                   f"CURRENT YEAR\n{_fmt_date(fy_e)} (₹)",
                                   f"PREVIOUS YEAR\n{_fmt_date(fy_p)} (₹)"],1):
            c=ws.cell(r,col,txt)
            c.font=self.HFONT; c.fill=self.HFILL
            c.alignment=AL("center",wrap=True); c.border=b_all
        ws.freeze_panes="A4"

    def _section(self, label):
        ws=self.ws; r=self._nr(); ws.row_dimensions[r].height=13
        ws.merge_cells(f"A{r}:E{r}")
        c=ws.cell(r,1,label)
        c.font=self.SFONT; c.fill=self.SFILL; c.alignment=AL("left"); c.border=b_all

    def _row(self, key, section, desc, cy, py):
        ws=self.ws; r=self._nr(); ws.row_dimensions[r].height=14
        fill=FILL(C_ALT) if r%2==0 else FILL(C_WHITE)
        tally_grp = TALLY_MAP.get(key, ("","",""))[1]
        sign_note = TALLY_MAP.get(key, ("","",""))[2]

        for col,val,fnt,fll,aln in [
            (1, section,   self.LFONT, fill,       AL("left")),
            (2, desc,      self.LFONT, fill,       AL("left")),
            (3, tally_grp, F(sz=8,color=C_GREEN), fill, AL("left",wrap=True)),
        ]:
            c=ws.cell(r,col,val); c.font=fnt; c.fill=fll; c.alignment=aln; c.border=b_all

        for col,val in [(4,cy),(5,py)]:
            v=_f(val) if val is not None else 0.0
            c=ws.cell(r,col,v)
            c.font=self.VFONT; c.fill=self.IFILL
            c.number_format='#,##0.00'; c.alignment=AL("right"); c.border=b_all

        # map stores (cy_float, py_float) for direct use in BS/P&L sheets
        self.map[key] = (_f(cy), _f(py) if py is not None else 0.0)
        return r

    def populate(self, f, fp, m, mp):
        R=self._row
        def g(d,k,default=None):
            v=d.get(k) if d else None
            return v if v is not None else default

        self._section("BS — EQUITY & LIABILITIES")
        R("eq_share",    "Equity & Liab", "Share Capital",                  g(f,"eq_share"),    g(fp,"eq_share"))
        R("reserves",    "Equity & Liab", "Reserves & Surplus",             g(f,"reserves"),    g(fp,"reserves"))
        R("lt_borrow",   "Equity & Liab", "Long-Term Borrowings",           g(f,"lt_borrow"),   g(fp,"lt_borrow"))
        R("dtl",         "Equity & Liab", "Deferred Tax Liabilities (Net)", g(f,"dtl"),         g(fp,"dtl"))
        R("st_borrow",   "Equity & Liab", "Short-Term Borrowings",          g(f,"st_borrow"),   g(fp,"st_borrow"))
        R("trade_pay",   "Equity & Liab", "Trade Payables",                 g(f,"trade_pay"),   g(fp,"trade_pay"))
        R("other_cl",    "Equity & Liab", "Other Current Liabilities",
            f.get("other_cl_total", f.get("other_cl",0)),
            fp.get("other_cl_total", fp.get("other_cl",0)) if fp else None)
        R("short_prov",  "Equity & Liab", "Short-Term Provisions",          g(f,"short_prov"),  g(fp,"short_prov"))

        self._section("BS — ASSETS")
        R("fa_net",       "Assets", "Tangible Assets (Net)",                g(f,"fa_net"),       g(fp,"fa_net"))
        R("inv_total",    "Assets", "Non-Current Investments",              g(f,"inv_total",0),  g(fp,"inv_total",0))
        R("lt_loans_adv", "Assets", "Long-Term Loans & Advances",           g(f,"lt_loans_adv"), g(fp,"lt_loans_adv"))
        R("other_nca",    "Assets", "Other Non-Current Assets",             g(f,"other_nca",0),  g(fp,"other_nca",0))
        R("inventories",  "Assets", "Inventories",                          g(f,"inventories"),  g(fp,"inventories"))
        R("trade_recv",   "Assets", "Trade Receivables",                    g(f,"trade_recv"),   g(fp,"trade_recv"))
        R("cash_bank",    "Assets", "Cash & Cash Equivalents",              g(f,"cash_bank"),    g(fp,"cash_bank"))
        R("st_loans_adv", "Assets", "Short-Term Loans & Advances",          g(f,"st_loans_adv"), g(fp,"st_loans_adv"))
        R("suspense_val", "Assets", "Other Current Assets (Suspense)",      g(f,"suspense_val",0),g(fp,"suspense_val",0))

        self._section("P&L — REVENUE")
        R("pnl_revenue",  "P&L Revenue", "Revenue from Operations", g(m,"revenue"),    g(mp,"revenue"))
        R("pnl_other_inc","P&L Revenue", "Other Income",            g(m,"other_inc"),  g(mp,"other_inc"))

        self._section("P&L — EXPENSES")
        R("pnl_cost_mat", "P&L Expenses","Cost of Material Consumed",       g(m,"cost_mat"),  g(mp,"cost_mat"))
        R("pnl_inv_chg",  "P&L Expenses","Changes in Inventories",          g(m,"inv_chg"),   g(mp,"inv_chg"))
        R("pnl_emp_exp",  "P&L Expenses","Employee Benefit Expense",        g(m,"emp_exp"),   g(mp,"emp_exp"))
        R("pnl_fin_cost", "P&L Expenses","Financial Costs",                 g(m,"fin_cost"),  g(mp,"fin_cost"))
        R("pnl_dep",      "P&L Expenses","Depreciation & Amortisation",     g(m,"dep"),       g(mp,"dep"))
        R("pnl_other_exp","P&L Expenses","Other Expenses",                  g(m,"other_exp"), g(mp,"other_exp"))
        R("pnl_tax",      "P&L Expenses","Current Tax",                     g(m,"tax"),       g(mp,"tax"))

        # Balance check row
        self._section("BALANCE CHECK (do not edit)")
        ws=self.ws; r=self._nr(); ws.row_dimensions[r].height=14
        ws.merge_cells(f"A{r}:C{r}")
        c=ws.cell(r,1,"BS Total Liabilities = BS Total Assets (Diff should be 0)")
        c.font=F(bold=True,sz=8.5,color=C_GOLD); c.fill=FILL(C_LGOLD); c.alignment=AL("left"); c.border=b_all

        l_keys=["eq_share","reserves","lt_borrow","dtl","st_borrow","trade_pay","other_cl","short_prov"]
        a_keys=["fa_net","inv_total","lt_loans_adv","other_nca","inventories","trade_recv","cash_bank","st_loans_adv","suspense_val"]

        for col_idx, col in [(0,4),(1,5)]:
            l_sum = round(sum(self.map[k][col_idx] for k in l_keys if k in self.map), 2)
            a_sum = round(sum(self.map[k][col_idx] for k in a_keys if k in self.map), 2)
            c=ws.cell(r,col, round(a_sum - l_sum, 2))
            c.font=F(bold=True,sz=8.5,color=C_DARK); c.fill=FILL(C_LGOLD)
            c.number_format='#,##0.00'; c.alignment=AL("right"); c.border=b_all


# ══════════════════════════════════════════════════════════════════════════════
# UNIT SELECTOR (shared helper)
# ══════════════════════════════════════════════════════════════════════════════
def _unit_box(ws, row, fy_e, fy_p):
    c=ws.cell(row,5,"Display Amounts In:")
    c.font=F(bold=True,sz=8.5,color=C_GOLD); c.alignment=AL("right")

    c=ws.cell(row,6,"Lakhs")
    c.font=F(bold=True,sz=10,color=C_DARK); c.fill=FILL(C_LGOLD)
    c.alignment=AL("center"); c.border=b_hdr
    unit_cell=f"F{row}"

    dv=DataValidation(type="list",formula1=f'"{",".join(UNIT_OPTIONS)}"',
                      allow_blank=False,showDropDown=False)
    dv.sqref=unit_cell; ws.add_data_validation(dv)

    r2=row+1
    ws.merge_cells(f"E{r2}:F{r2}")
    c=ws.cell(r2,5,"↑ Click to change unit — all amounts update")
    c.font=F(sz=7.5,color=C_GREY,italic=True); c.alignment=AL("center")
    return unit_cell


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 3 — BALANCE SHEET
# ══════════════════════════════════════════════════════════════════════════════
def _build_bs_sheet(ws, tb_map, company, fy_e, fy_p):
    ws.sheet_view.showGridLines=False
    ws.column_dimensions["A"].width=46
    ws.column_dimensions["B"].width=10
    ws.column_dimensions["C"].width=20
    ws.column_dimensions["D"].width=20
    ws.column_dimensions["E"].width=22
    ws.column_dimensions["F"].width=18

    ri=[1]
    def nr(): r=ri[0]; ri[0]+=1; return r

    # Title
    r=nr(); ws.merge_cells(f"A{r}:D{r}")
    c=ws.cell(r,1,company.upper())
    c.font=F(bold=True,sz=13,color=C_DARK); c.alignment=AL("center"); ws.row_dimensions[r].height=22

    r2=nr(); ws.merge_cells(f"A{r2}:D{r2}")
    c=ws.cell(r2,1,f"BALANCE SHEET AS AT {_fmt_date(fy_e)}")
    c.font=F(bold=True,sz=10,color=C_DARK); c.alignment=AL("center")
    unit_cell=_unit_box(ws,r2,fy_e,fy_p)

    r3=nr(); ws.merge_cells(f"C{r3}:D{r3}")
    lbl=_lbl_fml(unit_cell)
    ws.cell(r3,3,f'="(Amt. in "&{lbl}&")"').font=F(sz=7.5,color=C_GREY,italic=True)
    ws.cell(r3,3).alignment=AL("right")

    r4=nr(); ws.row_dimensions[r4].height=30
    for col,txt in enumerate(["PARTICULARS","Note No.",
        f"AS AT\n{_fmt_date(fy_e)}",f"AS AT\n{_fmt_date(fy_p)}"],1):
        c=ws.cell(r4,col,txt)
        c.font=F(bold=True,sz=9,color=C_WHITE); c.fill=FILL(C_DARK)
        c.alignment=AL("center" if col>1 else "left",wrap=True); c.border=b_all

    # tb_map now holds (cy_float, py_float) — write actual values, no formulas needed
    DIVISOR = 100_000  # Lakhs — matches default unit cell value

    def val(key, ci):
        if key not in tb_map: return None
        v = _f(tb_map[key][ci])
        return round(v / DIVISOR, 2) if v != 0 else None

    def _sec(lbl):
        r=nr(); ws.row_dimensions[r].height=13
        ws.merge_cells(f"A{r}:D{r}")
        c=ws.cell(r,1,lbl); c.font=F(bold=True,sz=8.5,color=C_WHITE)
        c.fill=FILL(C_MID); c.alignment=AL("left"); c.border=b_all

    def _drow(lbl,note,key,bold=False,ind=2,cy_ov=None,py_ov=None):
        r=nr(); ws.row_dimensions[r].height=13
        fill=FILL(C_ALT) if r%2==0 else FILL(C_WHITE)
        c=ws.cell(r,1,"    "*ind+lbl)
        c.font=F(bold=bold,sz=8.5,color=C_DARK); c.fill=fill
        c.alignment=AL("left"); c.border=b_all
        c=ws.cell(r,2,note or ""); c.font=F(sz=8)
        c.fill=fill; c.alignment=AL("center"); c.border=b_all
        for col,ci,ov in [(3,0,cy_ov),(4,1,py_ov)]:
            v = ov if ov is not None else val(key, ci)
            cv=ws.cell(r,col, v)
            cv.font=F(bold=bold,sz=8.5,color=C_DARK)
            cv.fill=FILL(C_LIGHT) if bold else fill
            cv.alignment=AL("right"); cv.border=b_all
            if v is not None: cv.number_format='#,##0.00'

    def _trow(lbl, cy_val, py_val):
        r=nr(); ws.row_dimensions[r].height=15
        c=ws.cell(r,1,lbl); c.font=F(bold=True,sz=9,color=C_DARK)
        c.fill=FILL(C_LIGHT); c.border=b_top; c.alignment=AL("left")
        ws.cell(r,2,"").fill=FILL(C_LIGHT); ws.cell(r,2).border=b_top
        for col,v in [(3,cy_val),(4,py_val)]:
            cv=ws.cell(r,col, round(v,2) if v else None)
            cv.font=F(bold=True,sz=9,color=C_DARK)
            cv.fill=FILL(C_LIGHT); cv.alignment=AL("right"); cv.border=b_top
            if v: cv.number_format='#,##0.00'

    def _brow(lbl,ind=2):
        r=nr(); ws.row_dimensions[r].height=11
        c=ws.cell(r,1,"    "*ind+lbl); c.font=F(sz=8.5,color=C_GREY)
        c.fill=FILL(C_WHITE); c.alignment=AL("left"); c.border=b_all
        for col in [2,3,4]:
            ws.cell(r,col).fill=FILL(C_WHITE); ws.cell(r,col).border=b_all

    def tsum(keys, ci):
        return round(sum(_f(tb_map[k][ci]) / DIVISOR for k in keys if k in tb_map), 2)

    L_KEYS=["eq_share","reserves","lt_borrow","dtl","st_borrow","trade_pay","other_cl","short_prov"]
    A_KEYS=["fa_net","inv_total","lt_loans_adv","other_nca","inventories","trade_recv","cash_bank","st_loans_adv","suspense_val"]

    _sec("I.  EQUITY & LIABILITIES")
    _sec("    SHAREHOLDER'S FUNDS")
    _drow("SHARE CAPITAL",            "1",  "eq_share",   ind=2)
    _drow("RESERVES & SURPLUS",       "2",  "reserves",   ind=2)
    _brow("MONEY RECEIVED AGAINST SHARE WARRANTS")
    _brow("SHARE APPLICATION MONEY PENDING ALLOTMENT",ind=1)
    _sec("    NON-CURRENT LIABILITIES")
    _drow("LONG-TERM BORROWINGS",          "3",  "lt_borrow",  ind=2)
    _drow("DEFERRED TAX LIABILITIES (NET)","32", "dtl",         ind=2)
    _brow("OTHER LONG TERM LIABILITIES")
    _brow("LONG-TERM PROVISIONS")
    _sec("    CURRENT LIABILITIES")
    _drow("SHORT-TERM BORROWINGS",    "4",  "st_borrow",  ind=2)
    _drow("TRADE PAYABLES",           "5",  "trade_pay",  ind=2)
    _drow("OTHER CURRENT LIABILITIES","6",  "other_cl",   ind=2)
    _drow("SHORT TERM PROVISION",     "7",  "short_prov", ind=2)
    _trow("TOTAL EQUITY & LIABILITIES", tsum(L_KEYS,0), tsum(L_KEYS,1))

    _sec("II.  ASSETS")
    _sec("    NON-CURRENT ASSETS")
    _sec("    Property Plant and Equipment")
    _drow("TANGIBLE ASSETS",              "8",   "fa_net",       ind=2)
    _brow("INTANGIBLE ASSETS")
    _brow("CAPITAL WORK-IN-PROGRESS")
    _brow("INTANGIBLE ASSETS UNDER DEVELOPMENT")
    _drow("NON-CURRENT INVESTMENTS",      "10",  "inv_total",    ind=1)
    _brow("DEFERRED TAX ASSET (NET)",ind=1)
    _drow("LONG-TERM LOANS AND ADVANCES", "9",   "lt_loans_adv", ind=1)
    _drow("OTHER NON-CURRENT ASSETS",     "9A",  "other_nca",    ind=1)
    _sec("    CURRENT ASSETS")
    _brow("CURRENT INVESTMENTS")
    _drow("INVENTORIES",                  "11",  "inventories",  ind=2)
    _drow("TRADE RECEIVABLES",            "12",  "trade_recv",   ind=2)
    _drow("CASH AND CASH EQUIVALENTS",    "13",  "cash_bank",    ind=2)
    _drow("SHORT TERM LOANS AND ADVANCES","14",  "st_loans_adv", ind=2)
    _drow("OTHER CURRENT ASSETS",         "15A", "suspense_val", ind=2)
    _trow("TOTAL ASSETS", tsum(A_KEYS,0), tsum(A_KEYS,1))
    _drow("III.  CONTINGENT LIABILITIES","31",None,ind=0)

    # Balance check
    r_chk=nr(); ws.row_dimensions[r_chk].height=13
    ws.merge_cells(f"A{r_chk}:B{r_chk}")
    c=ws.cell(r_chk,1,"Balance Check  (Assets − Liabilities = 0)")
    c.font=F(sz=7.5,color=C_GOLD,italic=True); c.alignment=AL("left")
    for col,ci in [(3,0),(4,1)]:
        diff = round(tsum(A_KEYS,ci) - tsum(L_KEYS,ci), 2)
        c=ws.cell(r_chk,col, diff)
        c.font=F(bold=True,sz=8.5,color=C_RED); c.number_format='#,##0.00'
        c.alignment=AL("right")

    r_ft=nr()
    ws.merge_cells(f"A{r_ft}:D{r_ft}")
    ws.cell(r_ft,1,"See accompanying notes to the financial statements").font=F(sz=7.5,color=C_GREY,italic=True)
    ws.cell(r_ft,1).alignment=AL("left")

    ws.print_area="A1:D60"
    ws.page_setup.orientation="portrait"
    ws.sheet_properties.tabColor=C_DARK


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 4 — P&L STATEMENT
# ══════════════════════════════════════════════════════════════════════════════
def _build_pnl_sheet(ws, tb_map, company, fy_e, fy_p):
    ws.sheet_view.showGridLines=False
    ws.column_dimensions["A"].width=54
    ws.column_dimensions["B"].width=10
    ws.column_dimensions["C"].width=20
    ws.column_dimensions["D"].width=20
    ws.column_dimensions["E"].width=22
    ws.column_dimensions["F"].width=18

    ri=[1]
    def nr(): r=ri[0]; ri[0]+=1; return r

    r=nr(); ws.merge_cells(f"A{r}:D{r}")
    c=ws.cell(r,1,company.upper())
    c.font=F(bold=True,sz=13,color=C_DARK); c.alignment=AL("center"); ws.row_dimensions[r].height=22

    r2=nr(); ws.merge_cells(f"A{r2}:D{r2}")
    c=ws.cell(r2,1,f"PROFIT & LOSS STATEMENT FOR THE YEAR ENDED {_fmt_date(fy_e)}")
    c.font=F(bold=True,sz=10,color=C_DARK); c.alignment=AL("center")
    unit_cell=_unit_box(ws,r2,fy_e,fy_p)

    r3=nr(); ws.merge_cells(f"C{r3}:D{r3}")
    lbl=_lbl_fml(unit_cell)
    ws.cell(r3,3,f'="(Amt. in "&{lbl}&")"').font=F(sz=7.5,color=C_GREY,italic=True)
    ws.cell(r3,3).alignment=AL("right")

    r4=nr(); ws.row_dimensions[r4].height=28
    for col,txt in enumerate(["PARTICULARS","Note No.",
        f"YEAR ENDED\n{_fmt_date(fy_e)}",f"YEAR ENDED\n{_fmt_date(fy_p)}"],1):
        c=ws.cell(r4,col,txt)
        c.font=F(bold=True,sz=9,color=C_WHITE); c.fill=FILL(C_DARK)
        c.alignment=AL("center" if col>1 else "left",wrap=True); c.border=b_all

    # tb_map holds (cy_float, py_float) — compute all values in Python
    DIVISOR = 100_000  # Lakhs

    def v(key, ci):
        if key not in tb_map: return None
        x = _f(tb_map[key][ci])
        return round(x / DIVISOR, 2) if x != 0 else None

    def vsum(keys, ci):
        return round(sum(_f(tb_map[k][ci]) / DIVISOR for k in keys if k in tb_map), 2)

    def _row(lbl,note,key,bold=False,cy_ov=None,py_ov=None):
        r=nr(); ws.row_dimensions[r].height=13
        fill=FILL(C_ALT) if r%2==0 else FILL(C_WHITE)
        c=ws.cell(r,1,lbl)
        c.font=F(bold=bold,sz=8.5,color=C_DARK)
        c.fill=FILL(C_LIGHT) if bold else fill; c.alignment=AL("left"); c.border=b_all
        c=ws.cell(r,2,note or ""); c.font=F(sz=8)
        c.fill=FILL(C_LIGHT) if bold else fill; c.alignment=AL("center"); c.border=b_all
        for col,ci,ov in [(3,0,cy_ov),(4,1,py_ov)]:
            val = ov if ov is not None else v(key, ci)
            cv=ws.cell(r,col, val)
            cv.font=F(bold=bold,sz=8.5,color=C_DARK)
            cv.fill=FILL(C_LIGHT) if bold else fill
            cv.alignment=AL("right"); cv.border=b_all
            if val is not None: cv.number_format='#,##0.00'

    def _trow(lbl, cy_val, py_val):
        r=nr(); ws.row_dimensions[r].height=15
        c=ws.cell(r,1,lbl); c.font=F(bold=True,sz=9,color=C_WHITE)
        c.fill=FILL(C_DARK); c.border=b_all; c.alignment=AL("left")
        ws.cell(r,2,"").font=F(sz=8,color=C_WHITE)
        ws.cell(r,2).fill=FILL(C_DARK); ws.cell(r,2).border=b_all
        ws.cell(r,2).alignment=AL("center")
        for col,val in [(3,cy_val),(4,py_val)]:
            cv=ws.cell(r,col, round(val,2) if val is not None else None)
            cv.font=F(bold=True,sz=9,color=C_WHITE)
            cv.fill=FILL(C_DARK); cv.alignment=AL("right"); cv.border=b_all
            if val is not None: cv.number_format='#,##0.00'

    def _sec(lbl):
        r=nr(); ws.merge_cells(f"A{r}:D{r}")
        c=ws.cell(r,1,lbl); c.font=F(bold=True,sz=8.5,color=C_WHITE)
        c.fill=FILL(C_MID); c.alignment=AL("left"); c.border=b_all; ws.row_dimensions[r].height=13

    def _brow(lbl):
        r=nr(); ws.row_dimensions[r].height=11
        c=ws.cell(r,1,lbl); c.font=F(sz=8.5,color=C_GREY)
        c.fill=FILL(C_WHITE); c.alignment=AL("left"); c.border=b_all
        for col in [2,3,4]:
            ws.cell(r,col).fill=FILL(C_WHITE); ws.cell(r,col).border=b_all

    exp_keys=["pnl_cost_mat","pnl_inv_chg","pnl_emp_exp","pnl_fin_cost","pnl_dep","pnl_other_exp"]
    rev_cy = vsum(["pnl_revenue","pnl_other_inc"], 0)
    rev_py = vsum(["pnl_revenue","pnl_other_inc"], 1)
    exp_cy = vsum(exp_keys, 0)
    exp_py = vsum(exp_keys, 1)
    pbt_cy = round(rev_cy - exp_cy, 2)
    pbt_py = round(rev_py - exp_py, 2)
    np_cy  = round(pbt_cy - (_f(tb_map["pnl_tax"][0])/DIVISOR if "pnl_tax" in tb_map else 0), 2)
    np_py  = round(pbt_py - (_f(tb_map["pnl_tax"][1])/DIVISOR if "pnl_tax" in tb_map else 0), 2)

    _sec("REVENUE")
    _row("I.    REVENUE FROM OPERATIONS",   "15","pnl_revenue")
    _row("II.   OTHER INCOME",              "16","pnl_other_inc")
    _trow("III.  TOTAL REVENUE  (I + II)",  rev_cy, rev_py)

    _sec("EXPENSES")
    _row("IV.   EXPENSES:",                 "",  None)
    _row("          COST OF MATERIAL CONSUMED",          "17","pnl_cost_mat")
    _brow("          PURCHASES OF STOCK IN TRADE")
    _row("          CHANGES IN INVENTORIES OF FG, WIP & STOCK","18","pnl_inv_chg")
    _row("          EMPLOYEE BENEFIT EXPENSE",           "19","pnl_emp_exp")
    _row("          FINANCIAL COST",                     "20","pnl_fin_cost")
    _row("          DEPRECIATION & AMORTIZATION EXPENSE", "8","pnl_dep")
    _row("          OTHER EXPENSES",                     "21","pnl_other_exp")
    _trow("          TOTAL EXPENSES  (IV)",  exp_cy, exp_py)

    _sec("PROFIT / LOSS")
    _trow("V.    PROFIT BEFORE EXCEPTIONAL ITEMS & TAX  (III−IV)", pbt_cy, pbt_py)
    _brow("VI.   EXCEPTIONAL ITEMS")
    _brow("VII.  PRIOR PERIOD ITEMS")
    _trow("VIII. PROFIT BEFORE EXTRAORDINARY ITEMS & TAX  (V−VI)", pbt_cy, pbt_py)
    _brow("IX.   EXTRAORDINARY ITEMS")
    _trow("X.    PROFIT BEFORE TAX  (VIII−IX)", pbt_cy, pbt_py)

    _sec("TAX EXPENSE")
    _brow("          EXCESS / SHORTER PROVISION OF EARLIER YEAR")
    _row("          CURRENT TAX",  "", "pnl_tax")
    _brow("          DEFERRED TAX")
    _brow("          MAT")
    _trow("          TOTAL TAX EXPENSE", v("pnl_tax",0), v("pnl_tax",1))

    _trow("XI.   PROFIT / (LOSS) FOR THE PERIOD  (X − Tax)", np_cy, np_py)
    _brow("XII.  EARNING PER EQUITY SHARE:")
    _brow("          BASIC"); _brow("          DILUTED")

    r_ft=nr(); ws.merge_cells(f"A{r_ft}:D{r_ft}")
    ws.cell(r_ft,1,"See accompanying notes to the financial statements").font=F(sz=7.5,color=C_GREY,italic=True)
    ws.cell(r_ft,1).alignment=AL("left")
    ws.print_area="A1:D70"
    ws.page_setup.orientation="portrait"
    ws.sheet_properties.tabColor=C_ORANGE


# ══════════════════════════════════════════════════════════════════════════════
# MASTER GENERATOR
# ══════════════════════════════════════════════════════════════════════════════
def generate_schedule3_excel(D, D_py, out_path=None):
    if bs_figures is None or pnl_figures is None:
        raise RuntimeError("sch3_classic.py not importable.")
    if D_py and _py_is_same_as_cy and _py_is_same_as_cy(D, D_py):
        D_py=None

    company=D.get("company","Company")
    fy_e=D.get("fy_end","")
    fy_p=(D_py.get("fy_end") if D_py else
          (f"{int(fy_e[:4])-1}{fy_e[4:]}" if len(fy_e)>=4 else ""))

    if out_path is None:
        safe=re.sub(r"[^\w]","_",company[:30])
        out_path=f"Sch3_Classic_{safe}_{fy_e}.xlsx"

    f  = bs_figures(D)
    fp = bs_figures(D_py) if D_py else {k:None for k in f}
    m  = pnl_figures(D)
    mp = pnl_figures(D_py) if D_py else ({k:None for k in m} if m else {})

    wb=openpyxl.Workbook()

    # Sheet 1 — How to Link
    ws_guide=wb.active; ws_guide.title="How to Link"
    _build_guide_sheet(ws_guide, company, fy_e, fy_p)

    # Sheet 2 — TB Data
    ws_tb=wb.create_sheet("TB Data")
    tb=TBDataSheet(ws_tb)
    tb.write_header(company,fy_e,fy_p)
    tb.populate(f,fp,m,mp)
    tb_map=tb.map
    ws_tb.sheet_properties.tabColor="375623"

    # Sheet 3 — Balance Sheet
    ws_bs=wb.create_sheet("Balance Sheet")
    _build_bs_sheet(ws_bs, tb_map, company, fy_e, fy_p)

    # Sheet 4 — P&L Statement
    ws_pnl=wb.create_sheet("P&L Statement")
    _build_pnl_sheet(ws_pnl, tb_map, company, fy_e, fy_p)

    # Open on BS by default
    ws_guide.sheet_properties.tabColor="1D6F42"
    wb.active=ws_bs

    wb.save(out_path)
    print(f"  Schedule III Classic Excel saved → {out_path}")

    return out_path


if __name__=="__main__":
    import argparse,json
    ap=argparse.ArgumentParser()
    ap.add_argument("--json",required=True)
    ap.add_argument("--json-py")
    ap.add_argument("--out")
    args=ap.parse_args()
    if tc is None: raise SystemExit("tally_core.py not found.")
    def _load(p):
        with open(p,encoding="utf-8") as fh: d=json.load(fh)
        return tc.parse_data(d["company"],d["data"],d["fy_start"],d["fy_end"])
    D=_load(args.json)
    D_py=_load(args.json_py) if args.json_py else None
    generate_schedule3_excel(D,D_py,args.out)