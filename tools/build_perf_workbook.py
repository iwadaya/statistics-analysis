"""Generate Performance_Tracker.xlsx (clean) and Performance_Tracker_DEMO.xlsx.

Layout contract with the VBA modules (vba/PerfConfig.bas):
  header row 5, data from row 6 on Treaties / Stats / master sheets;
  Treaties and Stats: A..P = inputs + sync state, Q.. = formulas;
  Stats T = premium USD, W = result USD, X = treaty key (read by the optimiser);
  Shares: header row 8, total row 9, treaties from row 10; H key, I current,
          J scenario (input), K min, L max, M optimised (written by the optimiser).
"""
import random
import sys
from datetime import date

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule, Rule
from openpyxl.styles import Alignment, Font
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles.numbers import NumberFormat
from openpyxl.utils import get_column_letter as L
from openpyxl.comments import Comment

from common import (BOX, CALC_BG, CLASSES, COUNTRIES, CURRENCIES, DATE, FONT, HEADER, INK_SOFT,
                    INPUT_BLUE, MONEY, PCT, PEGS, PEG_DATE, RATE, REGION, STAMP, TEAL, TOTAL_BG,
                    TOTAL_INK, TREATY_TYPES, UNIVERSE_CEDANTS, f, fill, header, name, style_range,
                    title_bar, widths)
from openpyxl.worksheet.datavalidation import DataValidation

FIRST = 6
LAST_TRT, LAST_STAT, LAST_CED, LAST_CTRY, LAST_RATE, LAST_LIST = 405, 6005, 505, 205, 105, 105
SH_HDR, SH_TOT, SH_FIRST, SH_LAST = 8, 9, 10, 409          # Shares sheet
PF_HDR, PF_TOT, PF_TOP, PF_BOT = 6, 7, 8, 407              # Performance sheet
N_YEARS = 20                                               # year columns on Performance / Dashboard
SC_YEARS = 30                                              # year rows on Scenarios
RATIO = '0.0%;[Red]-0.0%;"-"'

TT = [t for t in TREATY_TYPES if t[0] != "ALL"]
CB = [c for c in CLASSES if c[0] != "ALL"]
GROUPS = ["Cedant", "Country", "Class", "Treaty type", "Category", "Treaty"]
MEASURES = ["Premium", "Incurred claims", "Commission + brokerage", "Result", "Loss ratio", "Combined ratio"]
BASES = ["100%", "Current share", "Scenario share", "Optimised share"]
CAT_FILTER = ["All", "Proportional", "Non-proportional"]
CATS = ["PROPORTIONAL", "NON_PROPORTIONAL"]


def rs(col, last=LAST_STAT, sheet="Stats"):
    return f"{sheet}!${col}${FIRST}:${col}${last}"


def tkey(c, d, e, fcol, g, r):
    """Treaty key - must match PerfRepo.TreatyKey: UPPER(cedant|type|class|ccy|ref), blank ref = MAIN."""
    return (f'IF(${c}{r}="","",UPPER(TRIM(${c}{r})&"|"&TRIM(${d}{r})&"|"&TRIM(${e}{r})&"|"&TRIM(${fcol}{r})'
            f'&"|"&IF(TRIM(${g}{r})="","MAIN",TRIM(${g}{r}))))')


# ---- Demo data (fictional) -----------------------------------------------------
DEMO_CEDANTS = [('Demo - Alpha Insurance', 'Saudi Arabia'), ('Demo - Baobab General', 'Zambia'),
                ('Demo - Crescent Takaful', 'United Arab Emirates'), ('Demo - Delta Re-Assurance', 'Kenya'),
                ('Demo - Gulf Pearl Insurance', 'Bahrain')]
# cedant, type, class, ccy, ref, start, end, share, min, max, base premium, mean LR, comm/brok %, growth
DEMO_TREATIES = [
    ('Demo - Alpha Insurance', 'Quota Share', 'Property', 'SAR', 'MAIN', date(2016, 1, 1), date(2025, 12, 31), 0.10, 0.05, 0.20, 180e6, 0.52, 0.32, 0.05),
    ('Demo - Alpha Insurance', 'CAT XL', 'Property', 'SAR', 'LAYER 1', date(2016, 1, 1), date(2025, 12, 31), 0.15, 0.0, 0.25, 40e6, 0.45, 0.10, 0.03),
    ('Demo - Baobab General', 'Quota Share', 'Property', 'USD', 'MAIN', date(2018, 1, 1), date(2025, 12, 31), 0.20, 0.10, 0.30, 9e6, 0.68, 0.30, 0.08),
    ('Demo - Baobab General', 'Risk XL', 'Engineering', 'USD', 'MAIN', date(2019, 1, 1), date(2025, 12, 31), 0.25, 0.0, 0.40, 2.5e6, 0.55, 0.10, 0.04),
    ('Demo - Crescent Takaful', 'Quota Share', 'Medical', 'AED', 'MAIN', date(2016, 1, 1), date(2025, 12, 31), 0.075, 0.0, 0.15, 260e6, 0.78, 0.18, 0.06),
    ('Demo - Crescent Takaful', 'First Surplus', 'Engineering', 'AED', 'MAIN', date(2017, 7, 1), date(2025, 6, 30), 0.10, 0.05, 0.20, 45e6, 0.40, 0.35, 0.04),
    ('Demo - Delta Re-Assurance', 'Quota Share', 'Motor', 'USD', 'MAIN', date(2016, 1, 1), date(2025, 12, 31), 0.15, 0.0, 0.25, 22e6, 0.71, 0.27, 0.07),
    ('Demo - Gulf Pearl Insurance', 'Quota Share', 'Property', 'BHD', 'MAIN', date(2016, 1, 1), date(2025, 12, 31), 0.125, 0.05, 0.25, 14e6, 0.47, 0.33, 0.03),
    ('Demo - Gulf Pearl Insurance', 'CAT XL', 'Property', 'BHD', 'MAIN', date(2016, 1, 1), date(2025, 12, 31), 0.20, 0.10, 0.30, 3.2e6, 0.35, 0.10, 0.02),
]


def period_years(start, end):
    n = round(((end - start).days + 1) / 365.25)
    return max(1, n)


def demo_stats():
    """Deterministic fictional statistics: one row per UW year of each demo treaty."""
    rnd = random.Random(20261119)
    rows = []
    for (ced, tt, cob, ccy, ref, st, en, share, mn, mx, base, lr, cost, g) in DEMO_TREATIES:
        np_ = dict(TT)[tt] == "NON_PROPORTIONAL"
        for k in range(period_years(st, en)):
            y = st.year + k
            prem = round(base * (1 + g) ** k * rnd.uniform(0.92, 1.08), -3)
            if np_:
                # mostly quiet years with the odd large loss
                shock = rnd.random()
                f_lr = 0.05 + rnd.uniform(0, 0.25) if shock < 0.7 else rnd.uniform(0.9, 2.4)
                claims = round(prem * f_lr * lr / 0.45, -3)
                rows.append((ced, tt, cob, ccy, ref, y, prem, claims, None, round(prem * cost, -3)))
            else:
                claims = round(prem * max(0.05, rnd.gauss(lr, 0.12)), -3)
                rows.append((ced, tt, cob, ccy, ref, y, prem, claims, round(prem * cost, -3), None))
    return rows


def dv(ws, ref, **kw):
    d = DataValidation(allow_blank=True, showErrorMessage=True, **kw)
    ws.add_data_validation(d)
    d.add(ref)
    return d


def selector(ws, label_cell, value_cell, label, value, formula1, width_note=None):
    ws[label_cell] = label
    ws[label_cell].font = f(10, True)
    ws[label_cell].alignment = Alignment(horizontal="right")
    ws[value_cell] = value
    c = ws[value_cell]
    c.font = f(11, True, INPUT_BLUE)
    c.fill = fill("FFFF00")
    c.border = BOX
    c.alignment = Alignment(horizontal="center")
    if formula1:
        dv(ws, value_cell, type="list", formula1=formula1)


def ratio_format_rule(ws, ref, flag="sel_IsRatio"):
    dxf = DifferentialStyle(numFmt=NumberFormat(numFmtId=170, formatCode='0.0%'))
    ws.conditional_formatting.add(ref, Rule(type="expression", dxf=dxf, formula=[flag]))


def build(path, demo=False):
    wb = Workbook()
    st = wb.active
    st.title = "Start"
    dash = wb.create_sheet("Dashboard")
    pf = wb.create_sheet("Performance")
    shr = wb.create_sheet("Shares")
    sc = wb.create_sheet("Scenarios")
    tr = wb.create_sheet("Treaties")
    sx = wb.create_sheet("Stats")
    ce = wb.create_sheet("Cedants")
    co = wb.create_sheet("Countries")
    ra = wb.create_sheet("Rates")
    li = wb.create_sheet("Lists")
    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.tabColor = HEADER if ws.title in ("Start", "Dashboard", "Performance", "Shares", "Scenarios") else TEAL

    # ---------------- defined names ---------------------------------------------
    name(wb, "lst_Cedants", f"OFFSET(Cedants!$A${FIRST},0,0,MAX(1,COUNTA(Cedants!$A${FIRST}:$A${LAST_CED})),1)")
    name(wb, "lst_Countries", f"OFFSET(Countries!$A${FIRST},0,0,MAX(1,COUNTA(Countries!$A${FIRST}:$A${LAST_CTRY})),1)")
    name(wb, "lst_Ccy", f"OFFSET(Rates!$A${FIRST},0,0,MAX(1,COUNTA(Rates!$A${FIRST}:$A${LAST_RATE})),1)")
    name(wb, "lst_Treaty", f"OFFSET(Lists!$A${FIRST},0,0,MAX(1,COUNTA(Lists!$A${FIRST}:$A${LAST_LIST})),1)")
    name(wb, "lst_Cob", f"OFFSET(Lists!$D${FIRST},0,0,MAX(1,COUNTA(Lists!$D${FIRST}:$D${LAST_LIST})),1)")
    name(wb, "lst_Groups", f"Lists!$G$6:$G${5 + len(GROUPS)}")
    name(wb, "lst_Measures", f"Lists!$I$6:$I${5 + len(MEASURES)}")
    name(wb, "lst_Bases", f"Lists!$K$6:$K${5 + len(BASES)}")
    name(wb, "lst_CatFilter", f"Lists!$M$6:$M${5 + len(CAT_FILTER)}")
    name(wb, "cfg_DbPath", "Start!$C$7")
    name(wb, "cfg_LastPull", "Start!$C$9")
    name(wb, "cfg_LastPush", "Start!$C$10")
    name(wb, "sel_GroupIdx", f"MATCH(Performance!$B$4,Lists!$G$6:$G${5 + len(GROUPS)},0)")
    name(wb, "sel_MeasureIdx", f"MATCH(Performance!$D$4,Lists!$I$6:$I${5 + len(MEASURES)},0)")
    name(wb, "sel_IsRatio", f"MATCH(Performance!$D$4,Lists!$I$6:$I${5 + len(MEASURES)},0)>=5")
    name(wb, "sel_BasisIdx", f"MATCH(Performance!$F$4,Lists!$K$6:$K${5 + len(BASES)},0)")
    name(wb, "crit_Cat", f'CHOOSE(MATCH(Performance!$H$4,Lists!$M$6:$M${5 + len(CAT_FILTER)},0),"*","PROPORTIONAL","NON_PROPORTIONAL")')
    yrs = rs("H")
    name(wb, "sel_FirstYear", f'IF(Performance!$J$4="",IF(COUNT({yrs})=0,YEAR(TODAY())-{N_YEARS - 1},MIN({yrs})),Performance!$J$4)')
    name(wb, "opt_FromYear", "Shares!$C$4")
    name(wb, "opt_ToYear", "Shares!$E$4")
    name(wb, "opt_Lambda", "Shares!$C$5")
    name(wb, "opt_PremMin", "Shares!$C$6")
    name(wb, "opt_PremMax", "Shares!$E$6")
    name(wb, "opt_Status", "Shares!$C$7")
    name(wb, "shr_From", f'IF(Shares!$C$4="",IF(COUNT({yrs})=0,YEAR(TODAY()),MIN({yrs})),Shares!$C$4)')
    name(wb, "shr_To", f'IF(Shares!$E$4="",IF(COUNT({yrs})=0,YEAR(TODAY()),MAX({yrs})),Shares!$E$4)')

    blue = f(10, color=INPUT_BLUE)
    grey = f(9, color=INK_SOFT)

    def status_cf(ws, last):
        ref = f"A{FIRST}:A{last}"
        ws.conditional_formatting.add(ref, FormulaRule(formula=[f'LEFT($A{FIRST},5)="ERROR"'],
                                      fill=fill("F8D7DA"), font=Font(name=FONT, color="9C0006", bold=True)))
        ws.conditional_formatting.add(ref, FormulaRule(formula=[f'OR(LEFT($A{FIRST},3)="NOT",LEFT($A{FIRST},4)="SKIP",LEFT($A{FIRST},5)="EMPTY")'],
                                      fill=fill("FFF3CD"), font=Font(name=FONT, color="7A5B00")))
        ws.conditional_formatting.add(ref, FormulaRule(formula=[f'$A{FIRST}="OK"'], font=Font(name=FONT, color="2F8F6B")))

    # ================= Treaties =================================================
    title_bar(tr, 22, "Treaties",
              "One row per treaty (cedant x treaty type x class x currency x reference). The stats start and end dates set "
              "how many UW-year rows the treaty needs. Shares: current, and the min / max the optimiser may use.")
    labels = ["Status", "ID", "Cedant", "Treaty type", "Class of business", "Currency", "Reference",
              "Stats start date", "Stats end date", "Our share %", "Min share %", "Max share %", "Notes",
              "Delete (X)", "row_version", "_fp", "Country", "Category", "UW years", "UW year range",
              "Treaty key", "Rows on Stats"]
    header(tr, 5, labels)
    for c in range(17, 23):
        tr.cell(5, c).fill = fill(TEAL)
    tr.cell(4, 3, "Inputs (blue). Blank reference = MAIN; use LAYER 1, LAYER 2 ... for XL layers").font = f(8, True, INK_SOFT)
    tr.cell(4, 17, "Calculated").font = f(8, True, INK_SOFT)
    widths(tr, {"A": 24, "B": 6, "C": 28, "D": 18, "E": 18, "F": 9, "G": 11, "H": 13, "I": 13, "J": 10,
                "K": 10, "L": 10, "M": 22, "N": 8, "O": 6, "P": 6, "Q": 18, "R": 18, "S": 8, "T": 12,
                "U": 40, "V": 9})
    tr.column_dimensions["O"].hidden = True
    tr.column_dimensions["P"].hidden = True
    style_range(tr, f"A{FIRST}:A{LAST_TRT}", font=grey, border=False)
    style_range(tr, f"B{FIRST}:B{LAST_TRT}", font=grey, number_format="0", border=False)
    style_range(tr, f"C{FIRST}:N{LAST_TRT}", font=blue, border=False)
    style_range(tr, f"H{FIRST}:I{LAST_TRT}", number_format=DATE, border=False)
    style_range(tr, f"J{FIRST}:L{LAST_TRT}", number_format=PCT, border=False)
    style_range(tr, f"N{FIRST}:N{LAST_TRT}", align=Alignment(horizontal="center"), border=False)
    for r in range(FIRST, LAST_TRT + 1):
        tr[f"Q{r}"] = f'=IF($C{r}="","",IFERROR(INDEX(Cedants!$B${FIRST}:$B${LAST_CED},MATCH($C{r},Cedants!$A${FIRST}:$A${LAST_CED},0)),"?"))'
        tr[f"R{r}"] = f'=IF($D{r}="","",IFERROR(INDEX(Lists!$B${FIRST}:$B${LAST_LIST},MATCH($D{r},Lists!$A${FIRST}:$A${LAST_LIST},0)),"?"))'
        tr[f"S{r}"] = f'=IF(OR($H{r}="",$I{r}=""),"",IF($I{r}<=$H{r},"",MAX(1,ROUND(($I{r}-$H{r}+1)/365.25,0))))'
        tr[f"T{r}"] = f'=IF($S{r}="","",YEAR($H{r})&IF($S{r}>1,"-"&(YEAR($H{r})+$S{r}-1),""))'
        tr[f"U{r}"] = "=" + tkey("C", "D", "E", "F", "G", r)
        tr[f"V{r}"] = f'=IF($U{r}="","",COUNTIF({rs("X")},$U{r}))'
    style_range(tr, f"Q{FIRST}:V{LAST_TRT}", font=f(9), bg=CALC_BG, border=False)
    style_range(tr, f"S{FIRST}:S{LAST_TRT}", number_format="0", align=Alignment(horizontal="center"), border=False)
    style_range(tr, f"V{FIRST}:V{LAST_TRT}", number_format="0", align=Alignment(horizontal="center"), border=False)
    style_range(tr, f"U{FIRST}:U{LAST_TRT}", font=f(8, color=INK_SOFT), border=False)
    status_cf(tr, LAST_TRT)
    tr.conditional_formatting.add(f"V{FIRST}:V{LAST_TRT}", FormulaRule(
        formula=[f'AND($S{FIRST}<>"",$V{FIRST}<$S{FIRST})'], fill=fill("FFF3CD"), font=Font(name=FONT, bold=True, color="7A5B00")))
    tr.conditional_formatting.add(f"C{FIRST}:V{LAST_TRT}", FormulaRule(formula=[f'UPPER($N{FIRST})="X"'],
                                  font=Font(name=FONT, strike=True, color="999999")))
    dv(tr, f"C{FIRST}:C{LAST_TRT}", type="list", formula1="lst_Cedants", errorTitle="Cedant",
       error="Pick a cedant from the Cedants sheet (Add cedant to create one).")
    dv(tr, f"D{FIRST}:D{LAST_TRT}", type="list", formula1="lst_Treaty")
    dv(tr, f"E{FIRST}:E{LAST_TRT}", type="list", formula1="lst_Cob")
    dv(tr, f"F{FIRST}:F{LAST_TRT}", type="list", formula1="lst_Ccy")
    dv(tr, f"H{FIRST}:I{LAST_TRT}", type="date", operator="greaterThan", formula1="29221", errorTitle="Date",
       error="Enter a date, e.g. 01/01/2016.")
    dv(tr, f"J{FIRST}:L{LAST_TRT}", type="decimal", operator="between", formula1="0", formula2="1",
       errorTitle="Share", error="Enter a percentage between 0% and 100%.")
    dv(tr, f"N{FIRST}:N{LAST_TRT}", type="list", formula1='"X"')
    tr.freeze_panes = "D6"
    tr.auto_filter.ref = f"A5:V{LAST_TRT}"
    if demo:
        for i, t in enumerate(DEMO_TREATIES):
            r = FIRST + i
            ced, tt, cob, ccy, ref, s0, e0, share, mn, mx = t[:10]
            for col, v in zip("CDEFGHIJKL", (ced, tt, cob, ccy, ref, s0, e0, share, mn, mx)):
                tr[f"{col}{r}"] = v
            tr[f"A{r}"] = "NEW - not yet pushed"

    # ================= Stats ====================================================
    title_bar(sx, 32, "Statistics by UW year (100%)",
              "One row per treaty per UW year, at 100% of the cedant's figures in original currency. Use Generate stat rows "
              "to create the rows from each treaty's dates. Proportional: Commission. Non-proportional: Brokerage.")
    labels = ["Status", "ID", "Cedant", "Treaty type", "Class of business", "Currency", "Reference", "UW year",
              "Premium", "Incurred claims", "Commission", "Brokerage", "Notes", "Delete (X)", "row_version", "_fp",
              "Country", "Category", "FX (USD per 1)", "Premium USD", "Claims USD", "Comm + brok USD", "Result USD",
              "Treaty key", "_group", "_idx", "_cur", "_scen", "_opt", "_basis", "_num", "_den"]
    header(sx, 5, labels)
    for c in range(17, 33):
        sx.cell(5, c).fill = fill(TEAL)
    sx.cell(4, 3, "Inputs (blue) - 100% figures").font = f(8, True, INK_SOFT)
    sx.cell(4, 17, "Calculated - do not type here").font = f(8, True, INK_SOFT)
    w = {"A": 22, "B": 6, "C": 28, "D": 16, "E": 16, "F": 8, "G": 10, "H": 8, "I": 15, "J": 15, "K": 14,
         "L": 14, "M": 18, "N": 8, "O": 6, "P": 6, "Q": 16, "R": 18, "S": 11, "T": 15, "U": 15, "V": 15,
         "W": 15, "X": 40}
    widths(sx, w)
    for col in ["O", "P"] + [L(c) for c in range(25, 33)]:
        sx.column_dimensions[col].hidden = True
    style_range(sx, f"A{FIRST}:A{LAST_STAT}", font=grey, border=False)
    style_range(sx, f"B{FIRST}:B{LAST_STAT}", font=grey, number_format="0", border=False)
    style_range(sx, f"C{FIRST}:N{LAST_STAT}", font=blue, border=False)
    style_range(sx, f"H{FIRST}:H{LAST_STAT}", number_format="0", border=False)
    style_range(sx, f"I{FIRST}:L{LAST_STAT}", number_format=MONEY, border=False)
    style_range(sx, f"N{FIRST}:N{LAST_STAT}", align=Alignment(horizontal="center"), border=False)
    rate_rng = f"Rates!$A${FIRST}:$C${LAST_RATE}"
    for r in range(FIRST, LAST_STAT + 1):
        sx[f"Q{r}"] = f'=IF($C{r}="","",IFERROR(INDEX(Cedants!$B${FIRST}:$B${LAST_CED},MATCH($C{r},Cedants!$A${FIRST}:$A${LAST_CED},0)),""))'
        sx[f"R{r}"] = f'=IF($D{r}="","",IFERROR(INDEX(Lists!$B${FIRST}:$B${LAST_LIST},MATCH($D{r},Lists!$A${FIRST}:$A${LAST_LIST},0)),""))'
        sx[f"S{r}"] = (f'=IF(OR($C{r}="",$F{r}="",UPPER($N{r})="X"),"",IFERROR(IF(VLOOKUP($F{r},{rate_rng},3,FALSE)>0,'
                       f'VLOOKUP($F{r},{rate_rng},3,FALSE),""),""))')
        sx[f"T{r}"] = f'=IF($S{r}="",0,N($I{r})*$S{r})'
        sx[f"U{r}"] = f'=IF($S{r}="",0,N($J{r})*$S{r})'
        sx[f"V{r}"] = f'=IF($S{r}="",0,(N($K{r})+N($L{r}))*$S{r})'
        sx[f"W{r}"] = f'=T{r}-U{r}-V{r}'
        sx[f"X{r}"] = "=" + tkey("C", "D", "E", "F", "G", r)
        sx[f"Y{r}"] = f'=IF($X{r}="","",CHOOSE(sel_GroupIdx,$C{r},$Q{r},$E{r},$D{r},$R{r},$X{r}))'
        sx[f"Z{r}"] = f'=IF($X{r}="","",IFERROR(MATCH($X{r},Shares!$H${SH_FIRST}:$H${SH_LAST},0),""))'
        sx[f"AA{r}"] = f'=IF($Z{r}="",0,N(INDEX(Shares!$I${SH_FIRST}:$I${SH_LAST},$Z{r})))'
        sx[f"AB{r}"] = f'=IF($Z{r}="",0,N(INDEX(Shares!$V${SH_FIRST}:$V${SH_LAST},$Z{r})))'
        sx[f"AC{r}"] = f'=IF($Z{r}="",0,N(INDEX(Shares!$M${SH_FIRST}:$M${SH_LAST},$Z{r})))'
        sx[f"AD{r}"] = f'=CHOOSE(sel_BasisIdx,1,$AA{r},$AB{r},$AC{r})'
        sx[f"AE{r}"] = f'=CHOOSE(sel_MeasureIdx,$T{r},$U{r},$V{r},$W{r},$U{r},$U{r}+$V{r})*$AD{r}'
        sx[f"AF{r}"] = f'=$T{r}*$AD{r}'
    style_range(sx, f"Q{FIRST}:R{LAST_STAT}", font=f(9), bg=CALC_BG, border=False)
    style_range(sx, f"S{FIRST}:S{LAST_STAT}", font=f(9), number_format=RATE, bg=CALC_BG, border=False)
    style_range(sx, f"T{FIRST}:W{LAST_STAT}", font=f(10), number_format=MONEY, bg=CALC_BG, border=False)
    style_range(sx, f"X{FIRST}:X{LAST_STAT}", font=f(8, color=INK_SOFT), bg=CALC_BG, border=False)
    status_cf(sx, LAST_STAT)
    sx.conditional_formatting.add(f"C{FIRST}:X{LAST_STAT}", FormulaRule(formula=[f'UPPER($N{FIRST})="X"'],
                                  font=Font(name=FONT, strike=True, color="999999")))
    sx.conditional_formatting.add(f"S{FIRST}:S{LAST_STAT}", FormulaRule(
        formula=[f'AND($C{FIRST}<>"",$F{FIRST}<>"",$S{FIRST}="",UPPER($N{FIRST})<>"X")'], fill=fill("FFF3CD")))
    # commission on an NP row / brokerage on a proportional row
    sx.conditional_formatting.add(f"K{FIRST}:K{LAST_STAT}", FormulaRule(
        formula=[f'AND($R{FIRST}="NON_PROPORTIONAL",N($K{FIRST})<>0)'], fill=fill("F8D7DA")))
    sx.conditional_formatting.add(f"L{FIRST}:L{LAST_STAT}", FormulaRule(
        formula=[f'AND($R{FIRST}="PROPORTIONAL",N($L{FIRST})<>0)'], fill=fill("F8D7DA")))
    sx.conditional_formatting.add(f"K{FIRST}:K{LAST_STAT}", FormulaRule(
        formula=[f'$R{FIRST}="NON_PROPORTIONAL"'], fill=fill("EEEEEE")))
    sx.conditional_formatting.add(f"L{FIRST}:L{LAST_STAT}", FormulaRule(
        formula=[f'$R{FIRST}="PROPORTIONAL"'], fill=fill("EEEEEE")))
    dv(sx, f"C{FIRST}:C{LAST_STAT}", type="list", formula1="lst_Cedants")
    dv(sx, f"D{FIRST}:D{LAST_STAT}", type="list", formula1="lst_Treaty")
    dv(sx, f"E{FIRST}:E{LAST_STAT}", type="list", formula1="lst_Cob")
    dv(sx, f"F{FIRST}:F{LAST_STAT}", type="list", formula1="lst_Ccy")
    dv(sx, f"H{FIRST}:H{LAST_STAT}", type="whole", operator="between", formula1="1980", formula2="2100",
       errorTitle="UW year", error="Enter an underwriting year, e.g. 2024.")
    dv(sx, f"I{FIRST}:L{LAST_STAT}", type="decimal", operator="between", formula1="-1E+15", formula2="1E+15",
       errorTitle="Amount", error="Enter a number.")
    dv(sx, f"N{FIRST}:N{LAST_STAT}", type="list", formula1='"X"')
    sx.freeze_panes = "I6"
    sx.auto_filter.ref = f"A5:X{LAST_STAT}"
    if demo:
        for i, row in enumerate(demo_stats()):
            r = FIRST + i
            for col, v in zip("CDEFGHIJKL", row):
                if v is not None:
                    sx[f"{col}{r}"] = v
            sx[f"A{r}"] = "NEW - not yet pushed"

    # ================= Shares ===================================================
    title_bar(shr, 22, "Shares - what-if and optimiser",
              "One row per treaty (follows Treaties). Type scenario shares in column J (blank = current share), or click "
              "Run optimiser to fill column M. Figures are totals over the UW years in C4:E4, in USD.")
    selector(shr, "B4", "C4", "From UW year", None, None)
    selector(shr, "D4", "E4", "To UW year", None, None)
    shr["F4"] = '="Using UW "&shr_From&" - "&shr_To&"  (blank = all years with stats)"'
    shr["F4"].font = f(9, True, TEAL)
    selector(shr, "B5", "C5", "Risk weight (lambda)", 0, None)
    shr["D5"] = ("0 = maximise total result. Higher = prefer a steadier result year to year "
                 "(objective = average annual result - lambda x its standard deviation).")
    shr["D5"].font = f(9, italic=True, color=INK_SOFT)
    selector(shr, "B6", "C6", "Avg annual premium min (USD)", None, None)
    selector(shr, "D6", "E6", "max (USD)", None, None)
    shr["F6"] = "Optional: keep the portfolio's average annual premium at our share inside this range."
    shr["F6"].font = f(9, italic=True, color=INK_SOFT)
    for c in ("C6", "E6"):
        shr[c].number_format = MONEY
    shr["C5"].number_format = "0.00"
    for c in ("C4", "E4"):
        shr[c].number_format = "0"
    dv(shr, "C4", type="whole", operator="between", formula1="1980", formula2="2100")
    dv(shr, "E4", type="whole", operator="between", formula1="1980", formula2="2100")
    dv(shr, "C5", type="decimal", operator="greaterThanOrEqual", formula1="0", errorTitle="Risk weight",
       error="Enter 0 or more.")
    shr["B7"] = "Last optimiser run"; shr["B7"].font = f(10, True); shr["B7"].alignment = Alignment(horizontal="right")
    shr.merge_cells("C7:L7")
    shr["C7"].font = f(9, color=INK_SOFT)
    labels = ["Cedant", "Country", "Treaty type", "Category", "Class", "Ccy", "Reference", "Treaty key",
              "Current share", "Scenario share", "Min share", "Max share", "Optimised share",
              "Premium 100%", "Result 100%", "Loss ratio", "Combined ratio",
              "Result @ current", "Result @ scenario", "Result @ optimised", "Optimised vs current", "Scenario used"]
    header(shr, SH_HDR, labels)
    for c in (10,):
        shr.cell(SH_HDR, c).fill = fill("2F6F8F")
    shr.cell(SH_HDR, 13).fill = fill("8A5A00")
    for c in range(14, 23):
        shr.cell(SH_HDR, c).fill = fill(TEAL)
    shr.row_dimensions[SH_HDR].height = 32
    widths(shr, {"A": 26, "B": 14, "C": 15, "D": 17, "E": 14, "F": 6, "G": 9, "H": 34, "I": 10, "J": 10,
                 "K": 9, "L": 9, "M": 11, "N": 15, "O": 15, "P": 9, "Q": 10, "R": 15, "S": 15, "T": 15,
                 "U": 15, "V": 9})
    trs = lambda col, r: f"Treaties!{col}{r - (SH_FIRST - FIRST)}"
    yr_crit = f'{rs("H")},">="&shr_From,{rs("H")},"<="&shr_To'
    for r in range(SH_FIRST, SH_LAST + 1):
        shr[f"A{r}"] = f'=IF({trs("C", r)}="","",{trs("C", r)})'
        shr[f"B{r}"] = f'=IF($A{r}="","",{trs("Q", r)})'
        shr[f"C{r}"] = f'=IF($A{r}="","",{trs("D", r)}&"")'
        shr[f"D{r}"] = f'=IF($A{r}="","",{trs("R", r)})'
        shr[f"E{r}"] = f'=IF($A{r}="","",{trs("E", r)}&"")'
        shr[f"F{r}"] = f'=IF($A{r}="","",{trs("F", r)}&"")'
        shr[f"G{r}"] = f'=IF($A{r}="","",IF({trs("G", r)}="","MAIN",UPPER({trs("G", r)})))'
        shr[f"H{r}"] = f'={trs("U", r)}&""'
        shr[f"I{r}"] = f'=IF($A{r}="","",N({trs("J", r)}))'
        shr[f"K{r}"] = f'=IF($A{r}="","",IF({trs("K", r)}="",0,{trs("K", r)}))'
        shr[f"L{r}"] = f'=IF($A{r}="","",IF({trs("L", r)}="",MIN(1,2*$I{r}),{trs("L", r)}))'
        shr[f"N{r}"] = f'=IF($A{r}="","",SUMIFS({rs("T")},{rs("X")},$H{r},{yr_crit}))'
        shr[f"O{r}"] = f'=IF($A{r}="","",SUMIFS({rs("W")},{rs("X")},$H{r},{yr_crit}))'
        shr[f"P{r}"] = f'=IF(N($N{r})=0,"",SUMIFS({rs("U")},{rs("X")},$H{r},{yr_crit})/$N{r})'
        shr[f"Q{r}"] = f'=IF(N($N{r})=0,"",1-$O{r}/$N{r})'
        shr[f"R{r}"] = f'=IF($A{r}="","",$O{r}*$I{r})'
        shr[f"S{r}"] = f'=IF($A{r}="","",$O{r}*$V{r})'
        shr[f"T{r}"] = f'=IF(OR($A{r}="",$M{r}=""),"",$O{r}*$M{r})'
        shr[f"U{r}"] = f'=IF(OR($A{r}="",$M{r}=""),"",$T{r}-$R{r})'
        shr[f"V{r}"] = f'=IF($A{r}="","",IF($J{r}="",$I{r},$J{r}))'
    t = SH_TOT
    shr[f"A{t}"] = "TOTAL"
    for col in "NORSTU":
        shr[f"{col}{t}"] = f"=SUM({col}{SH_FIRST}:{col}{SH_LAST})"
    shr[f"P{t}"] = (f'=IF(N($N{t})=0,"",SUMIFS({rs("U")},{yr_crit})/SUMIFS({rs("T")},{yr_crit}))')
    shr[f"Q{t}"] = f'=IF(N($N{t})=0,"",1-$O{t}/$N{t})'
    style_range(shr, f"A{t}:V{t}", font=f(10, True, TOTAL_INK), bg=TOTAL_BG)
    style_range(shr, f"A{SH_FIRST}:H{SH_LAST}", font=f(10), border=False)
    style_range(shr, f"H{SH_FIRST}:H{SH_LAST}", font=f(8, color=INK_SOFT), border=False)
    style_range(shr, f"I{SH_FIRST}:M{SH_LAST}", font=f(10), number_format=PCT, border=False)
    style_range(shr, f"J{SH_FIRST}:J{SH_LAST}", font=f(10, True, INPUT_BLUE), number_format=PCT, bg="FFFDE7", border=False)
    style_range(shr, f"M{SH_FIRST}:M{SH_LAST}", font=f(10, True, "8A5A00"), number_format=PCT, border=False)
    style_range(shr, f"K{SH_FIRST}:L{SH_LAST}", font=f(9, color=INK_SOFT), number_format=PCT, border=False)
    style_range(shr, f"V{SH_FIRST}:V{SH_LAST}", font=f(9, color=INK_SOFT), number_format=PCT, border=False)
    for col in "NORSTU":
        style_range(shr, f"{col}{SH_FIRST}:{col}{SH_LAST}", font=f(10), number_format=MONEY, border=False)
        shr[f"{col}{t}"].number_format = MONEY
    for col in "PQ":
        style_range(shr, f"{col}{SH_FIRST}:{col}{SH_LAST}", font=f(10), number_format=RATIO, border=False)
        shr[f"{col}{t}"].number_format = RATIO
    dv(shr, f"J{SH_FIRST}:J{SH_LAST}", type="decimal", operator="between", formula1="0", formula2="1",
       errorTitle="Scenario share", error="Enter a percentage between 0% and 100%, or leave blank for the current share.")
    shr.conditional_formatting.add(f"Q{SH_FIRST}:Q{SH_LAST}", CellIsRule(operator="greaterThan", formula=["1"],
                                   font=Font(name=FONT, bold=True, color="9C0006")))
    shr.conditional_formatting.add(f"M{SH_FIRST}:M{SH_LAST}", FormulaRule(
        formula=[f'AND($M{SH_FIRST}<>"",$M{SH_FIRST}>$I{SH_FIRST}+0.00005)'], fill=fill("D6EFE3")))
    shr.conditional_formatting.add(f"M{SH_FIRST}:M{SH_LAST}", FormulaRule(
        formula=[f'AND($M{SH_FIRST}<>"",$M{SH_FIRST}<$I{SH_FIRST}-0.00005)'], fill=fill("FBE3E4")))
    shr.conditional_formatting.add(f"L{SH_FIRST}:L{SH_LAST}", FormulaRule(
        formula=[f'AND($A{SH_FIRST}<>"",$L{SH_FIRST}=0)'], fill=fill("FFF3CD")))
    shr.freeze_panes = f"B{SH_FIRST}"
    shr.auto_filter.ref = f"A{SH_HDR}:V{SH_LAST}"

    # ================= Scenarios ==================================================
    title_bar(sc, 16, "Scenarios - portfolio by UW year (USD)",
              "Our share of the whole portfolio under current, scenario and optimised shares, over the years set on Shares.")
    sc["A4"] = '="UW "&shr_From&" - "&shr_To&"   |   risk weight "&TEXT(opt_Lambda,"0.00")'
    sc["A4"].font = f(10, True, TEAL)
    grp = [("Current shares", "AA", 2), ("Scenario shares", "AB", 7), ("Optimised shares", "AC", 12)]
    for label, _, c0 in grp:
        sc.merge_cells(start_row=5, start_column=c0, end_row=5, end_column=c0 + 4)
        cell = sc.cell(5, c0, label)
        cell.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
        cell.fill = fill({"Current shares": TEAL, "Scenario shares": "2F6F8F", "Optimised shares": "8A5A00"}[label])
        cell.alignment = Alignment(horizontal="center")
    header(sc, 6, ["UW year"] + ["Premium", "Claims", "Result", "Loss ratio", "Combined ratio"] * 3)
    widths(sc, {"A": 26, **{L(c): 14 for c in range(2, 17)}})
    S0, S1 = 7, 7 + SC_YEARS - 1
    H, T_, U_, W_ = rs("H"), rs("T"), rs("U"), rs("W")
    for i in range(SC_YEARS):
        r = S0 + i
        sc[f"A{r}"] = f'=IF(shr_From+{i}>shr_To,"",shr_From+{i})'
        for _, sh_col, c0 in grp:
            share = rs(sh_col)
            p, cl, res, lr, cr = (L(c0 + k) for k in range(5))
            sc[f"{p}{r}"] = f'=IF($A{r}="","",SUMPRODUCT(({H}=$A{r})*{T_}*{share}))'
            sc[f"{cl}{r}"] = f'=IF($A{r}="","",SUMPRODUCT(({H}=$A{r})*{U_}*{share}))'
            sc[f"{res}{r}"] = f'=IF($A{r}="","",SUMPRODUCT(({H}=$A{r})*{W_}*{share}))'
            sc[f"{lr}{r}"] = f'=IF(N({p}{r})=0,"",{cl}{r}/{p}{r})'
            sc[f"{cr}{r}"] = f'=IF(N({p}{r})=0,"",1-{res}{r}/{p}{r})'
    style_range(sc, f"A{S0}:A{S1}", font=f(10, True), align=Alignment(horizontal="center"), border=False)
    for _, _, c0 in grp:
        style_range(sc, f"{L(c0)}{S0}:{L(c0 + 2)}{S1}", font=f(10), number_format=MONEY, border=False)
        style_range(sc, f"{L(c0 + 3)}{S0}:{L(c0 + 4)}{S1}", font=f(10), number_format=RATIO, border=False)
    sc.conditional_formatting.add(f"B{S0}:P{S1}", FormulaRule(formula=["MOD(ROW(),2)=0"], fill=fill("F4F9FA")))
    sm = S1 + 2
    summary = [("Total", "SUM"), ("Average annual result", "AVG"), ("Std deviation of annual result", "SD"),
               ("Objective (average - lambda x std dev)", "OBJ"), ("Average annual premium", "AVGP"),
               ("Worst year result", "MIN"), ("Years with a loss", "LOSS"), ("Combined ratio, all years", "CR")]
    sc.cell(sm - 1, 1, "Summary").font = f(12, True, TEAL)
    for k, (lab, kind) in enumerate(summary):
        r = sm + k
        sc.cell(r, 1, lab).font = f(10, True)
        for _, _, c0 in grp:
            p, cl, res = L(c0), L(c0 + 1), L(c0 + 2)
            pr, rr = f"{p}{S0}:{p}{S1}", f"{res}{S0}:{res}{S1}"
            col = res
            if kind == "SUM":
                sc[f"{p}{r}"] = f"=SUM({pr})"; sc[f"{cl}{r}"] = f"=SUM({cl}{S0}:{cl}{S1})"; sc[f"{res}{r}"] = f"=SUM({rr})"
                sc[f"{L(c0 + 3)}{r}"] = f'=IF(N({p}{r})=0,"",{cl}{r}/{p}{r})'
                sc[f"{L(c0 + 4)}{r}"] = f'=IF(N({p}{r})=0,"",1-{res}{r}/{p}{r})'
            elif kind == "AVG":
                sc[f"{col}{r}"] = f'=IFERROR(AVERAGE({rr}),"")'
            elif kind == "SD":
                sc[f"{col}{r}"] = f'=IFERROR(STDEVP({rr}),"")'
            elif kind == "OBJ":
                sc[f"{col}{r}"] = f'=IFERROR(AVERAGE({rr})-N(opt_Lambda)*STDEVP({rr}),"")'
            elif kind == "AVGP":
                sc[f"{p}{r}"] = f'=IFERROR(AVERAGE({pr}),"")'
                col = p
            elif kind == "MIN":
                sc[f"{col}{r}"] = f'=IF(COUNT({rr})=0,"",MIN({rr}))'
            elif kind == "LOSS":
                sc[f"{col}{r}"] = f'=COUNTIF({rr},"<0")'
            elif kind == "CR":
                sc[f"{L(c0 + 4)}{r}"] = f'=IF(N({p}{sm})=0,"",1-{res}{sm}/{p}{sm})'
            for cc in range(c0, c0 + 5):
                cell = sc.cell(r, cc)
                cell.font = f(10, True, TOTAL_INK)
                cell.fill = fill(TOTAL_BG)
                cell.number_format = RATIO if (cc - c0) >= 3 else MONEY
            if kind == "LOSS":
                sc[f"{col}{r}"].number_format = "0"
        sc.cell(r, 1).fill = fill(TOTAL_BG)
    sc.freeze_panes = "B7"

    def line_chart(title, cols, anchor, ytitle, fmt):
        ch = LineChart()
        ch.title = title
        ch.height, ch.width = 7.5, 17
        ch.y_axis.title = ytitle
        ch.y_axis.numFmt = fmt
        ch.x_axis.title = "UW year"
        for c in cols:
            ch.add_data(Reference(sc, min_col=c, min_row=6, max_row=S1), titles_from_data=True)
        ch.set_categories(Reference(sc, min_col=1, min_row=S0, max_row=S1))
        names = ["Current", "Scenario", "Optimised"]
        for s, n in zip(ch.series, names):
            from openpyxl.chart.series import SeriesLabel
            s.tx = SeriesLabel(v=n)
            s.smooth = False
        ch.roundedCorners = False
        ch.legend.position = "b"
        ch.legend.overlay = False
        ch.title.overlay = False
        ch.x_axis.delete = False
        ch.y_axis.delete = False
        sc.add_chart(ch, anchor)

    line_chart("Annual result at our share (USD)", [4, 9, 14], f"R5", "USD", '#,##0')
    line_chart("Combined ratio at our share", [6, 11, 16], f"R21", "Combined ratio", '0%')

    # ================= Performance =================================================
    title_bar(pf, N_YEARS + 3, "Performance by UW year",
              "Pick how to group, which measure, on what basis and which treaty category. Year columns start at the "
              "first UW year (blank = earliest with stats). Figures in USD at the latest FX rate.")
    selector(pf, "A4", "B4", "Group by", "Cedant", "lst_Groups")
    selector(pf, "C4", "D4", "Measure", "Result", "lst_Measures")
    selector(pf, "E4", "F4", "Basis", "100%", "lst_Bases")
    selector(pf, "G4", "H4", "Category", "All", "lst_CatFilter")
    selector(pf, "I4", "J4", "First UW year", None, None)
    pf["J4"].number_format = "0"
    dv(pf, "J4", type="whole", operator="between", formula1="1980", formula2="2100")
    pf["K4"] = '="Showing UW "&sel_FirstYear&" - "&(sel_FirstYear+' + str(N_YEARS - 1) + ')'
    pf["K4"].font = f(9, True, TEAL)
    pf[f"A{PF_HDR}"] = '=Performance!$B$4'
    pf[f"B{PF_HDR}"] = "Rows"
    for k in range(N_YEARS):
        pf.cell(PF_HDR, 3 + k, f"=sel_FirstYear+{k}")
    allc = L(3 + N_YEARS)
    pf.cell(PF_HDR, 3 + N_YEARS, "All years shown")
    for c in range(1, 4 + N_YEARS):
        cell = pf.cell(PF_HDR, c)
        cell.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
        cell.fill = fill(HEADER)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BOX
        if 3 <= c < 3 + N_YEARS:
            cell.number_format = "0"
    pf.row_dimensions[PF_HDR].height = 28
    widths(pf, {"A": 40, "B": 7, **{L(c): 13 for c in range(3, 3 + N_YEARS)}, allc: 15})
    NUM, DEN, GRPK, YR, CAT = rs("AE"), rs("AF"), rs("Y"), rs("H"), rs("R")
    span = f'{YR},">="&sel_FirstYear,{YR},"<="&(sel_FirstYear+{N_YEARS - 1})'

    def cellf(member, yearref):
        m = f"{GRPK},{member}," if member else ""
        y = f"{YR},{yearref}" if yearref else span
        crit = f"{m}{y},{CAT},crit_Cat"
        return (f'IF(sel_IsRatio,IFERROR(SUMIFS({NUM},{crit})/SUMIFS({DEN},{crit}),""),SUMIFS({NUM},{crit}))')

    srcs = ["Cedants!A{s}", "Countries!A{s}", "Lists!D{s}", "Lists!A{s}", "Lists!O{s}", "Treaties!U{s}"]
    for r in range(PF_TOP, PF_BOT + 1):
        s = r - PF_TOP + FIRST
        pf[f"A{r}"] = "=CHOOSE(sel_GroupIdx," + ",".join(x.format(s=s) for x in srcs) + ')&""'
        pf[f"B{r}"] = f'=IF($A{r}="","",COUNTIFS({GRPK},$A{r},{span},{CAT},crit_Cat))'
        for k in range(N_YEARS):
            col = L(3 + k)
            pf[f"{col}{r}"] = f'=IF($A{r}="","",{cellf(f"$A{r}", f"{col}${PF_HDR}")})'
        pf[f"{allc}{r}"] = f'=IF($A{r}="","",{cellf(f"$A{r}", None)})'
    pf[f"A{PF_TOT}"] = "TOTAL"
    pf[f"B{PF_TOT}"] = f'=COUNTIFS({span},{CAT},crit_Cat,{rs("X")},"<>")'
    for k in range(N_YEARS):
        col = L(3 + k)
        pf[f"{col}{PF_TOT}"] = "=" + cellf(None, f"{col}${PF_HDR}")
    pf[f"{allc}{PF_TOT}"] = "=" + cellf(None, None)
    style_range(pf, f"A{PF_TOT}:{allc}{PF_TOT}", font=f(10, True, TOTAL_INK), bg=TOTAL_BG, number_format=MONEY)
    pf[f"B{PF_TOT}"].number_format = "0"
    style_range(pf, f"A{PF_TOP}:A{PF_BOT}", font=f(10), border=False)
    style_range(pf, f"B{PF_TOP}:B{PF_BOT}", font=f(9, color=INK_SOFT), number_format='0;-0;""', border=False)
    style_range(pf, f"C{PF_TOP}:{allc}{PF_BOT}", font=f(10), number_format=MONEY, border=False)
    style_range(pf, f"{allc}{PF_TOP}:{allc}{PF_BOT}", font=f(10, True), number_format=MONEY, border=False)
    ratio_format_rule(pf, f"C{PF_TOT}:{allc}{PF_BOT}")
    pf.conditional_formatting.add(f"C{PF_TOP}:{allc}{PF_BOT}", FormulaRule(
        formula=[f'AND(sel_IsRatio,ISNUMBER(C{PF_TOP}),C{PF_TOP}>1)'], font=Font(name=FONT, bold=True, color="9C0006"),
        fill=fill("FBE3E4")))
    pf.conditional_formatting.add(f"C{PF_TOP}:{allc}{PF_BOT}", FormulaRule(
        formula=[f'AND(sel_IsRatio,ISNUMBER(C{PF_TOP}),C{PF_TOP}<=1)'], fill=fill("E6F4EE")))
    pf.freeze_panes = f"C{PF_TOP}"
    pf.auto_filter.ref = f"A{PF_HDR}:{allc}{PF_BOT}"

    # ================= Dashboard ================================================
    title_bar(dash, 11, "Dashboard", "Whole portfolio by UW year (all categories). 100% figures and our share at current "
              "shares, in USD at the latest FX rate.")
    widths(dash, {"A": 3, "B": 38, "C": 16, "D": 16, "E": 16, "F": 16, "G": 12, "H": 12, "I": 16, "J": 16, "K": 12})
    trk = f"Treaties!$C${FIRST}:$C${LAST_TRT}"
    stx = rs("X")
    kpis = [("Treaties", f'=COUNTIF({trk},"?*")', "#,##0"),
            ("Cedants with treaties", f'=SUMPRODUCT((Cedants!$A${FIRST}:$A${LAST_CED}<>"")*(COUNTIF({trk},Cedants!$A${FIRST}:$A${LAST_CED})>0))', "#,##0"),
            ("Stat rows with figures", f'=SUMPRODUCT(({stx}<>"")*(({rs("I")}<>"")+({rs("J")}<>"")>0))', "#,##0"),
            ("UW years covered", f'=IF(COUNT({rs("H")})=0,"-",MIN({rs("H")})&" - "&MAX({rs("H")}))', "@")]
    header(dash, 5, ["Portfolio", "Value"], 2)
    for i, (lab, fx, nf) in enumerate(kpis):
        r = 6 + i
        dash.cell(r, 2, lab); dash.cell(r, 3, fx)
        style_range(dash, f"B{r}:C{r}", font=f(10))
        dash.cell(r, 3).number_format = nf
        dash.cell(r, 3).font = f(11, True)
        dash.cell(r, 3).alignment = Alignment(horizontal="right")
    live = f'({stx}<>"")*(UPPER({rs("N")})<>"X")'
    checks = [
        ("Stat rows with no FX rate (left out of USD)", f'=SUMPRODUCT({live}*({rs("F")}<>"")*({rs("S")}=""))'),
        ("Generated rows still waiting for figures", f'=COUNTIF({rs("A")},"EMPTY*")'),
        ("Treaties with fewer stat rows than UW years",
         f'=SUMPRODUCT((Treaties!$S${FIRST}:$S${LAST_TRT}<>"")*(Treaties!$V${FIRST}:$V${LAST_TRT}<Treaties!$S${FIRST}:$S${LAST_TRT}))'),
        ("Stat rows whose treaty is not on Treaties", f'=SUMPRODUCT({live}*({rs("Z")}=""))'),
        ("Commission on NP / brokerage on proportional",
         f'=SUMPRODUCT({live}*({rs("R")}="NON_PROPORTIONAL")*({rs("K")}<>""))'
         f'+SUMPRODUCT({live}*({rs("R")}="PROPORTIONAL")*({rs("L")}<>""))'),
        ("Treaties or stat rows not OK in Access",
         f'=SUMPRODUCT(({trk}<>"")*(Treaties!$A${FIRST}:$A${LAST_TRT}<>"OK"))+SUMPRODUCT(({stx}<>"")*({rs("A")}<>"OK"))'),
    ]
    header(dash, 5, ["Data checks", "", "Rows"], 5)
    dash.merge_cells("E5:F5")
    for i, (lab, fx) in enumerate(checks):
        r = 6 + i
        dash.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)
        dash.cell(r, 5, lab); dash.cell(r, 7, fx)
        style_range(dash, f"E{r}:G{r}", font=f(10))
        dash.cell(r, 7).number_format = "#,##0"
        dash.cell(r, 7).font = f(11, True)
    dash.conditional_formatting.add("G6:G11", CellIsRule(operator="greaterThan", formula=["0"],
                                    fill=fill("FFF3CD"), font=Font(name=FONT, bold=True, color="7A5B00")))
    dash["I5"] = "Last pull"; dash["I6"] = '=IF(cfg_LastPull="","never",cfg_LastPull)'
    dash["J5"] = "Last push"; dash["J6"] = '=IF(cfg_LastPush="","never",cfg_LastPush)'
    for c in ("I5", "J5"):
        dash[c].font = f(9, True, INK_SOFT)
    for c in ("I6", "J6"):
        dash[c].font = f(10); dash[c].number_format = STAMP

    DT = 14
    dash.cell(DT - 1, 2, "Portfolio by UW year").font = f(12, True, TEAL)
    dash.cell(DT - 1, 5, '="Years follow Performance: UW "&sel_FirstYear&" onwards"').font = f(9, italic=True, color=INK_SOFT)
    header(dash, DT, ["UW year", "Premium 100%", "Claims 100%", "Comm + brok 100%", "Result 100%", "Loss ratio",
                      "Combined ratio", "Premium @ share", "Result @ share", "CR @ share"], 2)
    dash.row_dimensions[DT].height = 30
    D0 = DT + 1
    for k in range(N_YEARS):
        r = D0 + k
        dash[f"B{r}"] = f"=sel_FirstYear+{k}"
        y = f"$B{r}"
        dash[f"C{r}"] = f'=SUMIFS({rs("T")},{rs("H")},{y})'
        dash[f"D{r}"] = f'=SUMIFS({rs("U")},{rs("H")},{y})'
        dash[f"E{r}"] = f'=SUMIFS({rs("V")},{rs("H")},{y})'
        dash[f"F{r}"] = f'=SUMIFS({rs("W")},{rs("H")},{y})'
        dash[f"G{r}"] = f'=IF(N(C{r})=0,"",D{r}/C{r})'
        dash[f"H{r}"] = f'=IF(N(C{r})=0,"",1-F{r}/C{r})'
        dash[f"I{r}"] = f'=SUMPRODUCT(({rs("H")}={y})*{rs("T")}*{rs("AA")})'
        dash[f"J{r}"] = f'=SUMPRODUCT(({rs("H")}={y})*{rs("W")}*{rs("AA")})'
        dash[f"K{r}"] = f'=IF(N(I{r})=0,"",1-J{r}/I{r})'
    D1 = D0 + N_YEARS - 1
    style_range(dash, f"B{D0}:B{D1}", font=f(10, True), number_format="0", align=Alignment(horizontal="center"), border=False)
    for col in "CDEFIJ":
        style_range(dash, f"{col}{D0}:{col}{D1}", font=f(10), number_format=MONEY, border=False)
    for col in "GHK":
        style_range(dash, f"{col}{D0}:{col}{D1}", font=f(10), number_format=RATIO, border=False)
    tot = D1 + 1
    dash[f"B{tot}"] = "Total"
    for col in "CDEFIJ":
        dash[f"{col}{tot}"] = f"=SUM({col}{D0}:{col}{D1})"
    dash[f"G{tot}"] = f'=IF(N(C{tot})=0,"",D{tot}/C{tot})'
    dash[f"H{tot}"] = f'=IF(N(C{tot})=0,"",1-F{tot}/C{tot})'
    dash[f"K{tot}"] = f'=IF(N(I{tot})=0,"",1-J{tot}/I{tot})'
    style_range(dash, f"B{tot}:K{tot}", font=f(10, True, TOTAL_INK), bg=TOTAL_BG, number_format=MONEY)
    for col in "GHK":
        dash[f"{col}{tot}"].number_format = RATIO
    dash.conditional_formatting.add(f"G{D0}:H{D1}", CellIsRule(operator="greaterThan", formula=["1"],
                                    font=Font(name=FONT, bold=True, color="9C0006")))
    dash.conditional_formatting.add(f"K{D0}:K{D1}", CellIsRule(operator="greaterThan", formula=["1"],
                                    font=Font(name=FONT, bold=True, color="9C0006")))
    ch = LineChart()
    ch.title = "Loss and combined ratio by UW year (100%)"
    ch.height, ch.width = 8, 18
    ch.y_axis.numFmt = "0%"
    ch.y_axis.title = "Ratio"
    ch.x_axis.title = "UW year"
    ch.add_data(Reference(dash, min_col=7, max_col=8, min_row=DT, max_row=D1), titles_from_data=True)
    ch.set_categories(Reference(dash, min_col=2, min_row=D0, max_row=D1))
    ch.roundedCorners = False
    ch.legend.position = "b"; ch.legend.overlay = False; ch.title.overlay = False
    ch.x_axis.delete = False; ch.y_axis.delete = False
    dash.add_chart(ch, f"B{tot + 3}")

    # ================= Cedants / Countries / Rates / Lists =============================
    title_bar(ce, 6, "Cedants", "Every cedant needs a home country. Pre-loaded with Universe's cedant list. Add more with "
              "Add cedant, or type rows and Push lists + cedants. Set Active = N instead of deleting.")
    header(ce, 5, ["Cedant name", "Home country", "Active (Y/N)", "ID", "Status"])
    widths(ce, {"A": 36, "B": 24, "C": 11, "D": 7, "E": 40, "F": 4})
    style_range(ce, f"A{FIRST}:C{LAST_CED}", font=blue, border=False)
    style_range(ce, f"D{FIRST}:E{LAST_CED}", font=grey, border=False)
    style_range(ce, f"C{FIRST}:C{LAST_CED}", align=Alignment(horizontal="center"), border=False)
    dv(ce, f"B{FIRST}:B{LAST_CED}", type="list", formula1="lst_Countries")
    dv(ce, f"C{FIRST}:C{LAST_CED}", type="list", formula1='"Y,N"')
    ce.conditional_formatting.add(f"B{FIRST}:B{LAST_CED}", FormulaRule(formula=[f'AND($A{FIRST}<>"",$B{FIRST}="")'],
                                  fill=fill("F8D7DA")))
    ce.freeze_panes = "A6"
    ced_seed = sorted(UNIVERSE_CEDANTS + (DEMO_CEDANTS if demo else []), key=lambda x: x[0].lower())
    for i, (nm, hc) in enumerate(ced_seed):
        ce.cell(FIRST + i, 1, nm); ce.cell(FIRST + i, 2, hc); ce.cell(FIRST + i, 3, "Y")

    title_bar(co, 7, "Countries", "Reference list (Universe seed plus additional African markets). Add rows to extend.")
    header(co, 5, ["Country", "ISO2", "ISO3", "Region", "ID", "Status"])
    widths(co, {"A": 28, "B": 7, "C": 7, "D": 14, "E": 7, "F": 30, "G": 4})
    style_range(co, f"A{FIRST}:D{LAST_CTRY}", font=blue, border=False)
    style_range(co, f"E{FIRST}:F{LAST_CTRY}", font=grey, border=False)
    for i, (nm, a2, a3) in enumerate(COUNTRIES):
        r = FIRST + i
        co.cell(r, 1, nm); co.cell(r, 2, a2); co.cell(r, 3, a3); co.cell(r, 4, REGION[a2])
    co.freeze_panes = "A6"

    title_bar(ra, 6, "Currencies and FX rates",
              "Rate = USD per 1 unit of currency. Only USD and fixed pegs are pre-filled; enter the rest. "
              "A new effective date adds history in Access; the latest rate per currency is used.")
    header(ra, 5, ["Currency", "Name", "USD per 1 unit", "Effective date", "Status"])
    widths(ra, {"A": 10, "B": 26, "C": 16, "D": 14, "E": 44, "F": 4})
    style_range(ra, f"A{FIRST}:D{LAST_RATE}", font=blue, border=False)
    style_range(ra, f"C{FIRST}:C{LAST_RATE}", number_format=RATE, border=False)
    style_range(ra, f"D{FIRST}:D{LAST_RATE}", number_format=DATE, border=False)
    style_range(ra, f"E{FIRST}:E{LAST_RATE}", font=grey, border=False)
    for i, (code, nm) in enumerate(CURRENCIES):
        r = FIRST + i
        ra.cell(r, 1, code); ra.cell(r, 2, nm)
        if code in PEGS:
            rate, note = PEGS[code]
            ra.cell(r, 3, round(rate, 6)); ra.cell(r, 4, PEG_DATE)
            ra.cell(r, 5, f"{note} - verify before use")
        else:
            ra.cell(r, 5, "NO RATE - enter USD per 1 unit + date")
    ra.conditional_formatting.add(f"E{FIRST}:E{LAST_RATE}", FormulaRule(formula=[f'LEFT($E{FIRST},2)="NO"'], fill=fill("FFF3CD")))
    ra.conditional_formatting.add(f"E{FIRST}:E{LAST_RATE}", FormulaRule(formula=[f'LEFT($E{FIRST},5)="ERROR"'], fill=fill("F8D7DA")))
    ra.freeze_panes = "A6"

    title_bar(li, 15, "Lists", "Treaty types (with category) and classes of business are Universe's predefined lists, "
              "stored in Access. The category decides Commission (proportional) vs Brokerage (non-proportional). "
              "Columns G onwards are fixed lists used by the reports.")
    header(li, 5, ["Treaty type", "Category"], 1)
    header(li, 5, ["Class of business", "Code"], 4)
    header(li, 5, ["Group by"], 7)
    header(li, 5, ["Measure"], 9)
    header(li, 5, ["Basis"], 11)
    header(li, 5, ["Category filter"], 13)
    header(li, 5, ["Categories"], 15)
    widths(li, {"A": 24, "B": 20, "C": 3, "D": 24, "E": 9, "F": 3, "G": 13, "H": 3, "I": 24, "J": 3,
                "K": 16, "L": 3, "M": 16, "N": 3, "O": 20})
    style_range(li, f"A{FIRST}:B{LAST_LIST}", font=blue, border=False)
    style_range(li, f"D{FIRST}:E{LAST_LIST}", font=blue, border=False)
    for i, (a, b) in enumerate(TT):
        li.cell(FIRST + i, 1, a); li.cell(FIRST + i, 2, b)
    for i, (a, b) in enumerate(CB):
        li.cell(FIRST + i, 4, a); li.cell(FIRST + i, 5, b)
    for col, items in ((7, GROUPS), (9, MEASURES), (11, BASES), (13, CAT_FILTER), (15, CATS)):
        for i, v in enumerate(items):
            li.cell(FIRST + i, col, v).font = f(10)
    dv(li, f"B{FIRST}:B{LAST_LIST}", type="list", formula1='"PROPORTIONAL,NON_PROPORTIONAL"')
    li.freeze_panes = "A6"

    # ================= Start ====================================================
    title_bar(st, 6, "Start", "Excel front end over an Access database. Captures treaty statistics by UW year and lets "
              "you test shares. Excel = screens; Access = system of record.")
    widths(st, {"A": 3, "B": 30, "C": 42, "D": 74, "E": 3, "F": 3})
    st["B5"] = "Database"; st["B5"].font = f(12, True, TEAL)
    st["B7"] = "Database path (optional)"; st["B7"].font = f(10, True)
    st["C7"].fill = fill("FFFF00"); st["C7"].font = f(10, color=INPUT_BLUE); st["C7"].border = BOX
    st["D7"] = ("Leave blank to keep TreatyPerformance.accdb next to this workbook. Use a local or network-drive "
                "path (not a OneDrive web link).")
    st["D7"].font = f(9, italic=True, color=INK_SOFT); st["D7"].alignment = Alignment(wrap_text=True, vertical="top")
    st.row_dimensions[7].height = 28
    st["B9"] = "Last pull from Access"; st["B10"] = "Last push to Access"
    for c in ("C9", "C10"):
        st[c].number_format = STAMP; st[c].font = f(10); st[c].border = BOX
    st["B12"] = "Actions"; st["B12"].font = f(12, True, TEAL)
    actions = [
        (13, "Setup_Database", "First run: creates the Access file and tables, loads Universe's treaty types, classes and "
                               "cedants, then this workbook's lists, treaties and stats. Re-run any time: only adds what is missing."),
        (15, "Pull_All", "Reload every sheet from Access (do this before you start editing)."),
        (17, "Push_Masters", "Save Countries, Rates, Lists and Cedants. Every cedant needs a home country."),
        (19, "Push_Treaties", "Save the Treaties sheet (one transaction; blocks on any error or conflicting change)."),
        (21, "Generate_Stat_Rows", "Add one Stats row per UW year in each treaty's stats period (only the missing ones)."),
        (23, "Push_Stats", "Save Stats. Rows still waiting for figures stay on the sheet and are not saved."),
        (25, "Run_Optimiser", "Suggest shares on the Shares sheet within each treaty's min / max share."),
        (27, "Add_Cedant", "Add one cedant (name + home country) straight to Access."),
        (29, "Open_Database", "Open the .accdb in Access (tables, audit_log and the vw_perf_* queries)."),
    ]
    for r, macro, text in actions:
        st.cell(r, 3, f"macro: {macro}").font = f(8, color="AAAAAA")
        st.cell(r, 4, text).font = f(10)
        st.cell(r, 4).alignment = Alignment(wrap_text=True, vertical="top")
        st.row_dimensions[r].height = 30
    st["B13"] = "Buttons appear in column C once the macros are installed."
    st["B13"].font = f(8, italic=True, color=INK_SOFT)
    r0 = 32
    st[f"B{r0}"] = "Workflow"; st[f"B{r0}"].font = f(12, True, TEAL)
    flow = ["Install once: import the 8 .bas files from the vba folder (Developer > Visual Basic > File > Import File), "
            "save as .xlsm, run InstallButtons, then click 1. Setup / create database.",
            "Cedants: check the cedant is listed with its home country (Add cedant if not).",
            "Treaties: one row per treaty with cedant, treaty type, class, currency, stats start and end dates, our "
            "share and optional min / max shares. Push treaties.",
            "Click Generate stat rows: one row per UW year appears on Stats for every treaty.",
            "Stats: enter 100% figures - Premium, Incurred claims, and Commission (proportional) or Brokerage "
            "(non-proportional). Push stats.",
            "Performance: group by cedant, country, class, treaty type, category or treaty; pick the measure and basis.",
            "Shares: type scenario shares, or set the years / risk weight / premium range and Run optimiser. "
            "Compare current, scenario and optimised on Scenarios."]
    for i, t_ in enumerate(flow):
        r = r0 + 1 + i
        st.cell(r, 2, f"{i + 1}.").font = f(10, True)
        st.cell(r, 2).alignment = Alignment(horizontal="right", vertical="top")
        st.cell(r, 3, t_).font = f(10)
        st.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
        st.cell(r, 3).alignment = Alignment(wrap_text=True, vertical="top")
        st.row_dimensions[r].height = 30
    r0 = r0 + len(flow) + 2
    st[f"B{r0}"] = "Example treaty and stat row"; st[f"B{r0}"].font = f(12, True, TEAL)
    header(st, r0 + 1, ["Column", "Example value", "Rule"], 2)
    example = [
        ("Treaties: Cedant", "Tawuniya", "Must exist on Cedants (with a home country)"),
        ("Treaties: Treaty type / Class", "Quota Share / Property", "From Lists; the type's category decides commission vs brokerage"),
        ("Treaties: Currency / Reference", "SAR / MAIN", "Blank reference = MAIN; use LAYER 1, LAYER 2 for XL layers"),
        ("Treaties: Stats start / end", "01/01/2016 - 31/12/2025", "10 UW years: 2016-2025. 01/07/2017 - 30/06/2025 = 8 years: 2017-2024"),
        ("Treaties: Our / min / max share", "10% / 5% / 20%", "Blank min = 0%; blank max = 2 x our share (capped at 100%)"),
        ("Stats: UW year", 2024, "Inside the treaty's stats period"),
        ("Stats: Premium", 180000000, "100% figure, original currency; required (0 if none)"),
        ("Stats: Incurred claims", 95000000, "100%; blank = 0"),
        ("Stats: Commission", 57600000, "Proportional treaties only"),
        ("Stats: Brokerage", None, "Non-proportional treaties only"),
    ]
    for i, (h, v, rule) in enumerate(example):
        r = r0 + 2 + i
        st.cell(r, 2, h).font = f(10, True)
        c = st.cell(r, 3, v); c.font = f(10, color=INPUT_BLUE); c.alignment = Alignment(horizontal="left")
        if isinstance(v, (int, float)) and v > 3000:
            c.number_format = MONEY
        st.cell(r, 4, rule).font = f(10)
        style_range(st, f"B{r}:D{r}")
    r0 = r0 + len(example) + 3
    st[f"B{r0}"] = "Definitions"; st[f"B{r0}"].font = f(12, True, TEAL)
    defs = [("Result", "Premium - incurred claims - commission - brokerage"),
            ("Loss ratio", "Incurred claims / premium"),
            ("Combined ratio", "(Incurred claims + commission + brokerage) / premium"),
            ("Our share basis", "100% figures x share (current, scenario or optimised) per treaty"),
            ("Optimiser objective", "Average annual result - risk weight x standard deviation of annual result, "
                                    "within min / max shares and the optional premium range")]
    for i, (a, b) in enumerate(defs):
        r = r0 + 1 + i
        st.cell(r, 2, a).font = f(10, True); st.cell(r, 3, b).font = f(10)
        st.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
    r0 = r0 + len(defs) + 2
    st[f"B{r0}"] = "Colour legend"; st[f"B{r0}"].font = f(12, True, TEAL)
    legend = [("Blue text", "Input you type", None, INPUT_BLUE), ("Yellow cell", "Setting / filter", "FFFF00", None),
              ("Grey-blue cell", "Formula - do not type", CALC_BG, None),
              ("Status = ERROR", "Row refused on push - read the message", "F8D7DA", None),
              ("Status = EMPTY", "Generated row waiting for figures", "FFF3CD", None)]
    for i, (a, b, bg, fg) in enumerate(legend):
        r = r0 + 1 + i
        st[f"B{r}"] = a; st[f"C{r}"] = b
        st[f"B{r}"].font = f(10, color=fg or "16323B"); st[f"C{r}"].font = f(10)
        if bg:
            st[f"B{r}"].fill = fill(bg)
    r0 = r0 + len(legend) + 2
    st[f"B{r0}"] = "How this mirrors Universe"; st[f"B{r0}"].font = f(12, True, TEAL)
    header(st, r0 + 1, ["Universe", "Here", "What it does"], 2)
    arch = [("client (React screens)", "Excel sheets + PerfMacros.bas", "Entry grids, reports, buttons"),
            ("server/src/services, validation", "PerfService.bas", "Validation, push/pull, transactions, row generation"),
            ("services (analytics)", "PerfOptimiser.bas", "Share optimiser"),
            ("modules/*/repositories", "PerfRepo.bas", "Every SQL statement"),
            ("db/pool.js, withTransaction.js", "PerfDb.bas", "ADO connection, parameterised commands"),
            ("db/migrations + runMigrations.js", "PerfMigrate.bas + schema_migrations", "Numbered forward-only schema changes"),
            ("startup/ensureReferenceData.js", "PerfSeed.bas", "Predefined treaty types, classes, cedants"),
            ("treaty / statistics tables", "treaty, treaty_stat", "Treaty master + one row per UW year"),
            ("db/optimisticLock.js", "row_version column", "Refuses saves over someone else's change"),
            ("audit_log", "audit_log", "Who changed what, when")]
    for i, row in enumerate(arch):
        r = r0 + 2 + i
        for j, v in enumerate(row):
            st.cell(r, 2 + j, v).font = f(10)
        style_range(st, f"B{r}:D{r}")

    wb.active = 0
    wb.save(path)


if __name__ == "__main__":
    out = sys.argv[1]
    build(f"{out}/Performance_Tracker.xlsx", demo=False)
    build(f"{out}/Performance_Tracker_DEMO.xlsx", demo=True)
    print("built")
