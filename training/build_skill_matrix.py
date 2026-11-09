"""Builds the Frozen Production Skill Matrix workbook (frozen only, no baking).

Run:   python training/build_skill_matrix.py
Output: training/Frozen_Production_Skill_Matrix.xlsx

Tabs: Start Here, Dashboard, Skill Matrix, Person Card, Shift Coverage,
Training Plan, Skills Library, Role Requirements, Lists, Calc (hidden helper).
"""
import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, str(Path(__file__).parent))
from skill_data import CATEGORIES, LEVELS, ROLES, SKILLS  # noqa: E402

OUT = Path(__file__).with_name("Frozen_Production_Skill_Matrix.xlsx")
FONT = "Arial"
SHIFTS = ["1st Shift", "2nd Shift", "3rd Shift"]
N_EMP = 60
R0 = 6                         # first data row on Skill Matrix / Role Req / Calc
R1 = R0 + N_EMP - 1            # last employee row
SC = 10                        # first skill column (J)
NS = len(SKILLS)
FSC, LSC = L(SC), L(SC + NS - 1)
ROLE_NAMES = [r[1] for r in ROLES]
REQ_LAST = R0 + len(ROLES) - 1

NAVY, INK, MUTED = "1F3864", "1F2A36", "5B6B78"
GAP_FILL, GAP_FONT = "FFC7CE", "9C0006"
OK_FILL, OK_FONT = "C6EFCE", "006100"
AMB_FILL, AMB_FONT = "FFEB9C", "7F6000"
INPUT = PatternFill("solid", fgColor="FFF9DB")
GREYF = PatternFill("solid", fgColor="F2F4F6")
thin = Side(style="thin", color="D0D7DE")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
C = Alignment(horizontal="center", vertical="center", wrap_text=True)
LFT = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)


def f(size=10, bold=False, color=INK, italic=False, u=None):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic, underline=u)


def fill(c):
    return PatternFill("solid", fgColor=c)


def put(ws, ref, value=None, font=None, fl=None, al=None, border=None, fmt=None):
    cell = ws[ref] if isinstance(ref, str) else ws.cell(*ref)
    if value is not None:
        cell.value = value
    if font: cell.font = font
    if fl: cell.fill = fl
    if al: cell.alignment = al
    if border: cell.border = border
    if fmt: cell.number_format = fmt
    return cell


def head(ws, r, c, text, color=NAVY, rotate=False):
    cell = ws.cell(r, c, text)
    cell.font = f(10, True, "FFFFFF")
    cell.fill = fill(color)
    cell.border = BOX
    cell.alignment = Alignment(horizontal="center", vertical="bottom" if rotate else "center",
                               wrap_text=True, text_rotation=90 if rotate else 0)
    return cell


def merge(ws, rng, value=None, **kw):
    ws.merge_cells(rng)
    return put(ws, rng.split(":")[0], value, **kw)


def link(ws, ref, text, target, color=NAVY, bg=None, size=10):
    cell = put(ws, ref, text, f(size, True, color, u="single" if not bg else None),
               fill(bg) if bg else None, C)
    cell.hyperlink = f"#'{target}'!A1"
    return cell


def nav(ws, ref1, ref2):
    link(ws, ref1, "◂ Start Here", "Start Here")
    link(ws, ref2, "▦ Dashboard", "Dashboard")


def title(ws, text, sub):
    put(ws, "A1", text, f(18, True, NAVY))
    put(ws, "A2", sub, f(10, color=MUTED, italic=True))
    ws.row_dimensions[1].height = 30
    ws.sheet_view.showGridLines = False


wb = Workbook()
S_START = wb.active
S_START.title = "Start Here"
S_DASH = wb.create_sheet("Dashboard")
S_MX = wb.create_sheet("Skill Matrix")
S_CARD = wb.create_sheet("Person Card")
S_COV = wb.create_sheet("Shift Coverage")
S_PLAN = wb.create_sheet("Training Plan")
S_LIB = wb.create_sheet("Skills Library")
S_REQ = wb.create_sheet("Role Requirements")
S_LIST = wb.create_sheet("Lists")
S_CALC = wb.create_sheet("Calc")

# ------------------------------------------------------------------ Lists
ws = S_LIST
title(ws, "Lists", "Drop-down values. Shift names (yellow) can be renamed.")
for c, h in ((1, "Roles"), (3, "Shifts"), (5, "Dashboard filter"), (7, "Training status")):
    head(ws, 4, c, h)
for i, n in enumerate(ROLE_NAMES):
    put(ws, (5 + i, 1), n, f())
for i, s in enumerate(SHIFTS):
    put(ws, (5 + i, 3), s, f(), INPUT)
put(ws, "E5", "All shifts", f())
for i in range(3):
    put(ws, (6 + i, 5), f"=C{5 + i}", f())
for i, s in enumerate(["Not started", "In progress", "Ready for sign-off", "Complete", "On hold"]):
    put(ws, (5 + i, 7), s, f())
for col, w in zip("ABCDEFG", (32, 3, 14, 3, 16, 3, 20)):
    ws.column_dimensions[col].width = w
ROLE_LIST = f"Lists!$A$5:$A${4 + len(ROLES)}"
SHIFT_LIST = "Lists!$C$5:$C$7"
FILTER_LIST = "Lists!$E$5:$E$8"
SH = ["Lists!$C$5", "Lists!$C$6", "Lists!$C$7"]

# ------------------------------------------------------------------ skill header helper


def skill_headers(ws):
    col = SC
    for cat, color, items in CATEGORIES:
        start = col
        for sk in items:
            head(ws, 4, col, sk[0], color)
            head(ws, 5, col, sk[1], color, rotate=True)
            ws.column_dimensions[L(col)].width = 6
            col += 1
        ws.merge_cells(start_row=3, start_column=start, end_row=3, end_column=col - 1)
        head(ws, 3, start, cat, color)
    ws.row_dimensions[3].height = 32
    ws.row_dimensions[5].height = 200


# ------------------------------------------------------------------ Role Requirements
ws = S_REQ
title(ws, "Role Requirements", "Minimum level each role needs per skill (blank = not needed). Edit the yellow cells – everything else updates.")
nav(ws, "E1", "G1")
skill_headers(ws)
for c, h, w in ((2, "Code", 7), (3, "Role", 30)):
    ws.merge_cells(start_row=3, start_column=c, end_row=5, end_column=c)
    head(ws, 3, c, h)
    ws.column_dimensions[L(c)].width = w
ws.merge_cells(start_row=3, start_column=4, end_row=5, end_column=SC - 1)
head(ws, 3, 4, "Skills needed")
ws.column_dimensions["A"].width = 2
for c in range(4, SC):
    ws.column_dimensions[L(c)].width = 2.6
for i, (code, name) in enumerate(ROLES):
    r = R0 + i
    put(ws, (r, 2), code, f(10, True), border=BOX, al=C)
    put(ws, (r, 3), name, f(10, True), border=BOX, al=LFT)
    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=SC - 1)
    put(ws, (r, 4), f'=COUNTIF({FSC}{r}:{LSC}{r},">0")', f(10, True), al=C, border=BOX)
    for j, sk in enumerate(SKILLS):
        v = sk["req"][i]
        put(ws, (r, SC + j), v or None, f(10, True), INPUT, C, BOX)
    ws.row_dimensions[r].height = 22
REQ_GRID = f"'Role Requirements'!${FSC}${R0}:${LSC}${REQ_LAST}"
REQ_ROLES = f"'Role Requirements'!$C${R0}:$C${REQ_LAST}"
rng = f"{FSC}{R0}:{LSC}{REQ_LAST}"
for lvl, _n, _d, bg, fg in LEVELS[1:]:
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[str(lvl)], fill=fill(bg),
                                                  font=Font(name=FONT, bold=True, color=fg)))
dv = DataValidation(type="whole", operator="between", formula1="0", formula2="4", allow_blank=True,
                    error="Enter 0 to 4", errorTitle="Level")
ws.add_data_validation(dv); dv.add(rng)
put(ws, (REQ_LAST + 2, 3), "Starting values are a recommended baseline for a frozen dough line. Adjust to your equipment and job descriptions.",
    f(9, color=MUTED, italic=True))
ws.freeze_panes = ws.cell(R0, SC)

# ------------------------------------------------------------------ sample team


def sample_team():
    seed = [7]

    def rnd():
        seed[0] = (seed[0] * 16807) % 2147483647
        return seed[0] / 2147483647
    plan = [("PK", 0), ("PK", 1), ("PK", 2), ("FR", 0), ("FR", 1), ("MX", 0), ("MX", 2), ("FO", 1),
            ("BT", 0), ("WH", 1), ("PA", 2), ("LL", 0), ("LL", 1), ("SV", 0)]
    codes = [r[0] for r in ROLES]
    team = []
    for i, (code, sh) in enumerate(plan):
        ri = codes.index(code)
        lv = []
        for sk in SKILLS:
            q = sk["req"][ri]
            x = rnd()
            v = q
            if q > 0 and x < .16:
                v = q - 1 - (1 if x < .04 else 0)
            elif x > .86 and q < 4:
                v = q + 1
            elif q == 0 and x > .7:
                v = 1
            lv.append(max(0, min(4, v)))
        team.append((f"S-{101 + i}", f"Sample {chr(65 + i)}", ROLE_NAMES[ri], SHIFTS[sh], f"Line {1 + i % 2}", lv))
    return team


SAMPLE = sample_team()

# ------------------------------------------------------------------ Skill Matrix
ws = S_MX
title(ws, "Skill Matrix", "Pick a level 0–4 in each cell.  RED = below what the role needs.  0 Not trained · 1 In training · 2 Can do with help · 3 Qualified · 4 Trainer")
nav(ws, "F1", "H1")
skill_headers(ws)
fixed = [("A", "Emp ID", 9), ("B", "Name", 22), ("C", "Role", 26), ("D", "Shift", 11), ("E", "Line / Area", 11),
         ("F", "Skills needed", 6), ("G", "Skills met", 6), ("H", "Gaps", 6), ("I", "% Ready", 8)]
for col, h, w in fixed:
    ci = ws[col + "3"].column
    ws.merge_cells(start_row=3, start_column=ci, end_row=5, end_column=ci)
    head(ws, 3, ci, h, NAVY if ci < 6 else "404040", rotate=ci >= 6)
    ws.column_dimensions[col].width = w
for k in range(N_EMP):
    r = R0 + k
    ex = SAMPLE[k] if k < len(SAMPLE) else None
    for ci in range(1, 6):
        put(ws, (r, ci), ex[ci - 1] if ex else None, f(10), INPUT, LFT if ci in (2, 3) else C, BOX)
    req_row = f"INDEX({REQ_GRID},MATCH($C{r},{REQ_ROLES},0),0)"
    lv = f"${FSC}{r}:${LSC}{r}"
    put(ws, (r, 6), f'=IF($C{r}="","",SUMPRODUCT(--({req_row}>0)))')
    put(ws, (r, 7), f'=IF($C{r}="","",SUMPRODUCT(({req_row}>0)*({lv}>={req_row})))')
    put(ws, (r, 8), f'=IF($C{r}="","",F{r}-G{r})')
    put(ws, (r, 9), f'=IF(OR($C{r}="",F{r}=0),"",G{r}/F{r})', fmt="0%")
    for ci in range(6, 10):
        put(ws, (r, ci), font=f(10, True), fl=GREYF, al=C, border=BOX)
    for j in range(NS):
        put(ws, (r, SC + j), ex[5][j] if ex else None, f(10, True), al=C, border=BOX)
    ws.row_dimensions[r].height = 18
grid = f"{FSC}{R0}:{LSC}{R1}"
tl = f"{FSC}{R0}"
reqc = f"INDEX('Role Requirements'!{FSC}${R0}:{FSC}${REQ_LAST},MATCH($C{R0},{REQ_ROLES},0))"
ws.conditional_formatting.add(grid, FormulaRule(
    formula=[f'AND($C{R0}<>"",IFERROR({reqc},0)>0,IF(ISNUMBER({tl}),{tl},0)<IFERROR({reqc},0))'],
    stopIfTrue=True, fill=fill("FF7C80"), font=Font(name=FONT, bold=True, color=GAP_FONT)))
for lvl, _n, _d, bg, fg in LEVELS:
    ws.conditional_formatting.add(grid, FormulaRule(formula=[f"AND(ISNUMBER({tl}),{tl}={lvl})"],
                                                    fill=fill(bg), font=Font(name=FONT, bold=True, color=fg)))
pr = f"I{R0}:I{R1}"
ws.conditional_formatting.add(pr, FormulaRule(formula=[f"AND(ISNUMBER(I{R0}),I{R0}>=1)"], stopIfTrue=True, fill=fill(OK_FILL), font=Font(name=FONT, bold=True, color=OK_FONT)))
ws.conditional_formatting.add(pr, FormulaRule(formula=[f"AND(ISNUMBER(I{R0}),I{R0}>=0.8)"], stopIfTrue=True, fill=fill(AMB_FILL), font=Font(name=FONT, bold=True, color=AMB_FONT)))
ws.conditional_formatting.add(pr, FormulaRule(formula=[f"ISNUMBER(I{R0})"], fill=fill(GAP_FILL), font=Font(name=FONT, bold=True, color=GAP_FONT)))
ws.conditional_formatting.add(f"H{R0}:H{R1}", FormulaRule(formula=[f"AND(ISNUMBER(H{R0}),H{R0}>0)"], font=Font(name=FONT, bold=True, color="C00000")))
dv_l = DataValidation(type="list", formula1='"0,1,2,3,4"', allow_blank=True, error="Pick 0 to 4", errorTitle="Level",
                      prompt="0 Not trained | 1 In training | 2 Can do with help | 3 Qualified | 4 Trainer", promptTitle="Skill level")
dv_r = DataValidation(type="list", formula1=f"={ROLE_LIST}", allow_blank=True)
dv_s = DataValidation(type="list", formula1=f"={SHIFT_LIST}", allow_blank=True)
for d in (dv_l, dv_r, dv_s):
    ws.add_data_validation(d)
dv_l.add(grid); dv_r.add(f"C{R0}:C{R1}"); dv_s.add(f"D{R0}:D{R1}")
tr = R1 + 2
for i, (lab, crit) in enumerate((("# Qualified (3+)", ">=3"), ("# Trainers (4)", "=4"))):
    r = tr + i
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=SC - 1)
    put(ws, (r, 2), lab, f(10, True), al=Alignment(horizontal="right"))
    for j in range(NS):
        col = L(SC + j)
        put(ws, (r, SC + j), f'=COUNTIF({col}{R0}:{col}{R1},"{crit}")', f(10, True), GREYF, C, BOX)
ws.freeze_panes = ws.cell(R0, SC)
ws.auto_filter.ref = f"A5:{LSC}{R1}"
MGRID = f"'Skill Matrix'!${FSC}${R0}:${LSC}${R1}"
MNAMES = f"'Skill Matrix'!$B${R0}:$B${R1}"
MCOL = lambda c: f"'Skill Matrix'!${c}${R0}:${c}${R1}"  # noqa: E731

# ------------------------------------------------------------------ Calc (hidden helper)
ws = S_CALC
put(ws, "A1", "Helper calculations for the Dashboard – do not edit.", f(10, True))
put(ws, "A2", "Shift filter")
put(ws, "B2", '=IF(Dashboard!$K$2="All shifts","*",Dashboard!$K$2)')
for c, h in ((1, "Shift"), (2, "Role"), (3, "Gaps"), (4, "Rank key")):
    put(ws, (5, c), h, f(10, True))
cat_cols = {}
col = 5
skill_pos = 0
for cat, _color, items in CATEGORIES:
    a, b = L(SC + skill_pos), L(SC + skill_pos + len(items) - 1)
    put(ws, (4, col), cat, f(9, True))
    put(ws, (5, col), "need", f(9)); put(ws, (5, col + 1), "met", f(9))
    cat_cols[cat] = (L(col), L(col + 1))
    for k in range(N_EMP):
        r = R0 + k
        reqs = f"INDEX('Role Requirements'!${a}${R0}:${b}${REQ_LAST},MATCH($B{r},{REQ_ROLES},0),0)"
        put(ws, (r, col), f'=IF($B{r}="",0,SUMPRODUCT(--({reqs}>0)))')
        put(ws, (r, col + 1), f"=IF($B{r}=\"\",0,SUMPRODUCT(({reqs}>0)*('Skill Matrix'!{a}{r}:{b}{r}>={reqs})))")
    col += 2
    skill_pos += len(items)
for k in range(N_EMP):
    r = R0 + k
    put(ws, (r, 1), f"=IF('Skill Matrix'!D{r}=\"\",\"\",'Skill Matrix'!D{r})")
    put(ws, (r, 2), f"=IF('Skill Matrix'!C{r}=\"\",\"\",'Skill Matrix'!C{r})")
    put(ws, (r, 3), f"=IF('Skill Matrix'!H{r}=\"\",0,'Skill Matrix'!H{r})")
    put(ws, (r, 4), f'=IF(AND($B{r}<>"",OR($B$2="*",$A{r}=Dashboard!$K$2),$C{r}>0),$C{r}+(1000-ROW())/100000,0)')
CALC = lambda c: f"Calc!${c}${R0}:${c}${R1}"  # noqa: E731
ws.sheet_state = "hidden"

# ------------------------------------------------------------------ Shift Coverage
ws = S_COV
title(ws, "Shift Coverage", "How many people are Qualified (level 3+) per skill on each shift. Set the minimum per shift in the yellow column.")
nav(ws, "E1", "G1")
cols = [("ID", 7), ("Area", 22), ("Skill", 34), ("Min per shift", 9), (f"={SH[0]}", 10), (f"={SH[1]}", 10),
        (f"={SH[2]}", 10), ("Total qualified", 10), ("Trainers (4)", 10), ("People who need it", 10), ("Status", 14), ("key", 4)]
for i, (h, w) in enumerate(cols, start=1):
    head(ws, 4, i, h)
    ws.column_dimensions[L(i)].width = w
ws.row_dimensions[4].height = 36
mshift = MCOL("D")
mroles = MCOL("C")
for k, sk in enumerate(SKILLS):
    r = 5 + k
    c = L(SC + k)
    rng = f"'Skill Matrix'!${c}${R0}:${c}${R1}"
    reqcol = f"'Role Requirements'!${c}${R0}:${c}${REQ_LAST}"
    put(ws, (r, 1), sk["id"], f(10, True, "FFFFFF"), fill(sk["color"]), C, BOX)
    put(ws, (r, 2), sk["cat"], f(10), None, LFT, BOX)
    put(ws, (r, 3), sk["name"], f(10, True), None, LFT, BOX)
    put(ws, (r, 4), 2, f(10), INPUT, C, BOX)
    for i, sc in enumerate("EFG"):
        put(ws, (r, 5 + i), f'=COUNTIFS({mshift},{sc}$4,{rng},">=3")', f(10), None, C, BOX)
    put(ws, (r, 8), f'=COUNTIF({rng},">=3")', f(10, True), None, C, BOX)
    put(ws, (r, 9), f"=COUNTIF({rng},4)", f(10), None, C, BOX)
    put(ws, (r, 10), f"=SUMPRODUCT(({reqcol}>0)*COUNTIF({mroles},{REQ_ROLES}))", f(10), None, C, BOX)
    short = ",".join(f"AND(COUNTIF({mshift},{sc}$4)>0,{sc}{r}<$D{r})" for sc in "EFG")
    put(ws, (r, 11), f'=IF(J{r}=0,"Not needed",IF(H{r}<=1,"HIGH RISK",IF(OR({short}),"SHORT","OK")))', f(10, True), None, C, BOX)
    put(ws, (r, 12), f'=IF(K{r}="HIGH RISK",2,IF(K{r}="SHORT",1,0))+IF(OR(K{r}="HIGH RISK",K{r}="SHORT"),(100-H{r})/1000+(100-ROW())/1000000,0)',
        f(8, color="FFFFFF"))
st = f"K5:K{4 + NS}"
for val, bg, fg in (("OK", OK_FILL, OK_FONT), ("SHORT", AMB_FILL, AMB_FONT), ("HIGH RISK", GAP_FILL, GAP_FONT), ("Not needed", "EDEFF2", "7F7F7F")):
    ws.conditional_formatting.add(st, CellIsRule(operator="equal", formula=[f'"{val}"'], fill=fill(bg), font=Font(name=FONT, bold=True, color=fg)))
ws.conditional_formatting.add(f"E5:G{4 + NS}", FormulaRule(formula=[f"AND($K5<>\"Not needed\",COUNTIF({mshift},E$4)>0,E5<$D5)"],
                                                         fill=fill(GAP_FILL), font=Font(name=FONT, bold=True, color=GAP_FONT)))
ws.column_dimensions["L"].hidden = True
ws.freeze_panes = "D5"
ws.auto_filter.ref = f"A4:K{4 + NS}"
put(ws, (6 + NS, 2), "HIGH RISK = 0 or 1 qualified people in the whole team.  SHORT = a staffed shift has fewer than the minimum.  Not needed = no one on the team has a role that needs it.",
    f(9, color=MUTED, italic=True))
COV = lambda c: f"'Shift Coverage'!${c}$5:${c}${4 + NS}"  # noqa: E731

# ------------------------------------------------------------------ Dashboard
ws = S_DASH
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 2
ws.column_dimensions["B"].width = 28
for c in "CDEFGHIJKLM":
    ws.column_dimensions[c].width = 11.5
ws.column_dimensions["N"].hidden = True
merge(ws, "B2:H2", "Frozen Production – Skills Dashboard", font=f(20, True, NAVY), al=Alignment(vertical="center"))
ws.row_dimensions[2].height = 34
put(ws, "B3", "C.H. Guenther · frozen operations only (no baking) · updates automatically from the Skill Matrix", f(10, color=MUTED, italic=True))
put(ws, "J2", "Show shift ▸", f(11, True, NAVY), al=Alignment(horizontal="right", vertical="center"))
merge(ws, "K2:L2", "All shifts", font=f(12, True), fl=INPUT, al=C, border=BOX)
ws["L2"].border = BOX
dvf = DataValidation(type="list", formula1=f"={FILTER_LIST}", allow_blank=False, prompt="Pick a shift to filter the dashboard", promptTitle="Shift")
ws.add_data_validation(dvf); dvf.add("K2")
link(ws, "M2", "◂ Start", "Start Here")
CR = "Calc!$B$2"

tiles = [
    ("B5:C5", "B6:C7", "B8:C8", "TEAM READINESS",
     f"=IFERROR(SUMIFS({MCOL('G')},{MCOL('D')},{CR})/SUMIFS({MCOL('F')},{MCOL('D')},{CR}),0)", "0%",
     f'=SUMIFS({MCOL("G")},{MCOL("D")},{CR})&" of "&SUMIFS({MCOL("F")},{MCOL("D")},{CR})&" required skills met"'),
    ("D5:F5", "D6:F7", "D8:F8", "FULLY QUALIFIED PEOPLE",
     f'=COUNTIFS({MCOL("D")},{CR},{MCOL("H")},0)', "0",
     f'="out of "&COUNTIFS({MCOL("D")},{CR},{MCOL("C")},"<>")&" people"'),
    ("G5:I5", "G6:I7", "G8:I8", "OPEN SKILL GAPS",
     f'=SUMIFS({MCOL("H")},{MCOL("D")},{CR})', "0", '="red cells on the Skill Matrix"'),
    ("J5:L5", "J6:L7", "J8:L8", "HIGH-RISK SKILLS",
     f'=COUNTIF({COV("K")},"HIGH RISK")', "0", '="only 0–1 qualified people (all shifts)"'),
]
tile_fill = fill("EEF3F8")
for lab, val, sub, text, formula, fmt, subf in tiles:
    a, b = lab.split(":")
    for row in ws[f"{a[0]}5:{b[0]}8"]:
        for cell in row:
            cell.fill = tile_fill
    merge(ws, lab, text, font=f(9, True, MUTED), al=Alignment(horizontal="left", vertical="bottom", indent=1))
    merge(ws, val, formula, font=f(30, True, NAVY), al=Alignment(horizontal="left", vertical="center", indent=1), fmt=fmt)
    merge(ws, sub, subf, font=f(9, color=MUTED), al=Alignment(horizontal="left", vertical="top", indent=1))
ws.row_dimensions[6].height = 26
ws.row_dimensions[7].height = 26
ws.row_dimensions[8].height = 18
ws.conditional_formatting.add("B6", CellIsRule(operator="greaterThanOrEqual", formula=["0.95"], font=Font(name=FONT, size=30, bold=True, color=OK_FONT)))
ws.conditional_formatting.add("B6", CellIsRule(operator="lessThan", formula=["0.8"], font=Font(name=FONT, size=30, bold=True, color="C00000")))
ws.conditional_formatting.add("G6", CellIsRule(operator="greaterThan", formula=["0"], font=Font(name=FONT, size=30, bold=True, color="C00000")))
ws.conditional_formatting.add("J6", CellIsRule(operator="greaterThan", formula=["0"], font=Font(name=FONT, size=30, bold=True, color="C00000")))


def section(ref, text):
    put(ws, ref, text, f(13, True, NAVY))


# readiness by area table + heatmap (rows 10-22)
section("B10", "Readiness by area")
put(ws, "D10", "% of required skills met", f(9, color=MUTED, italic=True))
for c, h in zip("BCDEF", ("Area", "Selected", f"={SH[0]}", f"={SH[1]}", f"={SH[2]}")):
    head(ws, 11, ws[c + "11"].column, h)
area_rows = []
need_cols = [CALC(v[0]) for v in cat_cols.values()]
met_cols = [CALC(v[1]) for v in cat_cols.values()]


def ratio_all(crit):
    num = "+".join(f"SUMIFS({m},{CALC('A')},{crit})" for m in met_cols)
    den = "+".join(f"SUMIFS({n},{CALC('A')},{crit})" for n in need_cols)
    return f'=IFERROR(({num})/({den}),"–")'


for i, (cat, color, _items) in enumerate(CATEGORIES + [("All areas", NAVY, None)]):
    r = 12 + i
    area_rows.append(r)
    put(ws, (r, 2), cat, f(10, cat == "All areas"), None, LFT, BOX)
    if cat == "All areas":
        put(ws, (r, 3), ratio_all(CR), f(10, True), None, C, BOX, "0%")
        for j, s in enumerate(SH):
            put(ws, (r, 4 + j), ratio_all(s), f(10, True), None, C, BOX, "0%")
    else:
        nc, mc = cat_cols[cat]
        put(ws, (r, 3), f'=IFERROR(SUMIFS({CALC(mc)},{CALC("A")},{CR})/SUMIFS({CALC(nc)},{CALC("A")},{CR}),"–")', f(10, True), None, C, BOX, "0%")
        for j, s in enumerate(SH):
            put(ws, (r, 4 + j), f'=IFERROR(SUMIFS({CALC(mc)},{CALC("A")},{s})/SUMIFS({CALC(nc)},{CALC("A")},{s}),"–")', f(10, True), None, C, BOX, "0%")
    ws.row_dimensions[r].height = 20
last_area = area_rows[-1]
heat = f"C12:F{last_area}"
ws.conditional_formatting.add(heat, FormulaRule(formula=["AND(ISNUMBER(C12),C12>=0.95)"], stopIfTrue=True, fill=fill(OK_FILL), font=Font(name=FONT, bold=True, color=OK_FONT)))
ws.conditional_formatting.add(heat, FormulaRule(formula=["AND(ISNUMBER(C12),C12>=0.8)"], stopIfTrue=True, fill=fill(AMB_FILL), font=Font(name=FONT, bold=True, color=AMB_FONT)))
ws.conditional_formatting.add(heat, FormulaRule(formula=["ISNUMBER(C12)"], fill=fill(GAP_FILL), font=Font(name=FONT, bold=True, color=GAP_FONT)))
put(ws, (last_area + 1, 2), "Green 95%+ · Amber 80–94% · Red below 80% · – = no one on that shift needs these skills", f(8, color=MUTED, italic=True))

ch = BarChart()
ch.type = "bar"
ch.style = 10
ch.title = "Readiness by area (selected shift)"
ch.add_data(Reference(ws, min_col=3, min_row=11, max_row=last_area - 1), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=2, min_row=12, max_row=last_area - 1))
ch.legend = None
ch.y_axis.scaling.min = 0
ch.y_axis.scaling.max = 1
ch.y_axis.majorUnit = 0.25
ch.y_axis.number_format = "0%"
ch.x_axis.scaling.orientation = "maxMin"
ch.x_axis.delete = False
ch.y_axis.delete = False
ch.series[0].graphicalProperties.solidFill = "2E75B6"
ch.series[0].graphicalProperties.line.solidFill = "2E75B6"
ch.dataLabels = DataLabelList()
ch.dataLabels.showVal = True
ch.dataLabels.showSerName = False
ch.dataLabels.showCatName = False
ch.dataLabels.showLegendKey = False
ch.dataLabels.showPercent = False
ch.dataLabels.numFmt = "0%"
ch.gapWidth = 60
ch.width, ch.height = 17, 8.2
ws.add_chart(ch, "H10")

# who needs training most (rows 24-36)
section("B24", "Who needs training most")
put(ws, "D24", "most gaps first · selected shift", f(9, color=MUTED, italic=True))
for c, h in zip("BCDEG", ("Name", "Gaps", "% Ready", "Role", "Shift")):
    head(ws, 25, ws[c + "25"].column, h)
ws.merge_cells("E25:F25")
for k in range(1, 11):
    r = 25 + k
    put(ws, (r, 14), f"=LARGE({CALC('D')},{k})")
    idx = f"MATCH($N{r},{CALC('D')},0)"
    put(ws, (r, 2), f'=IF($N{r}<=0,"",INDEX({MNAMES},{idx}))', f(10, True), None, LFT, BOX)
    put(ws, (r, 3), f'=IF($N{r}<=0,"",INDEX({MCOL("H")},{idx}))', f(10, True, "C00000"), None, C, BOX)
    put(ws, (r, 4), f'=IF($N{r}<=0,"",INDEX({MCOL("I")},{idx}))', f(10), None, C, BOX, "0%")
    ws.merge_cells(f"E{r}:F{r}")
    put(ws, (r, 5), f'=IF($N{r}<=0,"",INDEX({MCOL("C")},{idx}))', f(9), None, LFT, BOX)
    put(ws, (r, 7), f'=IF($N{r}<=0,"",INDEX({MCOL("D")},{idx}))', f(9), None, C, BOX)
put(ws, "B36", '=IF(N26<=0,"No gaps – everyone meets their role.","Tip: open the Person Card tab to see exactly which skills each person is missing.")', f(9, color=MUTED, italic=True))

# % ready by shift chart
ch2 = BarChart()
ch2.type = "col"
ch2.style = 10
ch2.title = "Readiness by shift (all areas)"
ch2.add_data(Reference(ws, min_col=4, max_col=6, min_row=last_area), from_rows=True, titles_from_data=False)
ch2.set_categories(Reference(ws, min_col=4, max_col=6, min_row=11))
ch2.legend = None
ch2.y_axis.scaling.min = 0
ch2.y_axis.scaling.max = 1
ch2.y_axis.number_format = "0%"
ch2.y_axis.majorUnit = 0.25
ch2.x_axis.delete = False
ch2.y_axis.delete = False
ch2.series[0].graphicalProperties.solidFill = "548235"
ch2.series[0].graphicalProperties.line.solidFill = "548235"
ch2.dataLabels = DataLabelList()
ch2.dataLabels.showVal = True
ch2.dataLabels.showSerName = False
ch2.dataLabels.showCatName = False
ch2.dataLabels.showLegendKey = False
ch2.dataLabels.showPercent = False
ch2.dataLabels.numFmt = "0%"
ch2.gapWidth = 80
ch2.width, ch2.height = 13, 6.6
ws.add_chart(ch2, "I24")

# biggest coverage risks (rows 38-50)
section("B38", "Biggest coverage risks")
put(ws, "D38", "skills with too few qualified people · all shifts", f(9, color=MUTED, italic=True))
for c, h in zip("BCDEFG", ("Skill", "Status", f"={SH[0]}", f"={SH[1]}", f"={SH[2]}", "Trainers")):
    head(ws, 39, ws[c + "39"].column, h)
for k in range(1, 11):
    r = 39 + k
    put(ws, (r, 14), f"=LARGE({COV('L')},{k})")
    idx = f"MATCH($N{r},{COV('L')},0)"
    put(ws, (r, 2), f'=IF($N{r}<1,"",INDEX({COV("C")},{idx}))', f(10, True), None, LFT, BOX)
    put(ws, (r, 3), f'=IF($N{r}<1,"",INDEX({COV("K")},{idx}))', f(9, True), None, C, BOX)
    for j, col in enumerate("EFG"):
        put(ws, (r, 4 + j), f'=IF($N{r}<1,"",INDEX({COV(col)},{idx}))', f(10), None, C, BOX)
    put(ws, (r, 7), f'=IF($N{r}<1,"",INDEX({COV("I")},{idx}))', f(10), None, C, BOX)
rk = "C40:C49"
ws.conditional_formatting.add(rk, CellIsRule(operator="equal", formula=['"HIGH RISK"'], fill=fill(GAP_FILL), font=Font(name=FONT, bold=True, color=GAP_FONT)))
ws.conditional_formatting.add(rk, CellIsRule(operator="equal", formula=['"SHORT"'], fill=fill(AMB_FILL), font=Font(name=FONT, bold=True, color=AMB_FONT)))
ws.conditional_formatting.add("D40:F49", FormulaRule(formula=['AND(ISNUMBER(D40),D40<2)'], font=Font(name=FONT, bold=True, color="C00000")))
put(ws, "B50", '=IF(N40<1,"No coverage risks – every needed skill has enough qualified people.","Fix HIGH RISK first: train a second person, then a trainer (level 4).")', f(9, color=MUTED, italic=True))

# legend box (right of risks)
section("I38", "Rating scale")
for i, (lvl, name, desc, bg, fg) in enumerate(LEVELS):
    r = 39 + i
    put(ws, (r, 9), lvl, f(11, True, fg), fill(bg), C, BOX)
    merge(ws, f"J{r}:M{r}", f"{name} – {desc}", font=f(9), al=LFT)
put(ws, (44, 9), "RED", f(9, True, GAP_FONT), fill("FF7C80"), C, BOX)
merge(ws, "J44:M44", "Below the level the person's role needs (a gap)", font=f(9), al=LFT)
for r in range(39, 45):
    ws.row_dimensions[r].height = 28
ws.freeze_panes = "A4"

# ------------------------------------------------------------------ Person Card
ws = S_CARD
title(ws, "Person Card", "Pick a name to see every skill they need, what they have, and who can train them. Good for 1-to-1 reviews – print this tab.")
nav(ws, "G1", "H1")
for col, w in zip("ABCDEFGHI", (2, 7, 22, 34, 8, 8, 13, 20, 38)):
    ws.column_dimensions[col].width = w
ws.column_dimensions["M"].hidden = True
put(ws, "B4", "Person ▸", f(12, True, NAVY), al=Alignment(horizontal="right", vertical="center"))
merge(ws, "C4:D4", SAMPLE[0][1], font=f(13, True), fl=INPUT, al=LFT, border=BOX)
ws["D4"].border = BOX
dvp = DataValidation(type="list", formula1=f"={MNAMES}", allow_blank=True, prompt="Pick a name from the Skill Matrix", promptTitle="Person")
ws.add_data_validation(dvp); dvp.add("C4")
ws.row_dimensions[4].height = 26
put(ws, "M1", f"=IFERROR(MATCH($C$4,{MNAMES},0),\"\")")
for col, lab in (("C", "Role"), ("D", "Shift · Line"), ("E", "% Ready"), ("G", "Gaps"), ("H", "Emp ID")):
    put(ws, f"{col}6", lab, f(9, True, MUTED), al=Alignment(horizontal="left", indent=1))
for col in "CDEFGH":
    for r in (6, 7):
        ws[f"{col}{r}"].fill = fill("EEF3F8")
put(ws, "C7", f'=IF($M$1="","",INDEX({MCOL("C")},$M$1))', f(12, True), al=LFT)
put(ws, "D7", f'=IF($M$1="","",INDEX({MCOL("D")},$M$1)&"  ·  "&INDEX({MCOL("E")},$M$1))', f(12, True), al=LFT)
ws.merge_cells("E6:F6"); ws.merge_cells("E7:F7")
put(ws, "E7", f'=IF($M$1="","",INDEX({MCOL("I")},$M$1))', f(16, True, NAVY), al=LFT, fmt="0%")
put(ws, "G7", f'=IF($M$1="","",INDEX({MCOL("H")},$M$1))', f(16, True, "C00000"), al=LFT)
put(ws, "H7", f'=IF($M$1="","",INDEX({MCOL("A")},$M$1))', f(12, True), al=LFT)
put(ws, "I7", '=IF($M$1="","Pick a name from the drop-down in C4","")', f(10, True, "C00000"))
ws.row_dimensions[7].height = 28
for i, h in enumerate(("ID", "Area", "Skill", "Needs", "Has", "Status", "Who can train", "How it's checked"), start=2):
    head(ws, 9, i, h)
ws.row_dimensions[9].height = 24
for k, sk in enumerate(SKILLS):
    r = 10 + k
    put(ws, (r, 2), sk["id"], f(9, True, "FFFFFF"), fill(sk["color"]), C, BOX)
    put(ws, (r, 3), sk["cat"], f(9), None, LFT, BOX)
    put(ws, (r, 4), sk["name"], f(10, True), None, LFT, BOX)
    put(ws, (r, 5), f'=IF($M$1="","",IFERROR(INDEX({REQ_GRID},MATCH($C$7,{REQ_ROLES},0),{k + 1})+0,0))', f(11, True), None, C, BOX)
    put(ws, (r, 6), f'=IF($M$1="","",INDEX({MGRID},$M$1,{k + 1})+0)', f(11, True), None, C, BOX)
    put(ws, (r, 7), f'=IF($M$1="","",IF(E{r}=0,"Not needed",IF(F{r}>=E{r},"✓ Met","GAP")))', f(10, True), None, C, BOX)
    put(ws, (r, 8), f'=IF(OR($M$1="",G{r}<>"GAP"),"",IFERROR(INDEX({MNAMES},MATCH(4,INDEX({MGRID},0,{k + 1}),0)),"No trainer yet"))', f(9), None, LFT, BOX)
    put(ws, (r, 9), sk["verify"], f(9, color=MUTED), None, LFT, BOX)
    ws.row_dimensions[r].height = 20
last = 9 + NS
ws.conditional_formatting.add(f"G10:G{last}", CellIsRule(operator="equal", formula=['"GAP"'], fill=fill(GAP_FILL), font=Font(name=FONT, bold=True, color=GAP_FONT)))
ws.conditional_formatting.add(f"G10:G{last}", CellIsRule(operator="equal", formula=['"✓ Met"'], fill=fill(OK_FILL), font=Font(name=FONT, bold=True, color=OK_FONT)))
ws.conditional_formatting.add(f"C10:I{last}", FormulaRule(formula=['$G10="Not needed"'], font=Font(name=FONT, color="A6A6A6")))
ws.conditional_formatting.add(f"D10:D{last}", FormulaRule(formula=['$G10="GAP"'], font=Font(name=FONT, bold=True, color=GAP_FONT)))
for lvl, _n, _d, bg, fg in LEVELS:
    ws.conditional_formatting.add(f"F10:F{last}", FormulaRule(formula=[f'AND($G10<>"Not needed",$G10<>"",F10={lvl})'], fill=fill(bg), font=Font(name=FONT, bold=True, color=fg)))
ws.auto_filter.ref = f"B9:I{last}"
put(ws, (last + 2, 2), "Tip: use the filter arrow on 'Status' to show only GAP rows. To change a level, go to the Skill Matrix tab.", f(9, color=MUTED, italic=True))
ws.freeze_panes = "A10"

# ------------------------------------------------------------------ Skills Library
ws = S_LIB
title(ws, "Skills Library", "What each skill means and how it is checked. Add your site SOP / work instruction number in the yellow column.")
nav(ws, "E1", "F1")
cols = [("ID", 7), ("Area", 22), ("Skill", 32), ("What 'Qualified' (level 3) looks like", 60), ("Training method", 20),
        ("How it is checked", 26), ("Refresher", 13), ("SOP / WI #", 16), ("Roles that need it", 10)]
for i, (h, w) in enumerate(cols, start=1):
    head(ws, 4, i, h)
    ws.column_dimensions[L(i)].width = w
ws.row_dimensions[4].height = 30
for k, sk in enumerate(SKILLS):
    r = 5 + k
    vals = [sk["id"], sk["cat"], sk["name"], sk["qual"], sk["method"], sk["verify"], sk["refresh"], None]
    for i, v in enumerate(vals, start=1):
        put(ws, (r, i), v, f(10, i in (1, 3)), None, LFT if i in (2, 3, 4, 5, 6) else C, BOX)
    put(ws, (r, 1), font=f(10, True, "FFFFFF"), fl=fill(sk["color"]))
    ws.cell(r, 8).fill = INPUT
    c = L(SC + k)
    put(ws, (r, 9), f"=COUNTIF('Role Requirements'!{c}${R0}:{c}${REQ_LAST},\">0\")", f(10), None, C, BOX)
    ws.row_dimensions[r].height = 42
ws.freeze_panes = "D5"
ws.auto_filter.ref = f"A4:I{4 + NS}"

# ------------------------------------------------------------------ Training Plan
ws = S_PLAN
title(ws, "Training Plan", "One line per gap. Pick the person and skill from the drop-downs. Level 3+ needs trainer AND supervisor sign-off (plus QA for allergen / CCP skills).")
nav(ws, "F1", "G1")
tcols = [("Employee", 22), ("Skill ID", 9), ("Skill (auto)", 32), ("Now", 7), ("Target", 7), ("Trainer", 18),
         ("Start", 12), ("Due", 12), ("Status", 16), ("Trainer signed", 12), ("Supervisor signed", 12), ("QA signed (if needed)", 12), ("Notes", 30)]
for i, (h, w) in enumerate(tcols, start=1):
    head(ws, 4, i, h)
    ws.column_dimensions[L(i)].width = w
ws.row_dimensions[4].height = 36
lib_ids = f"'Skills Library'!$A$5:$A${4 + NS}"
lib_names = f"'Skills Library'!$C$5:$C${4 + NS}"
N_PLAN = 150
for k in range(N_PLAN):
    r = 5 + k
    for i in range(1, 14):
        put(ws, (r, i), font=f(10), fl=INPUT if i != 3 else GREYF, al=LFT if i in (1, 3, 6, 13) else C, border=BOX)
    put(ws, (r, 3), f'=IF(B{r}="","",IFERROR(INDEX({lib_names},MATCH(B{r},{lib_ids},0)),"Unknown ID"))')
    for i in (7, 8, 10, 11, 12):
        ws.cell(r, i).number_format = "dd-mmm-yy"
ex = ["Sample A", "PK05", None, 2, 3, "Sample L", date(2026, 11, 9), date(2026, 11, 30), "In progress", None, None, None, "Example row – overwrite"]
for i, v in enumerate(ex, start=1):
    if i != 3 and v is not None:
        put(ws, (5, i), v, f(10, italic=True, color=MUTED))
for d, rng_ in ((DataValidation(type="list", formula1=f"={MNAMES}", allow_blank=True), f"A5:A{4 + N_PLAN}"),
                (DataValidation(type="list", formula1=f"={lib_ids}", allow_blank=True), f"B5:B{4 + N_PLAN}"),
                (DataValidation(type="list", formula1='"0,1,2,3,4"', allow_blank=True), f"D5:E{4 + N_PLAN}"),
                (DataValidation(type="list", formula1=f"={MNAMES}", allow_blank=True), f"F5:F{4 + N_PLAN}"),
                (DataValidation(type="list", formula1="=Lists!$G$5:$G$9", allow_blank=True), f"I5:I{4 + N_PLAN}")):
    ws.add_data_validation(d); d.add(rng_)
st = f"I5:I{4 + N_PLAN}"
for val, bg, fg in (("Complete", OK_FILL, OK_FONT), ("In progress", AMB_FILL, AMB_FONT), ("Ready for sign-off", "DDEBF7", NAVY), ("On hold", "D9D9D9", "404040")):
    ws.conditional_formatting.add(st, CellIsRule(operator="equal", formula=[f'"{val}"'], fill=fill(bg), font=Font(name=FONT, bold=True, color=fg)))
ws.conditional_formatting.add(f"H5:H{4 + N_PLAN}", FormulaRule(formula=['AND(ISNUMBER(H5),H5<TODAY(),$I5<>"Complete")'], fill=fill(GAP_FILL), font=Font(name=FONT, bold=True, color=GAP_FONT)))
ws.freeze_panes = "B5"
ws.auto_filter.ref = f"A4:M{4 + N_PLAN}"

# ------------------------------------------------------------------ Start Here
ws = S_START
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 3
for c, w in zip("BCDEFG", (30, 3, 30, 3, 30, 3)):
    ws.column_dimensions[c].width = w
merge(ws, "B2:F2", "Frozen Production Skill Matrix", font=f(24, True, NAVY), al=Alignment(vertical="center"))
ws.row_dimensions[2].height = 40
merge(ws, "B3:F3", "C.H. Guenther · production team · frozen operations only (batching, mixing, forming, freezing, packaging, palletizing). No baking.",
      font=f(10, color=MUTED, italic=True), al=Alignment(wrap_text=True, vertical="top"))
ws.row_dimensions[3].height = 28
put(ws, "B5", "What do you want to do?  (click a box)", f(13, True, NAVY))
buttons = [
    ("See how the team is doing", "Dashboard", "2E75B6"),
    ("Update skill levels", "Skill Matrix", "548235"),
    ("Review one person", "Person Card", "7030A0"),
    ("Check shift coverage", "Shift Coverage", "BF8F00"),
    ("Log training & sign-offs", "Training Plan", "C55A11"),
    ("Look up what a skill means", "Skills Library", "404040"),
]
for i, (text, target, color) in enumerate(buttons):
    r = 7 + (i // 3) * 3
    col = "BDF"[i % 3]
    ws.merge_cells(f"{col}{r}:{col}{r + 1}")
    link(ws, f"{col}{r}", f"{text}\n▸ {target}", target, "FFFFFF", color, 12)
for r in (7, 8, 10, 11):
    ws.row_dimensions[r].height = 24
put(ws, "B14", "Getting started – 3 steps", f(13, True, NAVY))
steps = [
    ("1", "Enter your people", "On the Skill Matrix, type over the 14 'Sample' rows (or delete them) with each person's ID, name, role and shift. Yellow cells = you type here."),
    ("2", "Pick their levels", "In each skill cell choose 0–4 from the drop-down. Cells turn RED when someone is below what their role needs."),
    ("3", "Act on the gaps", "Dashboard shows where you stand. Person Card shows each person's gaps and who can train them. Log the training on Training Plan."),
]
for i, (n, h, t) in enumerate(steps):
    r = 15 + i
    put(ws, (r, 2), f"{n}.  {h}", f(11, True), al=Alignment(vertical="top", indent=1))
    merge(ws, f"D{r}:F{r}", t, font=f(10), al=Alignment(wrap_text=True, vertical="top"))
    ws.row_dimensions[r].height = 34
put(ws, "B19", "Rating scale", f(13, True, NAVY))
for i, (lvl, name, desc, bg, fg) in enumerate(LEVELS):
    r = 20 + i
    put(ws, (r, 2), f"{lvl}  {name}", f(11, True, fg), fill(bg), Alignment(horizontal="left", vertical="center", indent=1), BOX)
    merge(ws, f"D{r}:F{r}", desc, font=f(10), al=Alignment(vertical="center", wrap_text=True))
    ws.row_dimensions[r].height = 22
put(ws, "B25", "RED cell", f(11, True, GAP_FONT), fill("FF7C80"), Alignment(horizontal="left", vertical="center", indent=1), BOX)
merge(ws, "D25:F25", "Below the level that person's role needs = a gap to train.", font=f(10), al=Alignment(vertical="center"))
ws.row_dimensions[25].height = 22
put(ws, "B27", "Sign-off rules", f(13, True, NAVY))
rules = [
    "Only enter 3 after training, observed performance, and trainer + supervisor signatures.",
    "4 (Trainer) needs 6+ months at level 3, the On-the-job trainer skill (LL05) and supervisor approval.",
    "Allergen and CCP skills (FS02, FS03, FS05, FR03, PK03, PK04) also need QA verification.",
    "Review monthly, and after every new hire, transfer, new product or new equipment.",
    "Signed training records stay in document control – this workbook is the summary.",
]
for i, t in enumerate(rules):
    merge(ws, f"B{28 + i}:F{28 + i}", "•  " + t, font=f(10), al=Alignment(vertical="center"))
merge(ws, "B34:F34", "Role requirements can be changed on the Role Requirements tab. Shift names can be renamed on the Lists tab.",
      font=f(9, color=MUTED, italic=True))

# ------------------------------------------------------------------ finishing
for s in wb.worksheets:
    s.sheet_properties.pageSetUpPr.fitToPage = True
    s.page_setup.orientation = "landscape"
    s.page_setup.fitToWidth = 1
    s.page_setup.fitToHeight = 0
S_CARD.page_setup.orientation = "portrait"
S_MX.print_title_rows = "3:5"
S_MX.print_title_cols = "A:C"
for s, color in ((S_START, NAVY), (S_DASH, "2E75B6"), (S_MX, "548235"), (S_CARD, "7030A0"), (S_COV, "BF8F00"),
                 (S_PLAN, "C55A11"), (S_LIB, "404040"), (S_REQ, "7F7F7F"), (S_LIST, "BFBFBF")):
    s.sheet_properties.tabColor = color
wb.active = 0
wb.save(OUT)
print(f"Saved {OUT} ({NS} skills, {len(ROLES)} roles)")
