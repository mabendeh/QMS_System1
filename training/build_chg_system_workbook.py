"""Builds the CHG Production Skills workbook (simplified edition).

    python training/build_chg_system_workbook.py

Output: training/CHG_Skills_Career_System.xlsx
Sheets: Start Here, Dashboard, Skills Matrix, Employees, Person Card, CI Analytics,
        Job Requirements (+ hidden Lists, Calc, SkillCalc helpers).
Departments, competencies and job titles come from chg_config.json.
All people are fictional samples.
"""
import json
import random
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

HERE = Path(__file__).parent
CFG = json.loads((HERE / "chg_config.json").read_text())
OUT = HERE / "CHG_Skills_Career_System.xlsx"
TODAY = date.today()
FONT = "Arial"

AREAS = [a for a in CFG["AREAS"] if a["active"]]          # Proofing & Baking left out (frozen only)
DEPTS = [a["name"] for a in AREAS]
ND = len(DEPTS)
IS_PROD = [1 if a["kind"] == "PRODUCTION" else 0 for a in AREAS]
PROD_DEPTS = [a["name"] for a in AREAS if a["kind"] == "PRODUCTION"]
SKILLS = [(f'{a["code"]}-{i + 1:02d}', a["name"], c[0], len(c) > 2 and c[2] == "critical") for a in AREAS for i, c in enumerate(a["competencies"])]
NS = len(SKILLS)
JOBS = CFG["JOB_ROLES"]
SHIFTS = ["A", "B", "C"]

E0, E_N = 6, 150
E1 = E0 + E_N - 1
MC = 7
MCL, MLAST = L(MC), L(MC + NS - 1)
JR0 = 6
JR_N = 15
SC0 = 6                                   # SkillCalc first row
SC1 = SC0 + NS - 1

NAVY, INK, MUTED = "14284B", "1A2433", "5D6B7E"
ST = {"NT": ("C62828", "FFFFFF", "Needs Training"), "IT": ("F2C230", "1A2433", "In Training"), "TR": ("2E8B57", "FFFFFF", "Trained"),
      "C": ("1E63C4", "FFFFFF", "Competent"), "N/A": ("E3E7EC", "5D6B7E", "Not Applicable")}
REQ_BLANK = ("F4B6B6", "C62828")
OK_F, OK_T, AMB_F, AMB_T, BAD_F, BAD_T = "C6EFCE", "006100", "FFEB9C", "7F6000", "FFC7CE", "9C0006"
thin = Side(style="thin", color="D0D7DE")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
C_AL = Alignment(horizontal="center", vertical="center", wrap_text=True)
L_AL = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
INPUT = PatternFill("solid", fgColor="FFF9DB")
GREY = PatternFill("solid", fgColor="F2F4F6")


def f(size=10, bold=False, color=INK, italic=False, u=None):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic, underline=u)


def fill(c):
    return PatternFill("solid", fgColor=c)


def put(ws, ref, v=None, font=None, fl=None, al=None, border=None, fmt=None):
    c = ws[ref] if isinstance(ref, str) else ws.cell(*ref)
    if v is not None:
        c.value = v
    if font: c.font = font
    if fl: c.fill = fl
    if al: c.alignment = al
    if border: c.border = border
    if fmt: c.number_format = fmt
    return c


def head(ws, r, c, text, color=NAVY, rotate=0, size=10):
    cell = ws.cell(r, c, text)
    cell.font = f(size, True, "FFFFFF")
    cell.fill = fill(color)
    cell.border = BOX
    cell.alignment = Alignment(horizontal="center", vertical="bottom" if rotate else "center", wrap_text=True, text_rotation=rotate)
    return cell


def merge(ws, rng, v=None, **kw):
    ws.merge_cells(rng)
    return put(ws, rng.split(":")[0], v, **kw)


def link(ws, ref, text, target, color=NAVY, bg=None, size=10):
    c = put(ws, ref, text, f(size, True, color, u=None if bg else "single"), fill(bg) if bg else None, C_AL)
    c.hyperlink = f"#'{target}'!A1"
    return c


def title(ws, text, sub, navcol="H"):
    put(ws, "A1", text, f(18, True, NAVY))
    put(ws, "A2", sub, f(10, color=MUTED, italic=True))
    ws.row_dimensions[1].height = 30
    ws.sheet_view.showGridLines = False
    link(ws, f"{navcol}1", "◂ Start Here", "Start Here")
    link(ws, f"{L(ws[navcol + '1'].column + 1)}1", "▦ Dashboard", "Dashboard")


def dv(ws, formula, rng, prompt=None, title_=None, allow_blank=True):
    d = DataValidation(type="list", formula1=formula, allow_blank=allow_blank, prompt=prompt, promptTitle=title_,
                       showErrorMessage=True, error="Pick a value from the list.", errorTitle="Not in list")
    ws.add_data_validation(d)
    d.add(rng)


def cf(ws, rng, formula, fillc, fontc, bold=True, stop=False):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[formula], fill=fill(fillc), font=Font(name=FONT, bold=bold, color=fontc), stopIfTrue=stop))


wb = Workbook()
S = {}
for name in ["Start Here", "Dashboard", "Skills Matrix", "Employees", "Person Card", "CI Analytics", "Job Requirements", "Lists", "Calc", "SkillCalc"]:
    S[name] = wb.active if name == "Start Here" else wb.create_sheet(name)
S["Start Here"].title = "Start Here"
MX, EMP, JR, CALC, SKC = "'Skills Matrix'", "Employees", "'Job Requirements'", "Calc", "SkillCalc"

# ================================================================== Lists (hidden)
ws = S["Lists"]
lists = {"A": ("Shift", SHIFTS), "B": ("Skill status", list(ST)), "C": ("Department", DEPTS), "D": ("Yes/No", ["Yes", "No"]),
         "E": ("Dashboard filter", ["All shifts"] + SHIFTS), "F": ("Supervisors (edit)", ["Supervisor A (sample)", "Supervisor B (sample)", "Supervisor C (sample)"])}
for col, (h, vals) in lists.items():
    put(ws, f"{col}4", h, f(10, True))
    ws.column_dimensions[col].width = 24
    for i, v in enumerate(vals):
        put(ws, f"{col}{5 + i}", v)
LISTS = {k: f"Lists!${k}$5:${k}${4 + len(v[1])}" for k, v in lists.items()}
LISTS["F"] = "Lists!$F$5:$F$34"
ws.sheet_state = "hidden"

# ================================================================== Job Requirements
ws = S["Job Requirements"]
title(ws, "Job Requirements", "Mark “Yes” for every department a job title must be qualified in. Unassessed required skills show light red on the matrix.", navcol="D")
ws.column_dimensions["A"].width = 2
ws.column_dimensions["B"].width = 26
head(ws, 5, 2, "Job title")
for d, name in enumerate(DEPTS):
    head(ws, 5, 3 + d, name, rotate=90)
    ws.column_dimensions[L(3 + d)].width = 6
ws.row_dimensions[5].height = 150
JLAST = L(2 + ND)
for j in range(JR_N):
    r = JR0 + j
    job = JOBS[j] if j < len(JOBS) else None
    put(ws, (r, 2), job["title"] if job else None, f(10, True), INPUT, L_AL, BOX)
    for d, a in enumerate(AREAS):
        put(ws, (r, 3 + d), "Yes" if job and a["code"] in job["areas"] else None, f(10, True), INPUT, C_AL, BOX)
JR_GRID = f"{JR}!$C${JR0}:${JLAST}${JR0 + JR_N - 1}"
JR_JOBS = f"{JR}!$B${JR0}:$B${JR0 + JR_N - 1}"
JR_DEPTS = f"{JR}!$C$5:${JLAST}$5"
dv(ws, f"={LISTS['D']}", f"C{JR0}:{JLAST}{JR0 + JR_N - 1}")
ws.conditional_formatting.add(f"C{JR0}:{JLAST}{JR0 + JR_N - 1}", CellIsRule(operator="equal", formula=['"Yes"'], fill=fill("1E63C4"), font=Font(name=FONT, bold=True, color="FFFFFF")))
put(ws, (JR0 + JR_N + 1, 2), "Add new job titles in the empty yellow rows. Department columns are fixed.", f(9, italic=True, color=MUTED))
ws.freeze_panes = "C6"

# ================================================================== sample people (fictional)
rnd = random.Random(20261109)
first = ["Alex", "Jordan", "Taylor", "Casey", "Morgan", "Riley", "Jamie", "Avery", "Quinn", "Reese", "Skyler", "Dakota", "Rowan", "Emerson", "Finley", "Harper",
         "Kendall", "Logan", "Parker", "Sawyer", "Blake", "Cameron", "Drew", "Elliot", "Hayden", "Jesse", "Kai", "Lane", "Marley", "Noel", "Oakley", "Peyton",
         "Remy", "Sage", "Tatum", "Wren", "Ari", "Bailey"]
last = ["Sample", "Demo", "Example", "Testcase", "Placeholder", "Fictional", "Mockley", "Trialson"]
plan = []
for s in range(3):
    for role, n in [("Pre-Weigh Operator", 2), ("Mixer Operator", 2), ("Blending Operator", 1), ("Forming Operator", 1), ("Freezer Operator", 1),
                    ("Packaging Operator", 3), ("Palletizer Operator", 1), ("Line Lead", 1)]:
        plan += [(role, s)] * n
plan += [("Production Associate", 2), ("Sanitation Technician", 1)]
job_areas = {j["title"]: j["areas"] for j in JOBS}
cross = {"PW": "MX", "MX": "BL", "BL": "PW", "PK": "PL", "PL": "PK", "DF": "SF", "SF": "DF", "LS": "PK"}
code_name = {a["code"]: a["name"] for a in AREAS}
by_code = {}
for k, s in enumerate(SKILLS):
    by_code.setdefault(s[0].split("-")[0], []).append(k)
people = []
for i, (role, s) in enumerate(plan):
    primary = CFG["PRIMARY_AREA_FOR_ROLE"][role]
    trainee = role == "Production Associate" or i % 13 == 7
    p = {"id": f"DEMO-{1001 + i}", "first": first[i % len(first)], "last": last[i % len(last)], "shift": SHIFTS[s], "dept": code_name[primary], "job": role,
         "sup": lists["F"][1][s], "hire": TODAY - timedelta(days=int((rnd.uniform(1, 3) if trainee else rnd.uniform(6, 60)) * 30.4)), "st": {}}
    areas = list(job_areas[role])
    if not trainee and rnd.random() < 0.35 and primary in cross:
        areas.append(cross[primary])
    for code in areas:
        for k in by_code.get(code, []):
            x = rnd.random()
            if trainee:
                st = ("C" if x < .6 else "IT") if code in ("FS", "SA") else ("TR" if x < .15 else "IT" if x < .7 else "NT" if x < .85 else "")
            else:
                st = "C" if x < .9 else "TR" if x < .94 else "IT" if x < .98 else ""
            if st:
                p["st"][k] = st
    people.append(p)


def find(role, shift, nth=0):
    return [p for p in people if p["job"] == role and p["shift"] == shift][nth]


def set_area(p, code, st):
    for k in by_code[code]:
        p["st"][k] = st


m = find("Pre-Weigh Operator", "A")
for c in ("PW", "MX", "FS", "SA"):
    set_area(m, c, "C")
m.update(first="Alex", last="Multiskill-Sample")
t = find("Mixer Operator", "A", 1)
set_area(t, "MX", "IT")
t.update(first="Jordan", last="Trainee-Sample")
pk = find("Packaging Operator", "C")
for c in ("PK", "MD", "PL", "FS", "SA"):
    set_area(pk, c, "C")
pk.update(first="Morgan", last="Crosstrained-Sample")
na = find("Freezer Operator", "A")
for c in ("SF", "FS", "SA"):
    set_area(na, c, "C")
na["st"][by_code["SF"][-1]] = "N/A"          # example: one skill not applicable to this person
na.update(first="Riley", last="NA-Example-Sample")

# ================================================================== Calc layout (used by several sheets)
BL = ["C", "N", "T", "R", "Q"]           # competent, not-applicable, in-training, required(0/1), qualified(0/1)
BS = {b: 6 + i * ND for i, b in enumerate(BL)}
SUMC = 6 + len(BL) * ND + 1
SUMS = ["req", "comp", "gaps", "tr", "prod", "full", "key"]
SC = {n: SUMC + i for i, n in enumerate(SUMS)}
CL = {n: L(c) for n, c in SC.items()}

# ================================================================== Employees
ws = S["Employees"]
title(ws, "Employees", "One row per person. Type in the yellow cells; grey cells calculate. Rows marked Sample are fictional – overwrite or delete them.", navcol="K")
ecols = [("Employee ID", 11), ("First name", 13), ("Last name", 18), ("Shift", 7), ("Department", 20), ("Job title", 20), ("Supervisor", 20), ("Hire date", 11),
         ("Sample?", 8), ("Full name", 24), ("Skills required", 9), ("Competent", 10), ("Gaps", 7), ("% Ready", 8), ("In training (IT/TR)", 10),
         ("Production depts qualified", 11), ("Cross-trained", 9), ("Summary", 22)]
for i, (h, w) in enumerate(ecols, start=1):
    head(ws, 5, i, h, NAVY if i <= 9 else "404040")
    ws.column_dimensions[L(i)].width = w
ws.row_dimensions[5].height = 44
for k in range(E_N):
    r = E0 + k
    p = people[k] if k < len(people) else None
    vals = [p["id"], p["first"], p["last"], p["shift"], p["dept"], p["job"], p["sup"], p["hire"], "Yes"] if p else [None] * 9
    for c, v in enumerate(vals, start=1):
        put(ws, (r, c), v, f(10), INPUT, L_AL if c in (2, 3, 5, 6, 7) else C_AL, BOX, "dd-mmm-yy" if c == 8 else None)
    fm = [f'=IF(A{r}="","",TRIM(B{r}&" "&C{r}))', f'=IF(A{r}="","",{CALC}!{CL["req"]}{r})', f'=IF(A{r}="","",{CALC}!{CL["comp"]}{r})',
          f'=IF(A{r}="","",{CALC}!{CL["gaps"]}{r})', f'=IF(OR(A{r}="",N(K{r})=0),"",L{r}/K{r})', f'=IF(A{r}="","",{CALC}!{CL["tr"]}{r})',
          f'=IF(A{r}="","",{CALC}!{CL["prod"]}{r})', f'=IF(A{r}="","",IF(P{r}>=2,"Yes",""))',
          f'=IF(A{r}="","",IF(F{r}="","Choose a job title",IF(ISNA(MATCH(F{r},{JR_JOBS},0)),"Job title not on Job Requirements",IF(K{r}=0,"No skills required",IF(M{r}=0,"Fully qualified",M{r}&" skill gap"&IF(M{r}=1,"","s"))))))']
    for c, x in enumerate(fm, start=10):
        put(ws, (r, c), x, f(10, c in (10, 13, 14)), GREY, L_AL if c in (10, 18) else C_AL, BOX, "0%" if c == 14 else None)
dv(ws, f"={LISTS['A']}", f"D{E0}:D{E1}")
dv(ws, f"={LISTS['C']}", f"E{E0}:E{E1}")
dv(ws, f"={JR_JOBS}", f"F{E0}:F{E1}")
dv(ws, f"={LISTS['F']}", f"G{E0}:G{E1}")
dv(ws, f"={LISTS['D']}", f"I{E0}:I{E1}")
cf(ws, f"N{E0}:N{E1}", f"AND(ISNUMBER(N{E0}),N{E0}>=1)", OK_F, OK_T, stop=True)
cf(ws, f"N{E0}:N{E1}", f"AND(ISNUMBER(N{E0}),N{E0}>=0.8)", AMB_F, AMB_T, stop=True)
cf(ws, f"N{E0}:N{E1}", f"ISNUMBER(N{E0})", BAD_F, BAD_T)
cf(ws, f"M{E0}:M{E1}", f"AND(ISNUMBER(M{E0}),M{E0}>0)", "FFFFFF", "C00000")
cf(ws, f"I{E0}:I{E1}", f'I{E0}="Yes"', AMB_F, AMB_T)
cf(ws, f"A{E0}:A{E1}", f'AND(A{E0}<>"",COUNTIF($A${E0}:$A${E1},A{E0})>1)', "FF7C80", "FFFFFF")
put(ws, "M2", "Red Employee ID = duplicate ID.", f(9, italic=True, color="C00000"))
ws.freeze_panes = "D6"
ws.auto_filter.ref = f"A5:R{E1}"
EMP_IDS = f"{EMP}!$A${E0}:$A${E1}"
EMP_NAMES = f"{EMP}!$J${E0}:$J${E1}"
ECOL = lambda c: f"{EMP}!${c}${E0}:${c}${E1}"  # noqa: E731

# ================================================================== Skills Matrix
ws = S["Skills Matrix"]
title(ws, "Production Skills Matrix", "Pick a status in each cell: NT Needs Training · IT In Training · TR Trained · C Competent · N/A Not Applicable. Light red empty cell = required for the job, not assessed yet.", navcol="D")
ws.row_dimensions[2].height = 10
for i, (h, w) in enumerate([("Employee ID", 11), ("Employee", 24), ("Shift", 6), ("Department", 17), ("Job title", 19), ("% Ready", 8)], start=1):
    ws.merge_cells(start_row=3, start_column=i, end_row=5, end_column=i)
    head(ws, 3, i, h, NAVY if i < 6 else "404040")
    ws.column_dimensions[L(i)].width = w
for k, (sid, dept, name, crit) in enumerate(SKILLS):
    c = MC + k
    put(ws, (2, c), dept, f(6, color="FFFFFF"))
    head(ws, 3, c, None, "2E4A7A", size=9)
    head(ws, 4, c, sid, "2E4A7A", size=8)
    head(ws, 5, c, ("◆ " if crit else "") + name, "2E4A7A", rotate=90, size=9)
    ws.column_dimensions[L(c)].width = 4.6
start = MC
for a in AREAS:
    n = len(a["competencies"])
    ws.merge_cells(start_row=3, start_column=start, end_row=3, end_column=start + n - 1)
    ws.cell(3, start, a["name"]).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.column_dimensions.group(L(start), L(start + n - 1), outline_level=1, hidden=False)
    start += n
ws.sheet_properties.outlinePr.summaryRight = False
ws.row_dimensions[3].height = 26
ws.row_dimensions[5].height = 210
for k in range(E_N):
    r = E0 + k
    for c, src in zip("ABCDE", "AJDEF"):
        put(ws, f"{c}{r}", f'=IF({EMP}!$A{r}="","",{EMP}!{src}{r})', f(10, c == "B"), None, L_AL if c in "BDE" else C_AL, BOX)
    put(ws, f"F{r}", f'=IF({EMP}!$A{r}="","",{EMP}!N{r})', f(10, True), GREY, C_AL, BOX, "0%")
    p = people[k] if k < len(people) else None
    for j in range(NS):
        cell = ws.cell(r, MC + j, p["st"].get(j) if p else None)
        cell.font = f(9, True)
        cell.alignment = C_AL
        cell.border = BOX
    ws.row_dimensions[r].height = 18
grid = f"{MCL}{E0}:{MLAST}{E1}"
tl = f"{MCL}{E0}"
for code, (fc, tc, _) in ST.items():
    cf(ws, grid, f'{tl}="{code}"', fc, tc, stop=True)
req = f'IFERROR(INDEX({JR_GRID},MATCH($E{E0},{JR_JOBS},0),MATCH({MCL}$2,{JR_DEPTS},0)),"")="Yes"'
cf(ws, grid, f'AND({tl}="",$A{E0}<>"",{req})', REQ_BLANK[0], REQ_BLANK[1])
cf(ws, f"F{E0}:F{E1}", f"AND(ISNUMBER(F{E0}),F{E0}>=1)", OK_F, OK_T, stop=True)
cf(ws, f"F{E0}:F{E1}", f"AND(ISNUMBER(F{E0}),F{E0}>=0.8)", AMB_F, AMB_T, stop=True)
cf(ws, f"F{E0}:F{E1}", f"ISNUMBER(F{E0})", BAD_F, BAD_T)
dv(ws, f"={LISTS['B']}", grid, "NT Needs Training | IT In Training | TR Trained | C Competent | N/A Not Applicable", "Skill status")
tr = E1 + 2
for i, (lab, fm) in enumerate((("# Competent (C)", '"C"'), ("# In training (IT + TR)", None), ("# Needs training (NT)", '"NT"'))):
    r = tr + i
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=MC - 1)
    put(ws, (r, 1), lab, f(10, True), al=Alignment(horizontal="right"))
    for j in range(NS):
        col = L(MC + j)
        rng = f"{col}{E0}:{col}{E1}"
        x = f'=COUNTIF({rng},"IT")+COUNTIF({rng},"TR")' if fm is None else f"=COUNTIF({rng},{fm})"
        put(ws, (r, MC + j), x, f(9, True), GREY, C_AL, BOX)
ws.freeze_panes = ws.cell(E0, MC)
ws.auto_filter.ref = f"A5:F{E1}"
MDEPT = f"{MX}!${MCL}$2:${MLAST}$2"
MGRID = f"{MX}!${MCL}${E0}:${MLAST}${E1}"

# ================================================================== Calc (hidden)
ws = S["Calc"]
put(ws, "A1", "Helper calculations – do not edit.", f(10, True))
put(ws, "A2", "Shift filter")
put(ws, "B2", '=IF(Dashboard!$K$2="All shifts","*",Dashboard!$K$2)')
for b in BL:
    put(ws, (2, BS[b]), b, f(9, True))
    for d, name in enumerate(DEPTS):
        c = BS[b] + d
        put(ws, (3, c), name)
        put(ws, (4, c), f'=COUNTIF({MDEPT},{L(c)}$3)')
        put(ws, (5, c), IS_PROD[d])
for n, c in SC.items():
    put(ws, (5, c), n, f(9, True))


def bc(b, d, r):
    return f"{L(BS[b] + d)}{r}"


def brow(b, r, absrow=False):
    a, z = L(BS[b]), L(BS[b] + ND - 1)
    return f"${a}${r}:${z}${r}" if absrow else f"{a}{r}:{z}{r}"


for k in range(E_N):
    r = E0 + k
    put(ws, (r, 1), f'=IF({EMP}!A{r}="","",{EMP}!A{r})')
    put(ws, (r, 2), f'=IF(A{r}="","",{EMP}!D{r})')
    put(ws, (r, 3), f'=IF(A{r}="","",{EMP}!F{r})')
    put(ws, (r, 4), f'=IF(A{r}="","",{EMP}!E{r})')
    mrow = f"{MX}!${MCL}{r}:${MLAST}{r}"
    for d in range(ND):
        h = lambda b: f"{L(BS[b] + d)}$3"  # noqa: E731
        put(ws, (r, BS["C"] + d), f'=IF($A{r}="",0,COUNTIFS({MDEPT},{h("C")},{mrow},"C"))')
        put(ws, (r, BS["N"] + d), f'=IF($A{r}="",0,COUNTIFS({MDEPT},{h("N")},{mrow},"N/A"))')
        put(ws, (r, BS["T"] + d), f'=IF($A{r}="",0,COUNTIFS({MDEPT},{h("T")},{mrow},"IT")+COUNTIFS({MDEPT},{h("T")},{mrow},"TR"))')
        put(ws, (r, BS["R"] + d), f'=IF($A{r}="",0,IF(IFERROR(INDEX({JR_GRID},MATCH($C{r},{JR_JOBS},0),MATCH({h("R")},{JR_DEPTS},0)),"")="Yes",1,0))')
        tot = f"{L(BS['Q'] + d)}$4"
        put(ws, (r, BS["Q"] + d), f'=IF(AND({tot}>0,{bc("C", d, r)}>0,{bc("C", d, r)}+{bc("N", d, r)}>={tot}),1,0)')
    put(ws, (r, SC["req"]), f'=SUMPRODUCT({brow("R", r)},{brow("C", 4, True)})-SUMPRODUCT({brow("R", r)},{brow("N", r)})')
    put(ws, (r, SC["comp"]), f'=SUMPRODUCT({brow("R", r)},{brow("C", r)})')
    put(ws, (r, SC["gaps"]), f'={CL["req"]}{r}-{CL["comp"]}{r}')
    put(ws, (r, SC["tr"]), f'=SUMPRODUCT({brow("R", r)},{brow("T", r)})')
    put(ws, (r, SC["prod"]), f'=SUMPRODUCT({brow("Q", r)},{brow("Q", 5, True)})')
    put(ws, (r, SC["full"]), f'=AND($A{r}<>"",{CL["req"]}{r}>0,{CL["gaps"]}{r}=0)')
    put(ws, (r, SC["key"]), f'=IF(AND($A{r}<>"",OR($B$2="*",$B{r}=Dashboard!$K$2),{CL["gaps"]}{r}>0),{CL["gaps"]}{r}+(1000-ROW())/100000,0)')
ws.sheet_state = "hidden"
CA = {n: f"{CALC}!${L(c)}${E0}:${L(c)}${E1}" for n, c in SC.items()}
CSHIFT = f"{CALC}!$B${E0}:$B${E1}"
CR = f"{CALC}!$B$2"


def blockcol(b, dept_cell):
    """Whole-column range of block b for the department named in dept_cell."""
    a, z = L(BS[b]), L(BS[b] + ND - 1)
    return f"INDEX({CALC}!${a}${E0}:${z}${E1},0,MATCH({dept_cell},{CALC}!${a}$3:${z}$3,0))"


# ================================================================== SkillCalc (hidden)
ws = S["SkillCalc"]
for i, h in enumerate(["Skill ID", "Department", "Competency", "Competent", "Only person", "single key", "People who need it", "Nobody competent"], start=1):
    put(ws, (5, i), h, f(9, True))
for k, (sid, dept, name, crit) in enumerate(SKILLS):
    r = SC0 + k
    col = L(MC + k)
    put(ws, (r, 1), sid)
    put(ws, (r, 2), dept)
    put(ws, (r, 3), name)
    put(ws, (r, 4), f'=COUNTIF({MX}!{col}${E0}:{col}${E1},"C")')
    put(ws, (r, 5), f'=IF(D{r}=1,INDEX({MX}!$B${E0}:$B${E1},MATCH("C",{MX}!{col}${E0}:{col}${E1},0)),"")')
    put(ws, (r, 6), f'=IF(D{r}=1,1+(1000-ROW())/100000,0)')
    put(ws, (r, 7), f'=SUMPRODUCT({blockcol("R", f"B{r}")}*({MX}!{col}${E0}:{col}${E1}<>"N/A"))')
    put(ws, (r, 8), f'=IF(AND(D{r}=0,G{r}>0),1,0)')
ws.sheet_state = "hidden"
SK = lambda c: f"{SKC}!${c}${SC0}:${c}${SC1}"  # noqa: E731

# ================================================================== Dashboard
ws = S["Dashboard"]
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 2
ws.column_dimensions["B"].width = 26
for c in "CDEFGHIJKLMN":
    ws.column_dimensions[c].width = 11.5
ws.column_dimensions["P"].hidden = True
merge(ws, "B2:H2", "Production Skills Dashboard", font=f(20, True, NAVY), al=Alignment(vertical="center"))
ws.row_dimensions[2].height = 34
put(ws, "B3", "C.H. Guenther · Duncanville · updates automatically from the Skills Matrix · sample people are fictional", f(10, color=MUTED, italic=True))
put(ws, "J2", "Show shift ▸", f(11, True, NAVY), al=Alignment(horizontal="right", vertical="center"))
merge(ws, "K2:L2", "All shifts", font=f(12, True), fl=INPUT, al=C_AL, border=BOX)
dv(ws, f"={LISTS['E']}", "K2", "Pick a shift to filter the dashboard", "Shift", allow_blank=False)
link(ws, "M2", "◂ Start", "Start Here")
SHF = ECOL("D")
tiles = [
    ("B", "C", "EMPLOYEES", f'=COUNTIFS({EMP_IDS},"<>",{SHF},{CR})', "0", '="on the roster"'),
    ("D", "E", "TEAM READINESS", f'=IFERROR(SUMIFS({ECOL("L")},{SHF},{CR})/SUMIFS({ECOL("K")},{SHF},{CR}),0)', "0%",
     f'=SUMIFS({ECOL("L")},{SHF},{CR})&" of "&SUMIFS({ECOL("K")},{SHF},{CR})&" required skills"'),
    ("F", "G", "FULLY QUALIFIED", f'=COUNTIFS({CA["full"]},TRUE,{CSHIFT},{CR})', "0", '="no skill gaps"'),
    ("H", "I", "NEED TRAINING", f'=COUNTIFS({ECOL("M")},">0",{SHF},{CR})', "0", '="have a required gap"'),
    ("J", "K", "IN TRAINING", f'=COUNTIFS({ECOL("O")},">0",{SHF},{CR})', "0", '="have IT / TR skills"'),
    ("L", "M", "CROSS-TRAINED", f'=COUNTIFS({ECOL("Q")},"Yes",{SHF},{CR})', "0", '="2+ production depts"'),
]
tf = fill("EEF3F8")
for a, b, lab, fm, fmt, sub in tiles:
    for rr in range(5, 9):
        for cc in (a, b):
            ws[f"{cc}{rr}"].fill = tf
    merge(ws, f"{a}5:{b}5", lab, font=f(9, True, MUTED), al=Alignment(horizontal="left", vertical="bottom", indent=1))
    merge(ws, f"{a}6:{b}7", fm, font=f(28, True, NAVY), al=Alignment(horizontal="left", vertical="center", indent=1), fmt=fmt)
    merge(ws, f"{a}8:{b}8", sub, font=f(9, color=MUTED), al=Alignment(horizontal="left", vertical="top", indent=1))
for rr in (6, 7):
    ws.row_dimensions[rr].height = 24
ws.conditional_formatting.add("D6", CellIsRule(operator="greaterThanOrEqual", formula=["0.95"], font=Font(name=FONT, size=28, bold=True, color=OK_T)))
ws.conditional_formatting.add("D6", CellIsRule(operator="lessThan", formula=["0.8"], font=Font(name=FONT, size=28, bold=True, color="C00000")))
put(ws, "B10", "Qualified people by department", f(13, True, NAVY))
put(ws, "B11", "Qualified = every skill in the department is C (or N/A for that person)", f(9, italic=True, color=MUTED))
for c, h in zip("BCDEFG", ("Department", "A", "B", "C", "Selected shift", "Assigned to dept")):
    head(ws, 12, ws[c + "12"].column, h)
NP = len(PROD_DEPTS)
for i, dname in enumerate(PROD_DEPTS):
    r = 13 + i
    put(ws, f"B{r}", dname, f(10), None, L_AL, BOX)
    for j, s in enumerate(SHIFTS):
        put(ws, (r, 3 + j), f'=SUMPRODUCT(({CSHIFT}="{s}")*{blockcol("Q", f"$B{r}")})', f(10, True), None, C_AL, BOX)
    put(ws, f"F{r}", f'=IF($K$2="All shifts",SUM(C{r}:E{r}),INDEX(C{r}:E{r},MATCH($K$2,$C$12:$E$12,0)))', f(10, True), GREY, C_AL, BOX)
    put(ws, f"G{r}", f'=COUNTIFS({ECOL("E")},B{r},{SHF},{CR})', f(10), GREY, C_AL, BOX)
    cf(ws, f"C{r}:F{r}", f"C{r}=0", BAD_F, BAD_T)
last_dept = 12 + NP
put(ws, (last_dept + 1, 2), "Columns A / B / C = shifts.", f(8, italic=True, color=MUTED))
r0 = last_dept + 3
put(ws, f"B{r0}", "Who needs training most (selected shift)", f(13, True, NAVY))
for c, h in zip("BCDEF", ("Employee", "Gaps", "% Ready", "Job title", "Shift")):
    head(ws, r0 + 1, ws[c + "1"].column, h)
for k in range(1, 11):
    r = r0 + 1 + k
    put(ws, f"P{r}", f"=LARGE({CA['key']},{k})")
    mm = f"MATCH($P{r},{CA['key']},0)"
    put(ws, f"B{r}", f'=IF($P{r}<=0,"",INDEX({EMP_NAMES},{mm}))', f(10, True), None, L_AL, BOX)
    put(ws, f"C{r}", f'=IF($P{r}<=0,"",INDEX({ECOL("M")},{mm}))', f(10, True, "C00000"), None, C_AL, BOX)
    put(ws, f"D{r}", f'=IF($P{r}<=0,"",INDEX({ECOL("N")},{mm}))', f(10), None, C_AL, BOX, "0%")
    put(ws, f"E{r}", f'=IF($P{r}<=0,"",INDEX({ECOL("F")},{mm}))', f(9), None, L_AL, BOX)
    put(ws, f"F{r}", f'=IF($P{r}<=0,"",INDEX({ECOL("D")},{mm}))', f(9), None, C_AL, BOX)
put(ws, f"B{r0 + 12}", f'=IF(P{r0 + 2}<=0,"Nobody on this shift has a required skill gap.","")', f(9, italic=True, color=MUTED))
put(ws, "I10", "Watch list", f(13, True, NAVY))
watch = [("Skills only one person can do", f'=COUNTIF({SK("D")},1)'), ("Needed skills nobody is Competent in", f'=SUM({SK("H")})'),
         ("Cells marked NT (needs training)", f'=COUNTIF({MGRID},"NT")'), ("Cells in training (IT + TR)", f'=COUNTIF({MGRID},"IT")+COUNTIF({MGRID},"TR")')]
for i, (lab, fm) in enumerate(watch):
    r = 11 + i
    merge(ws, f"I{r}:L{r}", lab, font=f(10), al=L_AL, border=BOX)
    put(ws, f"M{r}", fm, f(11, True), None, C_AL, BOX)
put(ws, "I16", "Status legend", f(13, True, NAVY))
legend = [("NT", ST["NT"], "Needs Training"), ("IT", ST["IT"], "In Training – supervised"), ("TR", ST["TR"], "Trained – not yet signed off"),
          ("C", ST["C"], "Competent – works independently"), ("N/A", ST["N/A"], "Not applicable to this person"), ("", REQ_BLANK, "Required for the job – not assessed")]
for i, (code, cl, txt) in enumerate(legend):
    r = 17 + i
    put(ws, f"I{r}", code, f(10, True, cl[1]), fill(cl[0]), C_AL, BOX)
    merge(ws, f"J{r}:M{r}", txt, font=f(9), al=L_AL)
ch = BarChart()
ch.type = "bar"
ch.style = 10
ch.title = "Qualified vs assigned (selected shift)"
ch.add_data(Reference(ws, min_col=6, max_col=7, min_row=12, max_row=last_dept), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=2, min_row=13, max_row=last_dept))
ch.x_axis.scaling.orientation = "maxMin"
ch.x_axis.delete = False
ch.y_axis.delete = False
ch.y_axis.majorGridlines = None
ch.series[0].graphicalProperties.solidFill = "1E63C4"
ch.series[1].graphicalProperties.solidFill = "AFC1DC"
ch.gapWidth = 50
ch.width, ch.height = 15, 10.5
ws.add_chart(ch, "I24")
ws.freeze_panes = "A4"

# ================================================================== CI Analytics
ws = S["CI Analytics"]
title(ws, "Continuous Improvement Analytics", "All shifts. Counted from the Skills Matrix only – no productivity or cost savings are estimated.", navcol="H")
ws.column_dimensions["A"].width = 2
ws.column_dimensions["B"].width = 26
for c in "CDEFGHIJKLMN":
    ws.column_dimensions[c].width = 12
put(ws, "B4", "Cross-training", f(13, True, NAVY))
for c, h in zip("BC", ("Production depts qualified", "Employees")):
    head(ws, 5, ws[c + "5"].column, h)
for i, (lab, crit) in enumerate((("0", "0"), ("1", "1"), ("2", "2"), ("3 or more", ">=3"))):
    r = 6 + i
    put(ws, f"B{r}", lab, f(10), None, C_AL, BOX)
    put(ws, f"C{r}", f'=COUNTIFS({CA["prod"]},"{crit}",{EMP_IDS},"<>")' if crit.startswith(">") else f'=COUNTIFS({CA["prod"]},{crit},{EMP_IDS},"<>")', f(10, True), None, C_AL, BOX)
put(ws, "B10", "Total", f(10, True), None, C_AL, BOX)
put(ws, "C10", "=SUM(C6:C9)", f(10, True), GREY, C_AL, BOX)
put(ws, "B12", "By department", f(13, True, NAVY))
for c, h in zip("BCDEF", ("Department", "Assigned employees", "Qualified (any shift)", "Open required-skill gaps", "Skills only one person can do")):
    head(ws, 13, ws[c + "13"].column, h)
ws.row_dimensions[13].height = 32
for i, dname in enumerate(DEPTS):
    r = 14 + i
    put(ws, f"B{r}", dname, f(10), None, L_AL, BOX)
    put(ws, f"C{r}", f'=COUNTIF({ECOL("E")},B{r})', f(10), None, C_AL, BOX)
    put(ws, f"D{r}", f'=SUM({blockcol("Q", f"$B{r}")})', f(10, True), None, C_AL, BOX)
    tot = f"COUNTIF({MDEPT},$B{r})"
    put(ws, f"E{r}", f'=SUMPRODUCT({blockcol("R", f"$B{r}")}*({tot}-{blockcol("N", f"$B{r}")}-{blockcol("C", f"$B{r}")}))', f(10), None, C_AL, BOX)
    put(ws, f"F{r}", f'=COUNTIFS({SK("B")},B{r},{SK("D")},1)', f(10), None, C_AL, BOX)
    cf(ws, f"E{r}", f"E{r}>0", "FFFFFF", "C00000")
ci_last = 13 + ND
put(ws, (ci_last + 1, 2), "Total", f(10, True), None, L_AL, BOX)
for c in "CDEF":
    put(ws, f"{c}{ci_last + 1}", f"=SUM({c}14:{c}{ci_last})", f(10, True), GREY, C_AL, BOX)
r1 = ci_last + 4
put(ws, f"B{r1}", "Single-person dependencies (only one Competent person)", f(13, True, NAVY))
for c, h in zip("BCD", ("Competency", "Department", "Only person")):
    head(ws, r1 + 1, ws[c + "1"].column, h)
ws.column_dimensions["C"].width = 18
ws.column_dimensions["D"].width = 22
ws.column_dimensions["P"].hidden = True
for k in range(1, 21):
    r = r1 + 1 + k
    put(ws, f"P{r}", f"=LARGE({SK('F')},{k})")
    mm = f"MATCH($P{r},{SK('F')},0)"
    put(ws, f"B{r}", f'=IF($P{r}<=0,"",INDEX({SK("C")},{mm}))', f(9), None, L_AL, BOX)
    put(ws, f"C{r}", f'=IF($P{r}<=0,"",INDEX({SK("B")},{mm}))', f(9), None, L_AL, BOX)
    put(ws, f"D{r}", f'=IF($P{r}<=0,"",INDEX({SK("E")},{mm}))', f(9), None, L_AL, BOX)
put(ws, f"B{r1 + 22}", f'="Showing "&MIN(20,COUNTIF({SK("D")},1))&" of "&COUNTIF({SK("D")},1)&" single-person skills."', f(9, italic=True, color=MUTED))
c1 = BarChart(); c1.type = "col"; c1.style = 10; c1.title = "Employees by production departments qualified"
c1.add_data(Reference(ws, min_col=3, min_row=5, max_row=9), titles_from_data=True); c1.set_categories(Reference(ws, min_col=2, min_row=6, max_row=9))
c1.legend = None; c1.x_axis.delete = False; c1.y_axis.delete = False; c1.series[0].graphicalProperties.solidFill = "2E8B57"
c1.dataLabels = DataLabelList(); c1.dataLabels.showVal = True; c1.dataLabels.showSerName = False; c1.dataLabels.showCatName = False; c1.dataLabels.showLegendKey = False
c1.width, c1.height = 12, 6.5
ws.add_chart(c1, "E3")
c2 = BarChart(); c2.type = "bar"; c2.style = 10; c2.title = "Assigned vs qualified by department"
c2.add_data(Reference(ws, min_col=3, max_col=4, min_row=13, max_row=ci_last), titles_from_data=True); c2.set_categories(Reference(ws, min_col=2, min_row=14, max_row=ci_last))
c2.x_axis.scaling.orientation = "maxMin"; c2.x_axis.delete = False; c2.y_axis.delete = False
c2.series[0].graphicalProperties.solidFill = "AFC1DC"; c2.series[1].graphicalProperties.solidFill = "1E63C4"
c2.width, c2.height = 15, 10
ws.add_chart(c2, "H12")

# ================================================================== Person Card
ws = S["Person Card"]
title(ws, "Person Card", "Pick a name to see every skill, what the job requires and what to do next. Filter “Required” = Yes for a short list. Print this sheet for 1-to-1 reviews.", navcol="G")
for col, w in zip("ABCDEFGH", (2, 9, 22, 38, 10, 16, 24, 4)):
    ws.column_dimensions[col].width = w
ws.column_dimensions["M"].hidden = True
put(ws, "B4", "Person ▸", f(12, True, NAVY), al=Alignment(horizontal="right", vertical="center"))
merge(ws, "C4:D4", f'{m["first"]} {m["last"]}', font=f(13, True), fl=INPUT, al=L_AL, border=BOX)
dv(ws, f"={EMP_NAMES}", "C4", "Pick a name from the Employees sheet", "Person")
ws.row_dimensions[4].height = 26
put(ws, "M1", f'=IFERROR(MATCH($C$4,{EMP_NAMES},0),"")')
for col, lab, src, fmt in (("C", "Employee ID", "A", None), ("D", "Job title · Shift", None, None), ("E", "% Ready", "N", "0%"), ("F", "Gaps", "M", None), ("G", "Depts qualified", "P", None)):
    put(ws, f"{col}6", lab, f(9, True, MUTED), fill("EEF3F8"))
    fm = (f'=IF($M$1="","",INDEX({ECOL(src)},$M$1))' if src else f'=IF($M$1="","",INDEX({ECOL("F")},$M$1)&" · Shift "&INDEX({ECOL("D")},$M$1))')
    put(ws, f"{col}7", fm, f(12, True), fill("EEF3F8"), L_AL, None, fmt)
ws.row_dimensions[7].height = 26
put(ws, "B8", '=IF($M$1="","Pick a name in C4.",INDEX(' + ECOL("R") + ',$M$1))', f(10, True, NAVY))
for i, h in enumerate(["Skill ID", "Department", "Competency", "Required", "Status", "Next step"], start=2):
    head(ws, 10, i, h)
RB_A, RB_Z = L(BS["R"]), L(BS["R"] + ND - 1)
for k, (sid, dept, name, crit) in enumerate(SKILLS):
    r = 11 + k
    put(ws, (r, 2), sid, f(9, True), None, C_AL, BOX)
    put(ws, (r, 3), dept, f(9), None, L_AL, BOX)
    put(ws, (r, 4), ("◆ " if crit else "") + name, f(9), None, L_AL, BOX)
    put(ws, (r, 5), f'=IF($M$1="","",IF(IFERROR(INDEX({CALC}!${RB_A}${E0}:${RB_Z}${E1},$M$1,MATCH(C{r},{CALC}!${RB_A}$3:${RB_Z}$3,0)),0)=1,"Yes",""))', f(9, True), None, C_AL, BOX)
    put(ws, (r, 6), f'=IF($M$1="","",IF(INDEX({MGRID},$M$1,{k + 1})="",IF(E{r}="Yes","Not assessed","—"),INDEX({MGRID},$M$1,{k + 1})))', f(9, True), None, C_AL, BOX)
    put(ws, (r, 7), f'=IF($M$1="","",IF(E{r}<>"Yes","",IF(OR(F{r}="C",F{r}="N/A"),"",IF(F{r}="TR","Sign off when independent",IF(F{r}="IT","Continue training","Start training")))))', f(9, True, "C00000"), None, L_AL, BOX)
lastp = 10 + NS
for code, (fc, tc, _) in ST.items():
    cf(ws, f"F11:F{lastp}", f'F11="{code}"', fc, tc)
cf(ws, f"F11:F{lastp}", 'F11="Not assessed"', REQ_BLANK[0], REQ_BLANK[1])
cf(ws, f"B11:G{lastp}", '$E11=""', "FFFFFF", "A6A6A6", bold=False)
ws.auto_filter.ref = f"B10:G{lastp}"
ws.freeze_panes = "A11"

# ================================================================== Start Here
ws = S["Start Here"]
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 3
for c, w in zip("BCDEFG", (30, 3, 30, 3, 30, 3)):
    ws.column_dimensions[c].width = w
merge(ws, "B2:F2", "CHG Production Skills Matrix", font=f(22, True, NAVY), al=Alignment(vertical="center"))
ws.row_dimensions[2].height = 38
merge(ws, "B3:F3", "C.H. Guenther / Pioneer Frozen Foods · Duncanville. All people in this file are FICTIONAL samples (IDs start with DEMO-).",
      font=f(10, color=MUTED, italic=True), al=Alignment(wrap_text=True, vertical="top"))
ws.row_dimensions[3].height = 28
put(ws, "B5", "What do you want to do?  (click a box)", f(13, True, NAVY))
btns = [("See how the plant is doing", "Dashboard", "1E63C4"), ("Update skill levels", "Skills Matrix", "2E8B57"), ("Add or edit people", "Employees", "14284B"),
        ("Review one person", "Person Card", "7030A0"), ("CI analytics", "CI Analytics", "548235"), ("What each job requires", "Job Requirements", "7F7F7F")]
for i, (txt, tgt, col) in enumerate(btns):
    r = 7 + (i // 3) * 3
    cl = "BDF"[i % 3]
    ws.merge_cells(f"{cl}{r}:{cl}{r + 1}")
    link(ws, f"{cl}{r}", f"{txt}\n▸ {tgt}", tgt, "FFFFFF", col, 11)
for r in range(7, 13):
    ws.row_dimensions[r].height = 22
put(ws, "B14", "How it works", f(13, True, NAVY))
steps = [
    ("1. People", "On Employees enter each person's ID, name, shift (A/B/C), department, job title and supervisor. Delete or overwrite the Sample rows."),
    ("2. Skills", "On Skills Matrix pick NT / IT / TR / C / N/A in each cell. A light red empty cell means the job requires that skill and it hasn't been assessed."),
    ("3. Results", "Employees, Dashboard, Person Card and CI Analytics update by themselves. Use the shift box on the Dashboard to filter."),
]
for i, (a, b) in enumerate(steps):
    r = 15 + i
    put(ws, f"B{r}", a, f(11, True), al=Alignment(vertical="top", indent=1))
    merge(ws, f"D{r}:F{r}", b, font=f(10), al=Alignment(wrap_text=True, vertical="top"))
    ws.row_dimensions[r].height = 34
put(ws, "B19", "How things are counted", f(13, True, NAVY))
rules = [
    "Skills required = every skill in the departments marked “Yes” for the person's job title (Job Requirements), minus skills marked N/A for that person.",
    "Competent = cells marked C in those required skills.  Gaps = required − competent.  % Ready = competent ÷ required.",
    "Qualified in a department = every skill in that department is C (or N/A), with at least one C. Cross-trained = qualified in 2 or more production departments.",
    "Yellow cells = you type here. Grey cells = formulas – don't overwrite them.",
    "Proofing and Baking are not included (frozen operations only). Departments, skills and job requirements are starting assumptions – confirm them with area owners.",
]
for i, t_ in enumerate(rules):
    r = 20 + i
    merge(ws, f"B{r}:F{r}", "•  " + t_, font=f(10), al=Alignment(wrap_text=True, vertical="top"))
    ws.row_dimensions[r].height = 30
legend_r = 26
put(ws, f"B{legend_r}", "Status colours", f(13, True, NAVY))
for i, (code, (fc, tc, lab)) in enumerate(ST.items()):
    r = legend_r + 1 + i
    put(ws, f"B{r}", f"{code}  {lab}", f(11, True, tc), fill(fc), Alignment(horizontal="left", vertical="center", indent=1), BOX)
put(ws, f"B{legend_r + 6}", "(empty)  Required, not assessed", f(11, True, REQ_BLANK[1]), fill(REQ_BLANK[0]), Alignment(horizontal="left", vertical="center", indent=1), BOX)

# ================================================================== finishing
for s in wb.worksheets:
    s.sheet_properties.pageSetUpPr.fitToPage = True
    s.page_setup.orientation = "landscape"
    s.page_setup.fitToWidth = 1
    s.page_setup.fitToHeight = 0
S["Person Card"].page_setup.orientation = "portrait"
S["Skills Matrix"].print_title_rows = "3:5"
S["Skills Matrix"].print_title_cols = "A:F"
for n, c in {"Start Here": NAVY, "Dashboard": "1E63C4", "Skills Matrix": "2E8B57", "Employees": NAVY, "Person Card": "7030A0", "CI Analytics": "548235", "Job Requirements": "7F7F7F"}.items():
    S[n].sheet_properties.tabColor = c
wb.active = 0
wb.save(OUT)
print(f"Saved {OUT.name} ({NS} skills, {ND} departments, {len(people)} sample people)")
