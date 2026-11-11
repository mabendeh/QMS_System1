"""Builds the CHG Production Skills & Career Progression System as Excel workbooks.

    python training/build_chg_system_workbook.py

Outputs (in training/):
  CHG_Skills_Career_System.xlsx  - matrix, employees, training log, promotion review,
                                   coverage, dashboard, CI analytics, person card, setup
  (pay grades / skill premiums are on the red HR tabs in the same file)

Departments, competencies, job titles and career levels come from chg_config.json
(the same starting configuration as the web app). All people are fictional samples.
"""
import json
import random
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

HERE = Path(__file__).parent
CFG = json.loads((HERE / "chg_config.json").read_text())
OUT = HERE / "CHG_Skills_Career_System.xlsx"
TODAY = date.today()
FONT = "Arial"

# ------------------------------------------------------------------ sizes / anchors
AREAS = [a for a in CFG["AREAS"] if a["active"]]          # Proofing & Baking are inactive (frozen scope)
SKILLS = [(f'{a["code"]}-{i + 1:02d}', a["name"], c[0], len(c) > 2 and c[2] == "critical", c[1], a["code"])
          for a in AREAS for i, c in enumerate(a["competencies"])]
REQUAL = {"FS": 12, "SA": 12, "MD": 12}
REQUAL_SKILL = {"Allergen changeover": 12}
SK_N, SK0 = 140, 5                     # competency slots on Departments & Skills (rows 5..144)
SK1 = SK0 + SK_N - 1
D_N, D0 = 20, 5                        # department slots (rows 5..24)
D1 = D0 + D_N - 1
E_N, E0 = 150, 6                       # employee rows 6..155 on Employees / Matrix / Calc
E1 = E0 + E_N - 1
LOG_N, LOG0 = 2000, 6
LOG1 = LOG0 + LOG_N - 1
PR_N, PR0 = 100, 6
JR_N, JR0 = 15, 6                      # job titles rows on Job Requirements
MC = 7                                 # first skill column on matrix (G)
MCL = L(MC)
MLAST = L(MC + SK_N - 1)
SHIFTS = ["A", "B", "C"]
LEVELS = CFG["CAREER_LEVELS"]
JOBS = CFG["JOB_ROLES"]
CODE2NAME = {a["code"]: a["name"] for a in CFG["AREAS"]}

BLOCKS = ["C", "V", "T", "R", "Q", "P"]   # C count, verified-current count, IT/TR count, required, qualified, competent-required
BSTART = {b: 7 + i * D_N for i, b in enumerate(BLOCKS)}
ANCH = {name: 7 + 6 * D_N + i for i, name in enumerate(
    ["gap", "EZ_req", "FA_comp", "FB_gaps", "FC_tr", "FD_prod", "FE_fs", "FF_sa", "FG_prim", "FH_next", "FI_e", "FJ_f", "FK_g", "FL_h", "FM_i", "FN_ready", "FO_missing", "FP_key", "FQ_full", "FR_rank"])}
CL = {n: L(c) for n, c in ANCH.items()}

# ------------------------------------------------------------------ style helpers
NAVY, INK, MUTED = "14284B", "1A2433", "5D6B7E"
ST = {  # code: (fill, font, label)
    "NT": ("C62828", "FFFFFF", "Needs Training"),
    "IT": ("F2C230", "1A2433", "In Training"),
    "TR": ("2E8B57", "FFFFFF", "Trained"),
    "C": ("1E63C4", "FFFFFF", "Competent"),
    "N/A": ("E3E7EC", "5D6B7E", "Not Applicable / Not Assessed"),
}
EXP = ("8E1B1B", "FFFFFF")
UNVER = ("9DC3E6", "14284B")
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


def title(ws, text, sub, nav=True, navcol="H"):
    put(ws, "A1", text, f(18, True, NAVY))
    put(ws, "A2", sub, f(10, color=MUTED, italic=True))
    ws.row_dimensions[1].height = 30
    ws.sheet_view.showGridLines = False
    if nav:
        link(ws, f"{navcol}1", "◂ Start Here", "Start Here")
        link(ws, f"{L(ws[navcol + '1'].column + 1)}1", "▦ Dashboard", "Dashboard")


def dv(ws, formula, rng, prompt=None, title_=None, allow_blank=True):
    d = DataValidation(type="list", formula1=formula, allow_blank=allow_blank, prompt=prompt, promptTitle=title_,
                       showErrorMessage=True, error="Pick a value from the list.", errorTitle="Not in list")
    ws.add_data_validation(d)
    d.add(rng)


def cf_fill(ws, rng, formula, fillc, fontc, bold=True, stop=False):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[formula], fill=fill(fillc), font=Font(name=FONT, bold=bold, color=fontc), stopIfTrue=stop))


wb = Workbook()
S = {}
for name in ["Start Here", "Dashboard", "Skills Matrix", "Employees", "Training Log", "Person Card", "Promotion Review",
             "Shift Coverage", "CI Analytics", "Departments & Skills", "Job Requirements", "Career Levels", "Lists", "Calc"]:
    S[name] = wb.active if name == "Start Here" else wb.create_sheet(name)
    if name == "Start Here":
        S[name].title = name
SKN = "'Departments & Skills'"
EMP, MX, LOG, PRM, COV, LVL, JR, CALC, LST = "Employees", "'Skills Matrix'", "'Training Log'", "'Promotion Review'", "'Shift Coverage'", "'Career Levels'", "'Job Requirements'", "Calc", "Lists"

# ================================================================== Lists
ws = S["Lists"]
title(ws, "Lists", "Drop-down values. Yellow cells may be edited.", nav=False)
cols = {
    "A": ("Shift", SHIFTS), "B": ("Skill status", list(ST.keys())), "C": ("Yes/No", ["Yes", "No"]),
    "D": ("Practical result", ["Pass", "Not passed"]), "E": ("TWI step", ["1 Prepare", "2 Present", "3 Try Out", "4 Follow Up"]),
    "F": ("Decision", ["Approved", "Rejected"]), "G": ("Dashboard filter", ["All shifts"] + SHIFTS), "H": ("Employee status", ["Active", "Inactive"]),
    "I": ("Supervisors", ["Supervisor A (sample)", "Supervisor B (sample)", "Supervisor C (sample)"]),
    "J": ("Trainers / evaluators", ["Demo Trainer", "Demo Production Manager", "Demo Quality Lead"]),
    "K": ("Department type", ["Production", "Plant-wide"]),
}
for col, (h, vals) in cols.items():
    head(ws, 4, ws[col + "4"].column, h)
    ws.column_dimensions[col].width = 24
    for i, v in enumerate(vals):
        put(ws, f"{col}{5 + i}", v, f(), INPUT if col in "IJ" else None)
for i in range(len(cols["I"][1]), 30):
    put(ws, f"I{5 + i}", None, fl=INPUT)
    put(ws, f"J{5 + i}", None, fl=INPUT)
LISTS = {k: f"Lists!${k}$5:${k}${4 + len(v[1])}" for k, v in cols.items()}
LISTS["I"], LISTS["J"] = "Lists!$I$5:$I$34", "Lists!$J$5:$J$34"
ws.sheet_state = "hidden"

# ================================================================== Departments & Skills
ws = S["Departments & Skills"]
title(ws, "Departments & Skills", "Edit the yellow cells to rename, add or reorder. Starting checklists are assumptions – verify against the Duncanville plant. ◆ critical = needs practical evaluation AND evidence.", navcol="F")
for i, (h, w) in enumerate([("Skill ID", 9), ("Department", 22), ("Competency", 40), ("Critical", 8), ("Requalify every (months)", 11),
                            ("Standard / what Competent looks like", 54), ("People Competent", 10), ("Only qualified person", 22), ("key", 4)], start=1):
    head(ws, 4, i, h)
    ws.column_dimensions[L(i)].width = w
ws.row_dimensions[4].height = 42
dcol = {"J": ("Department", 26), "K": ("Code", 7), "L": ("Type", 12), "M": ("Competencies", 12), "N": ("Notes", 40), "O": ("prod#", 5)}
for col, (h, w) in dcol.items():
    head(ws, 4, ws[col + "4"].column, h)
    ws.column_dimensions[col].width = w
for k in range(SK_N):
    r = SK0 + k
    s = SKILLS[k] if k < len(SKILLS) else None
    months = (REQUAL_SKILL.get(s[2]) or REQUAL.get(s[5])) if s else None
    for c, v in zip("ABCDEF", (s[0], s[1], s[2], "Yes" if s[3] else "No", months, s[4]) if s else ("", "", "", "", "", "")):
        put(ws, f"{c}{r}", v if v != "" else None, f(10, c in "AC"), INPUT, L_AL if c in "BCF" else C_AL, BOX)
    mcol = L(MC + k)
    put(ws, f"G{r}", f'=IF(A{r}="","",COUNTIF({MX}!{mcol}${E0}:{mcol}${E1},"C"))', f(10, True), GREY, C_AL, BOX)
    put(ws, f"H{r}", f'=IF(AND(A{r}<>"",G{r}=1),INDEX({MX}!$B${E0}:$B${E1},MATCH("C",{MX}!{mcol}${E0}:{mcol}${E1},0)),"")', f(9), GREY, L_AL, BOX)
    put(ws, f"I{r}", f'=IF(AND(A{r}<>"",G{r}=1),1+(1000-ROW())/100000,0)', f(8, color="FFFFFF"))
    ws.row_dimensions[r].height = 30
for d in range(D_N):
    r = D0 + d
    a = AREAS[d] if d < len(AREAS) else None
    put(ws, f"J{r}", a["name"] if a else None, f(10, True), INPUT, L_AL, BOX)
    put(ws, f"K{r}", a["code"] if a else None, f(), INPUT, C_AL, BOX)
    put(ws, f"L{r}", ("Production" if a["kind"] == "PRODUCTION" else "Plant-wide") if a else None, f(), INPUT, C_AL, BOX)
    put(ws, f"M{r}", f'=IF(J{r}="","",COUNTIF($B${SK0}:$B${SK1},J{r}))', f(10, True), GREY, C_AL, BOX)
    put(ws, f"N{r}", a["description"] if a else None, f(9), INPUT, L_AL, BOX)
    put(ws, f"O{r}", f'=IF(L{r}="Production",COUNTIF(L${D0}:L{r},"Production"),"")', f(8, color="FFFFFF"))
put(ws, f"J{D1 + 2}", "Proofing and Baking are left out on purpose (frozen operations only). Add them in a spare row if needed.", f(9, italic=True, color=MUTED))
dv(ws, f"={LISTS['C']}", f"D{SK0}:D{SK1}")
dv(ws, f"=$J${D0}:$J${D1}", f"B{SK0}:B{SK1}")
dv(ws, f"={LISTS['K']}", f"L{D0}:L{D1}")
ws.column_dimensions["I"].hidden = True
ws.column_dimensions["O"].hidden = True
ws.freeze_panes = "D5"
DEPT_NAMES = f"{SKN}!$J${D0}:$J${D1}"

# ================================================================== Job Requirements
ws = S["Job Requirements"]
title(ws, "Job Requirements", "Mark “Yes” for every department a job title must be fully qualified in. Unassessed required skills show RED on the matrix.", navcol="D")
head(ws, 5, 2, "Job title")
ws.column_dimensions["A"].width = 2
ws.column_dimensions["B"].width = 26
for d in range(D_N):
    c = 3 + d
    head(ws, 5, c, f'={SKN}!J{D0 + d}', rotate=90)
    ws.column_dimensions[L(c)].width = 6
ws.row_dimensions[5].height = 150
for j in range(JR_N):
    r = JR0 + j
    job = JOBS[j] if j < len(JOBS) else None
    put(ws, (r, 2), job["title"] if job else None, f(10, True), INPUT, L_AL, BOX)
    for d in range(D_N):
        code = AREAS[d]["code"] if d < len(AREAS) else None
        put(ws, (r, 3 + d), "Yes" if job and code in job["areas"] else None, f(10, True), INPUT, C_AL, BOX)
JR_GRID = f"{JR}!$C${JR0}:${L(2 + D_N)}${JR0 + JR_N - 1}"
JR_JOBS = f"{JR}!$B${JR0}:$B${JR0 + JR_N - 1}"
JR_DEPTS = f"{JR}!$C$5:${L(2 + D_N)}$5"
dv(ws, f"={LISTS['C']}", f"C{JR0}:{L(2 + D_N)}{JR0 + JR_N - 1}")
ws.conditional_formatting.add(f"C{JR0}:{L(2 + D_N)}{JR0 + JR_N - 1}", CellIsRule(operator="equal", formula=['"Yes"'], fill=fill("1E63C4"), font=Font(name=FONT, bold=True, color="FFFFFF")))
ws.freeze_panes = "C6"

# ================================================================== Career Levels
ws = S["Career Levels"]
title(ws, "Career Levels & Promotion Rules", "DEMONSTRATION rules – not official CHG policy. HR edits the yellow cells. Time employed is never a requirement.", navcol="H")
hdr = [("Rank", 6), ("Level title", 24), ("Description", 40), ("Must be qualified in own department", 14), ("Min. production departments qualified", 14),
       ("Food safety required", 11), ("Safety required", 11), ("Must currently be level", 11), ("Confirmed during review (manual)", 44)]
for i, (h, w) in enumerate(hdr, start=2):
    head(ws, 5, i, h)
    ws.column_dimensions[L(i)].width = w
ws.column_dimensions["A"].width = 2
ws.row_dimensions[5].height = 54
RULES = {0: ("", 0, "", "", None, ""), 1: ("Yes", 0, "Yes", "Yes", 0, "Supervisor approval"), 2: ("No", 2, "Yes", "Yes", 1, "Supervisor approval"),
         3: ("No", 3, "Yes", "No", 2, "Supervisor confirms independent problem-solving"), 4: ("No", 0, "No", "No", 3, "Approved train-the-trainer program; supervisor and HR approval")}
for i, lv in enumerate(LEVELS):
    r = 6 + i
    e, mn, fs, sa, cur, man = RULES[lv["rank"]]
    put(ws, (r, 2), lv["rank"], f(10, True), None, C_AL, BOX)
    for c, v in zip(range(3, 11), (lv["name"], lv["description"], e, mn, fs, sa, cur, man)):
        put(ws, (r, c), v if v != "" else None, f(10, c == 3), INPUT, L_AL if c in (3, 4, 10) else C_AL, BOX)
    ws.row_dimensions[r].height = 32
put(ws, "B13", "Food safety department", f(10, True))
put(ws, "E13", "Food Safety & GMP", f(), INPUT, None, BOX)
put(ws, "B14", "Safety department", f(10, True))
put(ws, "E14", "Safety", f(), INPUT, None, BOX)
put(ws, "B16", "Promotion workflow: Supervisor recommends → Production Manager reviews → HR verifies qualification & pay policy → Plant Leadership gives final approval. "
               "No pay changes happen in this workbook.", f(9, italic=True, color=MUTED))
dv(ws, f"={LISTS['C']}", "E6:E10")
dv(ws, f"={LISTS['C']}", "G6:H10")
LV_NAMES = f"{LVL}!$C$6:$C$10"

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
skill_by = {}
for k, s in enumerate(SKILLS):
    skill_by.setdefault(s[5], []).append((k, s))
people = []
for i, (role, s) in enumerate(plan):
    primary = CFG["PRIMARY_AREA_FOR_ROLE"][role]
    trainee = role == "Production Associate" or i % 13 == 7
    p = {"id": f"DEMO-{1001 + i}", "first": first[i % len(first)], "last": last[i % len(last)], "shift": SHIFTS[s], "dept": CODE2NAME[primary],
         "job": role, "hire": TODAY - timedelta(days=int((rnd.uniform(1, 3) if trainee else rnd.uniform(6, 60)) * 30.4)),
         "sup": cols["I"][1][s], "level": LEVELS[0 if trainee else (2 if role == "Line Lead" else 1)]["name"], "status": {}, "log": {}}
    areas = list(job_areas[role])
    xa = cross[primary] if (not trainee and rnd.random() < 0.35 and primary in cross) else None
    if xa:
        areas.append(xa)
    for code in areas:
        for k, sk in skill_by.get(code, []):
            x = rnd.random()
            if trainee:
                st = ("C" if x < .6 else "IT") if code in ("FS", "SA") else ("TR" if x < .15 else "IT" if x < .7 else "")
            else:
                st = "C" if x < .86 else "TR" if x < .92 else "IT" if x < .97 else ""
            if st:
                p["status"][k] = st
    people.append(p)


def set_area(p, code, st, months_ago):
    for k, sk in skill_by[code]:
        p["status"][k] = st
        if st == "C":
            p["log"][k] = {"ver": TODAY - timedelta(days=int(months_ago * 30.4))}


def find(role, shift, nth=0):
    return [p for p in people if p["job"] == role and p["shift"] == shift][nth]


multi = find("Pre-Weigh Operator", "A")
for c in ("PW", "MX", "FS", "SA"):
    set_area(multi, c, "C", 3)
multi.update(first="Alex", last="Multiskill-Sample", level=LEVELS[1]["name"])
mixer = find("Mixer Operator", "B")
for c in ("MX", "FS", "SA"):
    set_area(mixer, c, "C", 5)
trainee = find("Mixer Operator", "A", 1)
set_area(trainee, "MX", "IT", 0)
trainee.update(first="Jordan", last="Trainee-Sample", level=LEVELS[0]["name"])
blender = find("Blending Operator", "B")
for c in ("BL", "FS", "SA"):
    set_area(blender, c, "C", 6)
btm = next(k for k, s in skill_by["BL"] if s[2] == "Blend time monitoring")
blender["status"][btm] = "TR"
blender["log"][btm] = {"pending": True}
blender.update(first="Taylor", last="Pending-Sample")
packer = find("Packaging Operator", "C")
for c in ("PK", "MD", "PL", "FS", "SA"):
    set_area(packer, c, "C", 4)
packer.update(first="Morgan", last="Promotion-Sample", level=LEVELS[1]["name"])
freezer = find("Freezer Operator", "A")
for c in ("SF", "FS", "SA"):
    set_area(freezer, c, "C", 5)
haccp = next(k for k, s in skill_by["FS"] if s[2].startswith("HACCP"))
gmp = next(k for k, s in skill_by["FS"] if s[2].startswith("GMPs"))
freezer["log"][haccp] = {"ver": TODAY - timedelta(days=395)}
freezer["log"][gmp] = {"ver": TODAY - timedelta(days=365 - 14)}
freezer.update(first="Riley", last="Expired-Sample")

log_rows = []
for p in people:
    for k, st in sorted(p["status"].items()):
        sk = SKILLS[k]
        lg = p["log"].get(k, {})
        if st == "C":
            months_req = REQUAL_SKILL.get(sk[2]) or REQUAL.get(sk[5])
            ver = lg.get("ver") or TODAY - timedelta(days=int((rnd.uniform(.5, 11) if months_req else rnd.uniform(1, 30)) * 30.4))
            log_rows.append([p["id"], sk[0], "Demo Trainer", ver - timedelta(days=30), ver - timedelta(days=7), "4 Follow Up", "Pass", "Yes", ver - timedelta(days=1),
                             "Signed checklist (sample)" if sk[3] else None, "Demo Production Manager", ver, "Sample record"])
        elif st == "TR":
            pend = lg.get("pending")
            log_rows.append([p["id"], sk[0], "Demo Trainer", TODAY - timedelta(days=30), TODAY + timedelta(days=7), "4 Follow Up",
                             "Pass" if pend else None, "Yes" if pend else None, TODAY - timedelta(days=3) if pend else None, None, None, None,
                             "Ready for sign-off (sample)" if pend else "Sample record"])
        elif st == "IT" and (p is trainee or rnd.random() < .5):
            due = TODAY + timedelta(days=rnd.randint(-6, 40))
            log_rows.append([p["id"], sk[0], "Demo Trainer", TODAY - timedelta(days=20), due, rnd.choice(["2 Present", "3 Try Out"]), None, None, None, None, None, None, "Sample record"])

# ================================================================== Employees
ws = S["Employees"]
title(ws, "Employees", "One row per person. Type in the yellow cells; grey cells calculate. Rows marked Sample are fictional – overwrite or delete them.", navcol="L")
ecols = [("Employee ID", 11), ("First name", 13), ("Last name", 17), ("Shift", 7), ("Department", 18), ("Job title", 20), ("Hire date", 11),
         ("Supervisor", 20), ("Career level", 20), ("Status", 9), ("Sample?", 8), ("Full name", 22), ("Skills required", 9), ("Competent (verified)", 10),
         ("Gaps", 7), ("% Ready", 8), ("In training", 9), ("Production depts qualified", 11), ("Cross-trained", 9), ("Next level", 20),
         ("Promotion check", 24), ("What is still missing", 44)]
for i, (h, w) in enumerate(ecols, start=1):
    head(ws, 5, i, h, NAVY if i <= 11 else "404040")
    ws.column_dimensions[L(i)].width = w
ws.row_dimensions[5].height = 44
for k in range(E_N):
    r = E0 + k
    p = people[k] if k < len(people) else None
    vals = [p["id"], p["first"], p["last"], p["shift"], p["dept"], p["job"], p["hire"], p["sup"], p["level"], "Active", "Yes"] if p else [None] * 11
    for c, v in enumerate(vals, start=1):
        put(ws, (r, c), v, f(10), INPUT, L_AL if c in (2, 3, 5, 6, 8, 9) else C_AL, BOX, "dd-mmm-yy" if c == 7 else None)
    cr = f"{CALC}!"
    formulas = [
        f'=IF(B{r}="","",B{r}&" "&C{r})', f'=IF(A{r}="","",{cr}{CL["EZ_req"]}{r})', f'=IF(A{r}="","",{cr}{CL["FA_comp"]}{r})', f'=IF(A{r}="","",M{r}-N{r})',
        f'=IF(OR(A{r}="",N(M{r})=0),"",N{r}/M{r})', f'=IF(A{r}="","",{cr}{CL["FC_tr"]}{r})', f'=IF(A{r}="","",{cr}{CL["FD_prod"]}{r})', f'=IF(A{r}="","",IF(R{r}>=2,"Yes",""))',
        f'=IF(A{r}="","",IFERROR(INDEX({LV_NAMES},{cr}{CL["FH_next"]}{r}+1),"Top level"))',
        f'=IF(A{r}="","",IF({cr}{CL["FN_ready"]}{r},"Meets requirements – ready for review","Not yet"))', f'=IF(A{r}="","",{cr}{CL["FO_missing"]}{r})',
    ]
    for c, fm in enumerate(formulas, start=12):
        put(ws, (r, c), fm, f(10, c in (12, 15, 16)), GREY, L_AL if c in (12, 20, 21, 22) else C_AL, BOX, "0%" if c == 16 else None)
dv(ws, f"={LISTS['A']}", f"D{E0}:D{E1}")
dv(ws, f"={DEPT_NAMES}", f"E{E0}:E{E1}")
dv(ws, f"={JR_JOBS}", f"F{E0}:F{E1}")
dv(ws, f"={LISTS['I']}", f"H{E0}:H{E1}")
dv(ws, f"={LV_NAMES}", f"I{E0}:I{E1}")
dv(ws, f"={LISTS['H']}", f"J{E0}:J{E1}")
dv(ws, f"={LISTS['C']}", f"K{E0}:K{E1}")
cf_fill(ws, f"P{E0}:P{E1}", f"AND(ISNUMBER(P{E0}),P{E0}>=1)", OK_F, OK_T, stop=True)
cf_fill(ws, f"P{E0}:P{E1}", f"AND(ISNUMBER(P{E0}),P{E0}>=0.8)", AMB_F, AMB_T, stop=True)
cf_fill(ws, f"P{E0}:P{E1}", f"ISNUMBER(P{E0})", BAD_F, BAD_T)
cf_fill(ws, f"U{E0}:U{E1}", f'LEFT(U{E0},5)="Meets"', "DDEBF7", NAVY)
cf_fill(ws, f"O{E0}:O{E1}", f"AND(ISNUMBER(O{E0}),O{E0}>0)", "FFFFFF", "C00000")
cf_fill(ws, f"K{E0}:K{E1}", f'K{E0}="Yes"', AMB_F, AMB_T)
ws.freeze_panes = "D6"
ws.auto_filter.ref = f"A5:V{E1}"
EMP_IDS = f"{EMP}!$A${E0}:$A${E1}"
EMP_NAMES = f"{EMP}!$L${E0}:$L${E1}"

# ================================================================== Training Log
ws = S["Training Log"]
title(ws, "Training & Qualification Log", "ONE row per employee per competency. Assign → train (TWI) → practical evaluation → verify. When someone requalifies, update the verification date on the same row.", navcol="J")
lcols = [("Employee ID", 11), ("Employee (auto)", 22), ("Skill ID", 9), ("Competency (auto)", 30), ("Department (auto)", 18), ("Critical (auto)", 8),
         ("Trainer / evaluator", 20), ("Start date", 11), ("Due date", 11), ("TWI step", 12), ("Practical result", 11), ("Operated independently", 11),
         ("Practical date", 11), ("Evidence reference", 22), ("Verified by", 22), ("Verification date", 12), ("Requalify months (auto)", 10),
         ("Expires (auto)", 11), ("Status (auto)", 24), ("Notes", 30), ("key", 4)]
for i, (h, w) in enumerate(lcols, start=1):
    head(ws, 5, i, h, NAVY if i not in (2, 4, 5, 6, 17, 18, 19) else "404040")
    ws.column_dimensions[L(i)].width = w
ws.row_dimensions[5].height = 44
SKA, SKB, SKC, SKD, SKE = (f"{SKN}!${c}${SK0}:${c}${SK1}" for c in "ABCDE")
inputs = {1: 0, 3: 1, 7: 2, 8: 3, 9: 4, 10: 5, 11: 6, 12: 7, 13: 8, 14: 9, 15: 10, 16: 11, 20: 12}
for k in range(LOG_N):
    r = LOG0 + k
    row = log_rows[k] if k < len(log_rows) else None
    for c, idx in inputs.items():
        put(ws, (r, c), row[idx] if row else None, f(9), INPUT, L_AL if c in (7, 14, 15, 20) else C_AL, BOX, "dd-mmm-yy" if c in (8, 9, 13, 16) else None)
    put(ws, (r, 2), f'=IF(A{r}="","",IFERROR(INDEX({EMP_NAMES},MATCH(A{r},{EMP_IDS},0)),"Unknown ID"))', f(9), GREY, L_AL, BOX)
    put(ws, (r, 4), f'=IF(C{r}="","",IFERROR(INDEX({SKC},MATCH(C{r},{SKA},0)),"Unknown skill"))', f(9), GREY, L_AL, BOX)
    put(ws, (r, 5), f'=IF(C{r}="","",IFERROR(INDEX({SKB},MATCH(C{r},{SKA},0)),""))', f(9), GREY, L_AL, BOX)
    put(ws, (r, 6), f'=IF(C{r}="","",IFERROR(INDEX({SKD},MATCH(C{r},{SKA},0)),""))', f(9), GREY, C_AL, BOX)
    put(ws, (r, 17), f'=IF(C{r}="","",IFERROR(IF(INDEX({SKE},MATCH(C{r},{SKA},0))="","",INDEX({SKE},MATCH(C{r},{SKA},0))),""))', f(9), GREY, C_AL, BOX)
    put(ws, (r, 18), f'=IF(OR(P{r}="",NOT(ISNUMBER(Q{r}))),"",EDATE(P{r},Q{r}))', f(9), GREY, C_AL, BOX, "dd-mmm-yy")
    put(ws, (r, 19),
        f'=IF(A{r}="","",IF(P{r}<>"",IF(OR(K{r}<>"Pass",L{r}<>"Yes"),"Invalid – no passed practical",IF(AND(F{r}="Yes",N{r}=""),"Invalid – evidence missing",'
        f'IF(AND(O{r}<>"",O{r}=G{r}),"Invalid – verifier ran the practical",IF(AND(R{r}<>"",R{r}<TODAY()),"Expired",IF(AND(R{r}<>"",R{r}<TODAY()+30),"Expiring soon","Qualified"))))),'
        f'IF(K{r}="Pass",IF(L{r}="Yes","Ready to verify","Pass must confirm independent operation"),IF(AND(I{r}<>"",I{r}<TODAY()),"Overdue",IF(H{r}<>"","In progress","Assigned")))))',
        f(9, True), GREY, L_AL, BOX)
    put(ws, (r, 21), f'=IF(A{r}="","",A{r}&"|"&C{r})', f(7, color="FFFFFF"))
LG = {c: f"{LOG}!${c}${LOG0}:${c}${LOG1}" for c in "ACESU"}
dv(ws, f"={EMP_IDS}", f"A{LOG0}:A{LOG1}")
dv(ws, f"={SKA}", f"C{LOG0}:C{LOG1}")
dv(ws, f"={LISTS['J']}", f"G{LOG0}:G{LOG1}")
dv(ws, f"={LISTS['E']}", f"J{LOG0}:J{LOG1}")
dv(ws, f"={LISTS['D']}", f"K{LOG0}:K{LOG1}")
dv(ws, f"={LISTS['C']}", f"L{LOG0}:L{LOG1}")
dv(ws, f"={LISTS['J']}", f"O{LOG0}:O{LOG1}")
sr = f"S{LOG0}:S{LOG1}"
for txt, fc, tc in (("Qualified", "1E63C4", "FFFFFF"), ("Expiring soon", AMB_F, AMB_T), ("Expired", EXP[0], "FFFFFF"), ("Ready to verify", "DDEBF7", NAVY),
                    ("Overdue", BAD_F, BAD_T), ("In progress", "F2C230", INK)):
    ws.conditional_formatting.add(sr, CellIsRule(operator="equal", formula=[f'"{txt}"'], fill=fill(fc), font=Font(name=FONT, bold=True, color=tc)))
cf_fill(ws, sr, f'LEFT(S{LOG0},7)="Invalid"', BAD_F, BAD_T)
cf_fill(ws, f"A{LOG0}:A{LOG1}", f'COUNTIF($U${LOG0}:$U${LOG1},U{LOG0})>1', "FF7C80", "FFFFFF")
ws.column_dimensions["U"].hidden = True
ws.freeze_panes = "C6"
ws.auto_filter.ref = f"A5:T{LOG1}"
put(ws, "L2", "Red Employee ID = duplicate row for the same skill (keep one row per person per skill).", f(9, italic=True, color="C00000"))

# ================================================================== Skills Matrix
ws = S["Skills Matrix"]
title(ws, "Production Skills Matrix", "Pick a status in each cell: NT Needs Training · IT In Training · TR Trained · C Competent · N/A. Red empty cell = required for the job, not assessed yet.", navcol="D")
ws.row_dimensions[2].height = 10
for i, (h, w) in enumerate([("Employee ID", 11), ("Employee", 22), ("Shift", 6), ("Department", 16), ("Job title", 18), ("% Ready", 8)], start=1):
    ws.merge_cells(start_row=3, start_column=i, end_row=5, end_column=i)
    head(ws, 3, i, h, NAVY if i < 6 else "404040")
    ws.column_dimensions[L(i)].width = w
for k in range(SK_N):
    c = MC + k
    col = L(c)
    sr_ = SK0 + k
    put(ws, (2, c), f"={SKN}!B{sr_}&\"\"", f(6, color="FFFFFF"))
    head(ws, 3, c, None, "2E4A7A", size=9)
    head(ws, 4, c, f'={SKN}!A{sr_}&""', "2E4A7A", size=8)
    h5 = head(ws, 5, c, f'=IF({SKN}!D{sr_}="Yes","◆ ","")&{SKN}!C{sr_}', "2E4A7A", rotate=90, size=9)
    ws.column_dimensions[col].width = 4.6
put(ws, "A2", None)
ws.row_dimensions[3].height = 22
ws.row_dimensions[5].height = 210
seed_by_row = {}
for k in range(E_N):
    r = E0 + k
    for c, src in zip("ABCDE", "ALDEF"):
        put(ws, f"{c}{r}", f'=IF({EMP}!$A{r}="","",{EMP}!{src}{r})', f(10, c == "B"), None, L_AL if c in "BDE" else C_AL, BOX)
    put(ws, f"F{r}", f'=IF({EMP}!$A{r}="","",{EMP}!P{r})', f(10, True), GREY, C_AL, BOX, "0%")
    p = people[k] if k < len(people) else None
    for j in range(SK_N):
        cell = ws.cell(r, MC + j, p["status"].get(j) if p else None)
        cell.font = f(9, True)
        cell.alignment = C_AL
        cell.border = BOX
    ws.row_dimensions[r].height = 18
grid = f"{MCL}{E0}:{MLAST}{E1}"
tl = f"{MCL}{E0}"
key = f'$A{E0}&"|"&{MCL}$4'
cf_fill(ws, grid, f'AND({tl}="C",COUNTIFS({LG["U"]},{key},{LG["S"]},"Expired")>0)', EXP[0], EXP[1], stop=True)
cf_fill(ws, grid, f'AND({tl}="C",COUNTIFS({LG["U"]},{key},{LG["S"]},"Qualified")+COUNTIFS({LG["U"]},{key},{LG["S"]},"Expiring soon")=0)', UNVER[0], UNVER[1], stop=True)
for code, (fc, tc, _) in ST.items():
    cf_fill(ws, grid, f'{tl}="{code}"', fc, tc, stop=True)
req = f'IFERROR(INDEX({JR_GRID},MATCH($E{E0},{JR_JOBS},0),MATCH({MCL}$2,{JR_DEPTS},0)),"")="Yes"'
cf_fill(ws, grid, f'AND({tl}="",$A{E0}<>"",{MCL}$4<>"",{req})', "F4B6B6", "C62828")
cf_fill(ws, f"F{E0}:F{E1}", f"AND(ISNUMBER(F{E0}),F{E0}>=1)", OK_F, OK_T, stop=True)
cf_fill(ws, f"F{E0}:F{E1}", f"AND(ISNUMBER(F{E0}),F{E0}>=0.8)", AMB_F, AMB_T, stop=True)
cf_fill(ws, f"F{E0}:F{E1}", f"ISNUMBER(F{E0})", BAD_F, BAD_T)
dv(ws, f"={LISTS['B']}", grid, "NT Needs Training | IT In Training | TR Trained | C Competent (needs a verified row in the Training Log) | N/A", "Skill status")
# department labels (merged per starting group) + one collapsible column group per department
start = MC
for a in AREAS:
    n = len(a["competencies"])
    ws.merge_cells(start_row=3, start_column=start, end_row=3, end_column=start + n - 1)
    lab = ws.cell(3, start, f"={L(start)}2")
    lab.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.column_dimensions.group(L(start), L(start + n - 1), outline_level=1, hidden=False)
    start += n
ws.merge_cells(start_row=3, start_column=start, end_row=3, end_column=MC + SK_N - 1)
ws.cell(3, start, "Spare columns – add skills on Departments & Skills").alignment = Alignment(horizontal="left", vertical="center")
ws.column_dimensions.group(L(start), MLAST, outline_level=1, hidden=False)
ws.sheet_properties.outlinePr.summaryRight = False
ws.freeze_panes = ws.cell(E0, MC)
ws.auto_filter.ref = f"A5:F{E1}"
c = ws["B3"]
c.comment = Comment("Edit people on the Employees sheet. Use the + / – buttons above the column letters to collapse departments.", "CHG")
MGRID = f"{MX}!${MCL}${E0}:${MLAST}${E1}"
MDEPT = f"{MX}!${MCL}$2:${MLAST}$2"

# ================================================================== Calc (hidden)
ws = S["Calc"]
put(ws, "A1", "Helper calculations – do not edit.", f(10, True))
put(ws, "A2", "Shift filter")
put(ws, "B2", '=IF(Dashboard!$K$2="All shifts","*",Dashboard!$K$2)')
for b in BLOCKS:
    for d in range(D_N):
        c = BSTART[b] + d
        put(ws, (3, c), f"={SKN}!$J${D0 + d}&\"\"")
        put(ws, (4, c), f"=N({SKN}!$M${D0 + d})")
        put(ws, (5, c), f'=IF({SKN}!$L${D0 + d}="Production",1,0)')
    put(ws, (2, BSTART[b]), b, f(9, True))


def blk(b, d=None, r=None, abs_row=False):
    if d is None:
        a, z = L(BSTART[b]), L(BSTART[b] + D_N - 1)
        return f"${a}${r}:${z}${r}" if abs_row else f"{a}{r}:{z}{r}"
    return f"{L(BSTART[b] + d)}{r}"


for k in range(E_N):
    r = E0 + k
    put(ws, (r, 1), f'=IF({EMP}!A{r}="","",{EMP}!A{r})')
    put(ws, (r, 2), f'=IF(A{r}="","",{EMP}!D{r})')
    put(ws, (r, 3), f'=IF(A{r}="","",{EMP}!F{r})')
    put(ws, (r, 4), f'=IF(A{r}="","",{EMP}!E{r})')
    put(ws, (r, 5), f'=IF(A{r}="",-1,IFERROR(MATCH({EMP}!I{r},{LV_NAMES},0)-1,-1))')
    mrow = f"{MX}!${MCL}{r}:${MLAST}{r}"
    for d in range(D_N):
        H = L(BSTART["C"] + d)
        put(ws, (r, BSTART["C"] + d), f'=IF(OR($A{r}="",{H}$3=""),0,COUNTIFS({MDEPT},{H}$3,{mrow},"C"))')
        HV = L(BSTART["V"] + d)
        put(ws, (r, BSTART["V"] + d), f'=IF(OR($A{r}="",{HV}$3=""),0,COUNTIFS({LG["A"]},$A{r},{LG["E"]},{HV}$3,{LG["S"]},"Qualified")+COUNTIFS({LG["A"]},$A{r},{LG["E"]},{HV}$3,{LG["S"]},"Expiring soon"))')
        HT = L(BSTART["T"] + d)
        put(ws, (r, BSTART["T"] + d), f'=IF(OR($A{r}="",{HT}$3=""),0,COUNTIFS({MDEPT},{HT}$3,{mrow},"IT")+COUNTIFS({MDEPT},{HT}$3,{mrow},"TR"))')
        HR = L(BSTART["R"] + d)
        put(ws, (r, BSTART["R"] + d), f'=IF(OR($A{r}="",{HR}$3=""),0,IF(IFERROR(INDEX({JR_GRID},MATCH($C{r},{JR_JOBS},0),MATCH({HR}$3,{JR_DEPTS},0)),"")="Yes",1,0))')
        HQ = L(BSTART["Q"] + d)
        put(ws, (r, BSTART["Q"] + d), f'=IF(AND({HQ}$4>0,{blk("C", d, r)}>={HQ}$4,{blk("V", d, r)}>={HQ}$4),1,0)')
        put(ws, (r, BSTART["P"] + d), f'={blk("R", d, r)}*MIN({blk("C", d, r)},{blk("V", d, r)},{L(BSTART["P"] + d)}$4)')
    A = {n: L(c) for n, c in ANCH.items()}
    qh = f"${L(BSTART['Q'])}$3:${L(BSTART['Q'] + D_N - 1)}$3"
    ws.cell(r, ANCH["EZ_req"]).value = f'=SUMPRODUCT({blk("R", r=r)},${L(BSTART["R"])}$4:${L(BSTART["R"] + D_N - 1)}$4)'
    put(ws, (r, ANCH["FA_comp"]), f'=SUM({blk("P", r=r)})')
    put(ws, (r, ANCH["FB_gaps"]), f'={A["EZ_req"]}{r}-{A["FA_comp"]}{r}')
    put(ws, (r, ANCH["FC_tr"]), f'=SUMPRODUCT({blk("R", r=r)},{blk("T", r=r)})')
    put(ws, (r, ANCH["FD_prod"]), f'=SUMPRODUCT({blk("Q", r=r)},${L(BSTART["Q"])}$5:${L(BSTART["Q"] + D_N - 1)}$5)')
    put(ws, (r, ANCH["FE_fs"]), f'=IFERROR(INDEX({blk("Q", r=r)},MATCH({LVL}!$E$13,{qh},0)),0)')
    put(ws, (r, ANCH["FF_sa"]), f'=IFERROR(INDEX({blk("Q", r=r)},MATCH({LVL}!$E$14,{qh},0)),0)')
    put(ws, (r, ANCH["FG_prim"]), f'=IFERROR(INDEX({blk("Q", r=r)},MATCH($D{r},{qh},0)),0)')
    put(ws, (r, ANCH["FH_next"]), f'=IF(OR($A{r}="",$E{r}<0,$E{r}>=COUNTA({LV_NAMES})-1),"",$E{r}+1)')
    nr = f"{A['FH_next']}{r}"
    idx = lambda col: f"INDEX({LVL}!${col}$6:${col}$10,{nr}+1)"
    put(ws, (r, ANCH["FI_e"]), f'=IF({nr}="",FALSE,OR({idx("E")}<>"Yes",{A["FG_prim"]}{r}=1))')
    put(ws, (r, ANCH["FJ_f"]), f'=IF({nr}="",FALSE,{A["FD_prod"]}{r}>=N({idx("F")}))')
    put(ws, (r, ANCH["FK_g"]), f'=IF({nr}="",FALSE,OR({idx("G")}<>"Yes",{A["FE_fs"]}{r}=1))')
    put(ws, (r, ANCH["FL_h"]), f'=IF({nr}="",FALSE,OR({idx("H")}<>"Yes",{A["FF_sa"]}{r}=1))')
    put(ws, (r, ANCH["FM_i"]), f'=IF({nr}="",FALSE,$E{r}>=N({idx("I")}))')
    put(ws, (r, ANCH["FN_ready"]), f'=IF({nr}="",FALSE,AND({A["FI_e"]}{r},{A["FJ_f"]}{r},{A["FK_g"]}{r},{A["FL_h"]}{r},{A["FM_i"]}{r}))')
    put(ws, (r, ANCH["FO_missing"]),
        f'=IF($A{r}="","",IF({nr}="","Top level reached",IF({A["FN_ready"]}{r},"Only the review confirmations remain: "&{idx("J")},'
        f'TRIM(IF(NOT({A["FI_e"]}{r}),"Qualify in own department. ","")&IF(NOT({A["FJ_f"]}{r}),"Qualify in "&{idx("F")}&" production departments (has "&{A["FD_prod"]}{r}&"). ","")'
        f'&IF(NOT({A["FK_g"]}{r}),"Complete food safety. ","")&IF(NOT({A["FL_h"]}{r}),"Complete safety. ","")&IF(NOT({A["FM_i"]}{r}),"Must first reach level "&{idx("I")}&". ","")))))')
    put(ws, (r, ANCH["FP_key"]), f'=IF(AND($A{r}<>"",OR($B$2="*",$B{r}=Dashboard!$K$2),{A["FB_gaps"]}{r}>0),{A["FB_gaps"]}{r}+(1000-ROW())/100000,0)')
    put(ws, (r, ANCH["FQ_full"]), f'=AND($A{r}<>"",{A["EZ_req"]}{r}>0,{A["FB_gaps"]}{r}=0)')
for n, c in ANCH.items():
    put(ws, (5, c), n)
ws.sheet_state = "hidden"
CA = {n: f"{CALC}!${L(c)}${E0}:${L(c)}${E1}" for n, c in ANCH.items()}
QB = f"{CALC}!${L(BSTART['Q'])}${E0}:${L(BSTART['Q'] + D_N - 1)}${E1}"
QH = f"{CALC}!${L(BSTART['Q'])}$3:${L(BSTART['Q'] + D_N - 1)}$3"
CSHIFT = f"{CALC}!$B${E0}:$B${E1}"
CR = f"{CALC}!$B$2"

# ================================================================== Shift Coverage
ws = S["Shift Coverage"]
title(ws, "Shift Coverage", "Qualified = every competency in the department is Competent AND has a current verification in the Training Log. Targets (yellow) are demonstration values.", navcol="H")
ccols = [("Department", 26), ("Shift A qualified", 10), ("Shift B qualified", 10), ("Shift C qualified", 10), ("Shift A target", 9), ("Shift B target", 9),
         ("Shift C target", 9), ("Total qualified", 10), ("Total target", 9), ("Status", 12), ("Shifts below target", 10), ("Selected shift qualified", 11), ("Selected shift target", 11), ("key", 4)]
for i, (h, w) in enumerate(ccols, start=2):
    head(ws, 5, i, h)
    ws.column_dimensions[L(i)].width = w
ws.column_dimensions["A"].width = 2
ws.row_dimensions[5].height = 44
NPROD = 18
for k in range(NPROD):
    r = 6 + k
    put(ws, (r, 2), f'=IFERROR(INDEX({SKN}!$J${D0}:$J${D1},MATCH({k + 1},{SKN}!$O${D0}:$O${D1},0)),"")', f(10, True), None, L_AL, BOX)
    for i, s in enumerate(SHIFTS):
        put(ws, (r, 3 + i), f'=IF($B{r}="","",SUMPRODUCT(({CSHIFT}="{s}")*INDEX({QB},0,MATCH($B{r},{QH},0))))', f(10, True), None, C_AL, BOX)
        prod = [a for a in AREAS if a["kind"] == "PRODUCTION"]
        tgt = (3 if prod[k]["code"] == "PK" else 2) if k < len(prod) and prod[k]["code"] not in ("SH", "LM", "SN", "MH") else None
        put(ws, (r, 6 + i), tgt, f(10), INPUT, C_AL, BOX)
    put(ws, (r, 9), f'=IF(B{r}="","",SUM(C{r}:E{r}))', f(10, True), GREY, C_AL, BOX)
    put(ws, (r, 10), f'=IF(OR(B{r}="",COUNT(F{r}:H{r})=0),"",SUM(F{r}:H{r}))', f(10), GREY, C_AL, BOX)
    put(ws, (r, 12), None, f(10), GREY, C_AL, BOX)
    ws.cell(r, 12).value = f'=IF(B{r}="","",' + "+".join(f"IF(AND(ISNUMBER({t}{r}),{q}{r}<{t}{r}),1,0)" for q, t in zip("CDE", "FGH")) + ")"
    put(ws, (r, 11), f'=IF(B{r}="","",IF(COUNT(F{r}:H{r})=0,"No target",IF(L{r}>0,"Gap","Covered")))', f(10, True), GREY, C_AL, BOX)
    put(ws, (r, 13), f'=IF(B{r}="","",IF(Dashboard!$K$2="All shifts",I{r},INDEX(C{r}:E{r},MATCH(Dashboard!$K$2,{{"A","B","C"}},0))))', f(10, True), GREY, C_AL, BOX)
    put(ws, (r, 14), f'=IF(B{r}="","",IF(Dashboard!$K$2="All shifts",J{r},INDEX(F{r}:H{r},MATCH(Dashboard!$K$2,{{"A","B","C"}},0))))', f(10), GREY, C_AL, BOX)
    put(ws, (r, 15), f'=IF(N(L{r})>0,L{r}+(100-ROW())/1000,0)', f(7, color="FFFFFF"))
    for q, t in zip("CDE", "FGH"):
        cf_fill(ws, f"{q}{r}", f"AND(ISNUMBER({t}{r}),{q}{r}<{t}{r})", BAD_F, BAD_T)
        cf_fill(ws, f"{q}{r}", f"AND(ISNUMBER({t}{r}),{q}{r}>={t}{r})", "DDEBF7", "1E63C4")
for txt, fc, tc in (("Covered", "DDEBF7", "1E63C4"), ("Gap", BAD_F, BAD_T), ("No target", "EDEFF2", MUTED)):
    ws.conditional_formatting.add(f"K6:K{5 + NPROD}", CellIsRule(operator="equal", formula=[f'"{txt}"'], fill=fill(fc), font=Font(name=FONT, bold=True, color=tc)))
ws.column_dimensions["O"].hidden = True
put(ws, (7 + NPROD, 2), "Scheduled qualified headcount: UNKNOWN · Available qualified headcount: UNKNOWN – no scheduling or attendance system is connected, so these are not estimated.", f(10, True, "7F6000"))
put(ws, (8 + NPROD, 2), "Plant-wide trainings (Food Safety & GMP, Safety) are not production stations and are not listed here.", f(9, italic=True, color=MUTED))
ws.freeze_panes = "C6"
COV_ROWS = (6, 5 + NPROD)

# ================================================================== Promotion Review
ws = S["Promotion Review"]
title(ws, "Promotion Review", "Supervisor recommends → Production Manager → HR (qualification & pay policy) → Final approval. Recommendations only: nothing here changes pay. Keep signed copies.", navcol="F")
pcols = [("Employee ID", 11), ("Employee (auto)", 20), ("Current level (auto)", 18), ("Recommended level", 20), ("System check (auto)", 34), ("Why recommended", 34),
         ("Submitted by", 18), ("Submitted date", 11), ("Manager decision", 11), ("Manager name", 18), ("Manager date", 11), ("Manager comments", 26),
         ("HR decision", 11), ("HR name", 18), ("HR date", 11), ("HR comments", 26), ("Final decision", 11), ("Final approver", 18), ("Final date", 11),
         ("Effective date", 11), ("Final comments", 26), ("Status (auto)", 22), ("Checks (auto)", 34)]
for i, (h, w) in enumerate(pcols, start=1):
    head(ws, 5, i, h, "404040" if h.endswith("(auto)") else NAVY)
    ws.column_dimensions[L(i)].width = w
ws.row_dimensions[5].height = 40
sample_pr = [packer["id"], None, None, LEVELS[2]["name"], None, "Qualified in Packaging and Palletizing; covers both stations on Shift C. (Sample request)",
             "Supervisor C (sample)", TODAY - timedelta(days=2)]
for k in range(PR_N):
    r = PR0 + k
    for c in [1, 4] + list(range(6, 22)):
        v = sample_pr[c - 1] if k == 0 and c <= len(sample_pr) else None
        put(ws, (r, c), v, f(9), INPUT, L_AL if c in (6, 7, 10, 12, 14, 16, 18, 21) else C_AL, BOX, "dd-mmm-yy" if c in (8, 11, 15, 19, 20) else None)
    m = f"MATCH(A{r},{EMP_IDS},0)"
    put(ws, (r, 2), f'=IF(A{r}="","",IFERROR(INDEX({EMP_NAMES},{m}),"Unknown ID"))', f(9), GREY, L_AL, BOX)
    put(ws, (r, 3), f'=IF(A{r}="","",IFERROR(INDEX({EMP}!$I${E0}:$I${E1},{m}),""))', f(9), GREY, L_AL, BOX)
    put(ws, (r, 5),
        f'=IF(OR(A{r}="",D{r}=""),"",IFERROR(IF(D{r}=INDEX({EMP}!$T${E0}:$T${E1},{m}),IF(INDEX({CA["FN_ready"]},{m}),"Meets documented requirements","NOT MET: "&INDEX({CA["FO_missing"]},{m})),"Not the next level – check manually"),"Unknown ID"))',
        f(9), GREY, L_AL, BOX)
    put(ws, (r, 22), f'=IF(A{r}="","",IF(OR(I{r}="Rejected",M{r}="Rejected",Q{r}="Rejected"),"Rejected",IF(Q{r}="Approved","Approved",IF(M{r}="Approved","Awaiting final approval",IF(I{r}="Approved","Awaiting HR","Awaiting manager")))))', f(9, True), GREY, C_AL, BOX)
    checks = (f'=IF(A{r}="","",TRIM(IF(AND(M{r}<>"",I{r}<>"Approved"),"HR decided before manager approval. ","")&IF(AND(Q{r}<>"",M{r}<>"Approved"),"Final decided before HR approval. ","")'
              f'&IF(OR(AND(I{r}="Rejected",L{r}=""),AND(M{r}="Rejected",P{r}=""),AND(Q{r}="Rejected",U{r}="")),"Rejection needs a reason. ","")'
              f'&IF(OR(AND(J{r}<>"",OR(J{r}=G{r},J{r}=N{r},J{r}=R{r})),AND(N{r}<>"",OR(N{r}=G{r},N{r}=R{r})),AND(R{r}<>"",R{r}=G{r})),"Same person on two steps. ","")'
              f'&IF(AND(Q{r}="Approved",T{r}=""),"Effective date missing. ","")&IF(AND(Q{r}="Approved",LEFT(E{r},5)<>"Meets"),"Approved although requirements are not met. ","")))')
    put(ws, (r, 23), checks, f(9, color="C00000"), GREY, L_AL, BOX)
dv(ws, f"={EMP_IDS}", f"A{PR0}:A{PR0 + PR_N - 1}")
dv(ws, f"={LV_NAMES}", f"D{PR0}:D{PR0 + PR_N - 1}")
for col in "IMQ":
    dv(ws, f"={LISTS['F']}", f"{col}{PR0}:{col}{PR0 + PR_N - 1}")
dv(ws, f"={LISTS['I']}", f"G{PR0}:G{PR0 + PR_N - 1}")
st_r = f"V{PR0}:V{PR0 + PR_N - 1}"
for txt, fc, tc in (("Approved", "1E63C4", "FFFFFF"), ("Rejected", BAD_F, BAD_T)):
    ws.conditional_formatting.add(st_r, CellIsRule(operator="equal", formula=[f'"{txt}"'], fill=fill(fc), font=Font(name=FONT, bold=True, color=tc)))
cf_fill(ws, st_r, f'LEFT(V{PR0},8)="Awaiting"', AMB_F, AMB_T)
cf_fill(ws, f"E{PR0}:E{PR0 + PR_N - 1}", f'LEFT(E{PR0},7)="NOT MET"', BAD_F, BAD_T)
cf_fill(ws, f"E{PR0}:E{PR0 + PR_N - 1}", f'LEFT(E{PR0},5)="Meets"', "DDEBF7", NAVY)
ws.freeze_panes = "C6"

# ================================================================== Person Card
ws = S["Person Card"]
title(ws, "Person Card", "Pick a name to see every skill, what is required, training status and expiry. Use the filter on “Required” for a short list. Print this sheet for 1-to-1 reviews.", navcol="H")
for col, w in zip("ABCDEFGHIJ", (2, 9, 20, 34, 9, 13, 22, 18, 12, 22)):
    ws.column_dimensions[col].width = w
ws.column_dimensions["M"].hidden = True
put(ws, "B4", "Person ▸", f(12, True, NAVY), al=Alignment(horizontal="right", vertical="center"))
merge(ws, "C4:D4", f'{multi["first"]} {multi["last"]}', font=f(13, True), fl=INPUT, al=L_AL, border=BOX)
dv(ws, f"={EMP_NAMES}", "C4", "Pick a name from the Employees sheet", "Person")
ws.row_dimensions[4].height = 26
put(ws, "M1", f'=IFERROR(MATCH($C$4,{EMP_NAMES},0),"")')
infos = [("C", "Employee ID", "A"), ("D", "Job title · Shift", None), ("E", "% Ready", "P"), ("F", "Gaps", "O"), ("G", "Career level", "I"), ("H", "Next level", "T"), ("I", "Areas qualified", "R")]
for col, lab, src in infos:
    put(ws, f"{col}6", lab, f(9, True, MUTED), fill("EEF3F8"))
    put(ws, f"{col}7", f'=IF($M$1="","",INDEX({EMP}!${src}${E0}:${src}${E1},$M$1))' if src else
        f'=IF($M$1="","",INDEX({EMP}!$F${E0}:$F${E1},$M$1)&" · Shift "&INDEX({EMP}!$D${E0}:$D${E1},$M$1))', f(12, True), fill("EEF3F8"), L_AL, None, "0%" if src == "P" else None)
put(ws, "B8", f'=IF($M$1="","Pick a name in C4.","Promotion: "&INDEX({EMP}!$U${E0}:$U${E1},$M$1)&" – "&INDEX({EMP}!$V${E0}:$V${E1},$M$1))', f(10, True, NAVY))
ws.row_dimensions[7].height = 26
for i, h in enumerate(["Skill ID", "Department", "Competency", "Required", "Status", "Training log", "Trainer", "Expires", "Action"], start=2):
    head(ws, 10, i, h)
RB = f"{CALC}!${L(BSTART['R'])}${E0}:${L(BSTART['R'] + D_N - 1)}${E1}"
RH = f"{CALC}!${L(BSTART['R'])}$3:${L(BSTART['R'] + D_N - 1)}$3"
for k in range(SK_N):
    r = 11 + k
    put(ws, (r, 2), f'={SKN}!A{SK0 + k}&""', f(9, True), None, C_AL, BOX)
    put(ws, (r, 3), f'={SKN}!B{SK0 + k}&""', f(9), None, L_AL, BOX)
    put(ws, (r, 4), f'=IF({SKN}!C{SK0 + k}="","",IF({SKN}!D{SK0 + k}="Yes","◆ ","")&{SKN}!C{SK0 + k})', f(9), None, L_AL, BOX)
    put(ws, (r, 5), f'=IF(OR($M$1="",B{r}=""),"",IF(IFERROR(INDEX({RB},$M$1,MATCH(C{r},{RH},0)),0)=1,"Yes",""))', f(9, True), None, C_AL, BOX)
    put(ws, (r, 6), f'=IF(OR($M$1="",B{r}=""),"",IF(INDEX({MGRID},$M$1,{k + 1})="",IF(E{r}="Yes","NT (not assessed)","—"),INDEX({MGRID},$M$1,{k + 1})))', f(9, True), None, C_AL, BOX)
    lk = f'MATCH(INDEX({EMP_IDS},$M$1)&"|"&B{r},{LG["U"]},0)'
    put(ws, (r, 7), f'=IF(OR($M$1="",B{r}=""),"",IFERROR(INDEX({LG["S"]},{lk}),""))', f(9), None, L_AL, BOX)
    put(ws, (r, 8), f'=IF(OR($M$1="",B{r}=""),"",IFERROR(INDEX({LOG}!$G${LOG0}:$G${LOG1},{lk})&"",""))', f(9), None, L_AL, BOX)
    put(ws, (r, 9), f'=IF(OR($M$1="",B{r}=""),"",IFERROR(INDEX({LOG}!$R${LOG0}:$R${LOG1},{lk}),""))', f(9), None, C_AL, BOX, "dd-mmm-yy")
    put(ws, (r, 10), f'=IF(OR($M$1="",B{r}=""),"",IF(G{r}="Expired","Requalify",IF(G{r}="Expiring soon","Plan requalification",IF(G{r}="Ready to verify","Verify now",IF(AND(F{r}="C",G{r}=""),"Add verification record",IF(AND(E{r}="Yes",F{r}<>"C"),IF(G{r}="","Assign training","Continue training"),""))))))', f(9, True, "C00000"), None, L_AL, BOX)
fr = "F11:F150"
for code, (fc, tc, _) in ST.items():
    cf_fill(ws, fr, f'LEFT(F11,{len(code)})="{code}"' if code != "C" else 'F11="C"', fc, tc)
cf_fill(ws, "B11:J150", '$E11=""', "FFFFFF", "A6A6A6", bold=False)
ws.auto_filter.ref = "B10:J150"
ws.freeze_panes = "A11"

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
put(ws, "B3", "C.H. Guenther · Duncanville · updates automatically · sample data is fictional", f(10, color=MUTED, italic=True))
put(ws, "J2", "Show shift ▸", f(11, True, NAVY), al=Alignment(horizontal="right", vertical="center"))
merge(ws, "K2:L2", "All shifts", font=f(12, True), fl=INPUT, al=C_AL, border=BOX)
dv(ws, f"={LISTS['G']}", "K2", "Pick a shift to filter the dashboard", "Shift", allow_blank=False)
link(ws, "M2", "◂ Start", "Start Here")
tiles = [
    ("B", "C", "TOTAL EMPLOYEES", f'=COUNTIFS({EMP_IDS},"<>",{EMP}!$D${E0}:$D${E1},{CR})', "0", '="active roster"'),
    ("D", "E", "FULLY QUALIFIED", f'=COUNTIFS({CA["FQ_full"]},TRUE,{CSHIFT},{CR})', "0", '="all job skills verified"'),
    ("F", "G", "IN TRAINING", f'=COUNTIFS({EMP}!$Q${E0}:$Q${E1},">0",{EMP}!$D${E0}:$D${E1},{CR})', "0", '="have IT / TR skills"'),
    ("H", "I", "NEED TRAINING", f'=COUNTIFS({EMP}!$O${E0}:$O${E1},">0",{EMP}!$D${E0}:$D${E1},{CR})', "0", '="have a required gap"'),
    ("J", "K", "CROSS-TRAINED", f'=COUNTIFS({EMP}!$S${E0}:$S${E1},"Yes",{EMP}!$D${E0}:$D${E1},{CR})', "0", '="2+ production depts"'),
    ("L", "M", "PROMOTION REVIEW", f'=COUNTIF({PRM}!$V${PR0}:$V${PR0 + PR_N - 1},"Awaiting*")', "0", '="requests in review"'),
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
put(ws, "B10", "Needs attention", f(13, True, NAVY))
att = [("Qualifications ready to verify", f'=COUNTIF({LG["S"]},"Ready to verify")'), ("Training overdue", f'=COUNTIF({LG["S"]},"Overdue")'),
       ("Qualifications expiring in 30 days", f'=COUNTIF({LG["S"]},"Expiring soon")'), ("Qualifications expired", f'=COUNTIF({LG["S"]},"Expired")'),
       ("Invalid log rows (missing practical / evidence / same verifier)", f'=COUNTIF({LG["S"]},"Invalid*")'),
       ("“C” entered without a current verification record", f'=MAX(0,COUNTIF({MGRID},"C")-COUNTIF({LG["S"]},"Qualified")-COUNTIF({LG["S"]},"Expiring soon")-COUNTIF({LG["S"]},"Expired"))'),
       ("Departments below a shift target", f'=COUNTIF({COV}!$K$6:$K${5 + NPROD},"Gap")'), ("Skills only one person can do", f'=COUNTIF({SKN}!$G${SK0}:$G${SK1},1)')]
for i, (lab, fm) in enumerate(att):
    r = 11 + i
    merge(ws, f"B{r}:E{r}", lab, font=f(10), al=L_AL, border=BOX)
    put(ws, f"F{r}", fm, f(11, True), None, C_AL, BOX)
    cf_fill(ws, f"F{r}", f"F{r}>0", BAD_F, BAD_T)
put(ws, "H10", "Coverage by department (selected shift)", f(13, True, NAVY))
for c, h in zip("HIJK", ("Department", "", "Qualified", "Target")):
    head(ws, 11, ws[c + "11"].column, h)
ws.merge_cells("H11:I11")
for i in range(NPROD):
    r = 12 + i
    cr_ = COV_ROWS[0] + i
    ws.merge_cells(f"H{r}:I{r}")
    put(ws, f"H{r}", f"={COV}!B{cr_}&\"\"", f(9), None, L_AL, BOX)
    put(ws, f"J{r}", f'=IF(H{r}="","",{COV}!M{cr_})', f(9, True), None, C_AL, BOX)
    put(ws, f"K{r}", f'=IF(H{r}="","",{COV}!N{cr_})', f(9), None, C_AL, BOX)
    cf_fill(ws, f"J{r}", f"AND(ISNUMBER(K{r}),J{r}<K{r})", BAD_F, BAD_T)
    cf_fill(ws, f"J{r}", f"AND(ISNUMBER(K{r}),J{r}>=K{r})", "DDEBF7", "1E63C4")
last_cov = 11 + len([a for a in AREAS if a["kind"] == "PRODUCTION"])
put(ws, "B21", "Who needs training most (selected shift)", f(13, True, NAVY))
for c, h in zip("BCDEF", ("Employee", "Gaps", "% Ready", "Job title", "Shift")):
    head(ws, 22, ws[c + "22"].column, h)
for k in range(1, 11):
    r = 22 + k
    put(ws, f"P{r}", f"=LARGE({CA['FP_key']},{k})")
    m = f"MATCH($P{r},{CA['FP_key']},0)"
    put(ws, f"B{r}", f'=IF($P{r}<=0,"",INDEX({EMP_NAMES},{m}))', f(10, True), None, L_AL, BOX)
    put(ws, f"C{r}", f'=IF($P{r}<=0,"",INDEX({EMP}!$O${E0}:$O${E1},{m}))', f(10, True, "C00000"), None, C_AL, BOX)
    put(ws, f"D{r}", f'=IF($P{r}<=0,"",INDEX({EMP}!$P${E0}:$P${E1},{m}))', f(10), None, C_AL, BOX, "0%")
    put(ws, f"E{r}", f'=IF($P{r}<=0,"",INDEX({EMP}!$F${E0}:$F${E1},{m}))', f(9), None, L_AL, BOX)
    put(ws, f"F{r}", f'=IF($P{r}<=0,"",INDEX({EMP}!$D${E0}:$D${E1},{m}))', f(9), None, C_AL, BOX)
put(ws, "B34", "Status legend", f(13, True, NAVY))
legend = [(ST["NT"], "NT", "Needs Training"), (ST["IT"], "IT", "In Training – supervised"), (ST["TR"], "TR", "Trained – awaiting verification"),
          (ST["C"], "C", "Competent – verified, independent"), (ST["N/A"], "N/A", "Not applicable / not assessed"), ((EXP[0], EXP[1]), "C", "Expired – requalify (not counted)"),
          ((UNVER[0], UNVER[1]), "C", "Competent entered but NO verification record"), (("F4B6B6", "C62828"), "", "Required for the job – not assessed")]
for i, (cl, code, txt) in enumerate(legend):
    r = 35 + i
    put(ws, f"B{r}", code, f(10, True, cl[1]), fill(cl[0]), C_AL, BOX)
    merge(ws, f"C{r}:F{r}", txt, font=f(9), al=L_AL)
ch = BarChart()
ch.type = "bar"
ch.style = 10
ch.title = "Qualified people vs target (selected shift)"
ch.add_data(Reference(ws, min_col=10, max_col=11, min_row=11, max_row=last_cov), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=8, min_row=12, max_row=last_cov))
ch.x_axis.scaling.orientation = "maxMin"
ch.x_axis.delete = False
ch.y_axis.delete = False
ch.y_axis.majorGridlines = None
ch.series[0].graphicalProperties.solidFill = "1E63C4"
ch.series[1].graphicalProperties.solidFill = "AFC1DC"
ch.gapWidth = 50
ch.width, ch.height = 15, 11
ws.add_chart(ch, "H31")
ws.freeze_panes = "A4"

# ================================================================== CI Analytics
ws = S["CI Analytics"]
title(ws, "Continuous Improvement Analytics", "Counted from the matrix and the Training Log only. No productivity or cost savings are estimated – those need real before/after measurements.", navcol="H")
ws.column_dimensions["A"].width = 2
ws.column_dimensions["B"].width = 26
for c in "CDEFGHIJKLMN":
    ws.column_dimensions[c].width = 12
put(ws, "B4", "Cross-training", f(13, True, NAVY))
for c, h in zip("BC", ("Production depts qualified", "Employees")):
    head(ws, 5, ws[c + "5"].column, h)
for i, (lab, crit) in enumerate((("0", 0), ("1", 1), ("2", 2), ("3 or more", 3))):
    r = 6 + i
    put(ws, f"B{r}", lab, f(10), None, C_AL, BOX)
    fm = f'=COUNTIFS({CA["FD_prod"]},">=3",{EMP_IDS},"<>")' if crit == 3 else f'=COUNTIFS({CA["FD_prod"]},{crit},{EMP_IDS},"<>")'
    put(ws, f"C{r}", fm, f(10, True), None, C_AL, BOX)
put(ws, "B11", "Coverage and gaps by department (all shifts)", f(13, True, NAVY))
for c, h in zip("BCDE", ("Department", "Assigned employees", "Qualified (any dept)", "Open required-skill gaps")):
    head(ws, 12, ws[c + "12"].column, h)
ws.row_dimensions[12].height = 32
for i in range(NPROD):
    r = 13 + i
    cr_ = COV_ROWS[0] + i
    put(ws, f"B{r}", f"={COV}!B{cr_}&\"\"", f(9), None, L_AL, BOX)
    put(ws, f"C{r}", f'=IF(B{r}="","",COUNTIF({EMP}!$E${E0}:$E${E1},B{r}))', f(10), None, C_AL, BOX)
    put(ws, f"D{r}", f'=IF(B{r}="","",{COV}!I{cr_})', f(10, True), None, C_AL, BOX)
    rb = f"INDEX({RB},0,MATCH(B{r},{RH},0))"
    pb = f"INDEX({CALC}!${L(BSTART['P'])}${E0}:${L(BSTART['P'] + D_N - 1)}${E1},0,MATCH(B{r},{CALC}!${L(BSTART['P'])}$3:${L(BSTART['P'] + D_N - 1)}$3,0))"
    put(ws, f"E{r}", f'=IF(B{r}="","",SUMPRODUCT({rb})*INDEX({SKN}!$M${D0}:$M${D1},MATCH(B{r},{DEPT_NAMES},0))-SUM({pb}))', f(10), None, C_AL, BOX)
ci_last = 12 + len([a for a in AREAS if a["kind"] == "PRODUCTION"])
put(ws, "B32", "Qualifications verified per month (last 12 months)", f(13, True, NAVY))
for c, h in zip("BCD", ("Month", "Verified", "Training started")):
    head(ws, 33, ws[c + "33"].column, h)
for i in range(12):
    r = 34 + i
    put(ws, f"B{r}", f"=EOMONTH(TODAY(),{i - 12})+1", f(10), None, C_AL, BOX, "mmm yy")
    put(ws, f"C{r}", f'=COUNTIFS({LOG}!$P${LOG0}:$P${LOG1},">="&B{r},{LOG}!$P${LOG0}:$P${LOG1},"<="&EOMONTH(B{r},0))', f(10, True), None, C_AL, BOX)
    put(ws, f"D{r}", f'=COUNTIFS({LOG}!$H${LOG0}:$H${LOG1},">="&B{r},{LOG}!$H${LOG0}:$H${LOG1},"<="&EOMONTH(B{r},0))', f(10), None, C_AL, BOX)
put(ws, "H32", "Single-person dependencies (only one Competent person)", f(13, True, NAVY))
for c, h in zip("HIJ", ("Competency", "Department", "Only person")):
    head(ws, 33, ws[c + "33"].column, h)
ws.column_dimensions["H"].width = 30
ws.column_dimensions["J"].width = 22
ws.column_dimensions["P"].hidden = True
for k in range(1, 21):
    r = 33 + k
    put(ws, f"P{r}", f"=LARGE({SKN}!$I${SK0}:$I${SK1},{k})")
    m = f"MATCH($P{r},{SKN}!$I${SK0}:$I${SK1},0)"
    put(ws, f"H{r}", f'=IF($P{r}<=0,"",INDEX({SKC},{m}))', f(9), None, L_AL, BOX)
    put(ws, f"I{r}", f'=IF($P{r}<=0,"",INDEX({SKB},{m}))', f(9), None, L_AL, BOX)
    put(ws, f"J{r}", f'=IF($P{r}<=0,"",INDEX({SKN}!$H${SK0}:$H${SK1},{m}))', f(9), None, L_AL, BOX)
put(ws, "B48", f'="Competencies with nobody Competent: "&COUNTIFS({SKN}!$G${SK0}:$G${SK1},0,{SKA},"<>")', f(10, True, "C00000"))
put(ws, "B49", f'="First-time practical pass rate: "&IFERROR(TEXT(COUNTIF({LOG}!$K${LOG0}:$K${LOG1},"Pass")/COUNTA({LOG}!$K${LOG0}:$K${LOG1}),"0%"),"no evaluations yet")', f(10, True))
c1 = BarChart(); c1.type = "col"; c1.style = 10; c1.title = "Employees by production departments qualified"
c1.add_data(Reference(ws, min_col=3, min_row=5, max_row=9), titles_from_data=True); c1.set_categories(Reference(ws, min_col=2, min_row=6, max_row=9))
c1.legend = None; c1.x_axis.delete = False; c1.y_axis.delete = False; c1.series[0].graphicalProperties.solidFill = "2E8B57"
c1.dataLabels = DataLabelList(); c1.dataLabels.showVal = True; c1.dataLabels.showSerName = False; c1.dataLabels.showCatName = False; c1.dataLabels.showLegendKey = False
c1.width, c1.height = 12, 6.5
ws.add_chart(c1, "E4")
c2 = BarChart(); c2.type = "bar"; c2.style = 10; c2.title = "Assigned vs qualified by department"
c2.add_data(Reference(ws, min_col=3, max_col=4, min_row=12, max_row=ci_last), titles_from_data=True); c2.set_categories(Reference(ws, min_col=2, min_row=13, max_row=ci_last))
c2.x_axis.scaling.orientation = "maxMin"; c2.x_axis.delete = False; c2.y_axis.delete = False
c2.series[0].graphicalProperties.solidFill = "AFC1DC"; c2.series[1].graphicalProperties.solidFill = "1E63C4"
c2.width, c2.height = 15, 9
ws.add_chart(c2, "G11")
c3 = LineChart(); c3.title = "Qualifications verified per month"; c3.style = 12
c3.add_data(Reference(ws, min_col=3, max_col=4, min_row=33, max_row=45), titles_from_data=True); c3.set_categories(Reference(ws, min_col=2, min_row=34, max_row=45))
c3.x_axis.number_format = "mmm yy"; c3.x_axis.delete = False; c3.y_axis.delete = False
c3.width, c3.height = 15, 7
ws.add_chart(c3, "B51")

# ================================================================== Start Here
ws = S["Start Here"]
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 3
for c, w in zip("BCDEFG", (30, 3, 30, 3, 30, 3)):
    ws.column_dimensions[c].width = w
merge(ws, "B2:F2", "CHG Production Skills & Career Progression", font=f(22, True, NAVY), al=Alignment(vertical="center"))
ws.row_dimensions[2].height = 38
merge(ws, "B3:F3", "C.H. Guenther / Pioneer Frozen Foods · Duncanville · Excel edition. All people in this file are FICTIONAL samples (IDs start with DEMO-).",
      font=f(10, color=MUTED, italic=True), al=Alignment(wrap_text=True, vertical="top"))
ws.row_dimensions[3].height = 28
put(ws, "B5", "What do you want to do?  (click a box)", f(13, True, NAVY))
btns = [("See how the plant is doing", "Dashboard", "1E63C4"), ("Update skill levels", "Skills Matrix", "2E8B57"), ("Add or edit people", "Employees", "14284B"),
        ("Record training & verification", "Training Log", "C55A11"), ("Review one person", "Person Card", "7030A0"), ("Promotion recommendations", "Promotion Review", "404040"),
        ("Check shift coverage", "Shift Coverage", "BF8F00"), ("CI analytics", "CI Analytics", "548235"), ("Departments & skills setup", "Departments & Skills", "7F7F7F"), ("HR pay grades (HR only)", "HR Pay Grades", "C00000")]
for i, (txt, tgt, col) in enumerate(btns):
    r = 7 + (i // 3) * 3
    cl = "BDF"[i % 3]
    ws.merge_cells(f"{cl}{r}:{cl}{r + 1}")
    link(ws, f"{cl}{r}", f"{txt}\n▸ {tgt}", tgt, "FFFFFF", col, 11)
for r in range(7, 19):
    ws.row_dimensions[r].height = 22
put(ws, "B20", "How it works", f(13, True, NAVY))
steps = [
    ("1. People", "Add each person on Employees (ID, name, shift A/B/C, department, job title, supervisor, career level). Delete or overwrite the Sample rows."),
    ("2. Skills", "On Skills Matrix pick NT / IT / TR / C / N/A. A red empty cell means the job requires that skill and it hasn't been assessed."),
    ("3. Training", "On Training Log add one row per person per skill: trainer, dates, TWI step, practical result, evidence, verifier. Status calculates itself."),
    ("4. Competent", "“C” only counts when the log shows: practical PASSED, operated independently, evidence for ◆ critical skills, and a verifier who is not the trainer. Otherwise the cell turns pale blue."),
    ("5. Expiry", "Skills with a requalification interval expire automatically (dark red) and stop counting toward coverage."),
    ("6. Promotion", "Employees and Person Card show the next level and what is missing. Supervisors add a row on Promotion Review; manager, HR and final approver fill their columns."),
]
for i, (a, b) in enumerate(steps):
    r = 21 + i
    put(ws, f"B{r}", a, f(11, True), al=Alignment(vertical="top", indent=1))
    merge(ws, f"D{r}:F{r}", b, font=f(10), al=Alignment(wrap_text=True, vertical="top"))
    ws.row_dimensions[r].height = 40
put(ws, "B28", "Important", f(13, True, NAVY))
notes = [
    "Pay sheets (red tabs, HR Pay …) are in this file. Excel cannot hide them from anyone who opens it – encrypt the file (File ▸ Info ▸ Protect Workbook ▸ Encrypt with Password) and share it only with people allowed to see pay. If supervisors need the file, delete the HR tabs from their copy.",
    "Excel cannot enforce who edits what or keep a tamper-proof history. Use folder permissions, turn on version history (SharePoint/OneDrive), and keep signed copies of verifications and promotion decisions.",
    "Departments, checklists, job requirements, career levels, rules and targets are STARTING ASSUMPTIONS – confirm them with the area owners and HR.",
    "Yellow cells = you type here. Grey cells = formulas, don't overwrite. Dates as dd-mmm-yy.",
    "Proofing and Baking are left out (frozen operations only).",
]
for i, t in enumerate(notes):
    r = 29 + i
    merge(ws, f"B{r}:F{r}", "•  " + t, font=f(10), al=Alignment(wrap_text=True, vertical="top"))
    ws.row_dimensions[r].height = 30

# ================================================================== finishing
for s in wb.worksheets:
    s.sheet_properties.pageSetUpPr.fitToPage = True
    s.page_setup.orientation = "landscape"
    s.page_setup.fitToWidth = 1
    s.page_setup.fitToHeight = 0
S["Person Card"].page_setup.orientation = "portrait"
S["Skills Matrix"].print_title_rows = "3:5"
S["Skills Matrix"].print_title_cols = "A:F"
colors = {"Start Here": NAVY, "Dashboard": "1E63C4", "Skills Matrix": "2E8B57", "Employees": NAVY, "Training Log": "C55A11", "Person Card": "7030A0",
          "Promotion Review": "404040", "Shift Coverage": "BF8F00", "CI Analytics": "548235", "Departments & Skills": "7F7F7F", "Job Requirements": "7F7F7F", "Career Levels": "7F7F7F"}
for n, c in colors.items():
    S[n].sheet_properties.tabColor = c
# ================================================================== HR pay sheets (in the same workbook)
hb = wb
ws = hb.create_sheet("HR Pay Grades")
ws.sheet_view.showGridLines = False
put(ws, "A1", "CHG HR Compensation – CONFIDENTIAL", f(18, True, NAVY))
put(ws, "A2", "HR Administrators and Plant Leadership only. Anyone who can open this workbook can see these sheets – encrypt the file (File ▸ Info ▸ Protect Workbook ▸ Encrypt with Password) "
              "and store it where only approved people have access. No amounts are pre-filled – enter only APPROVED values.", f(10, italic=True, color="C00000"))
link(ws, "H1", "◂ Start Here", "Start Here")
for i, (h, w) in enumerate([("Code", 9), ("Pay grade name", 28), ("Career level", 24), ("Min hourly rate (USD)", 13), ("Max hourly rate (USD)", 13),
                            ("Effective date", 12), ("End date", 12), ("Approval authority", 22), ("Policy reference / notes", 40), ("Check", 30)], start=1):
    head(ws, 4, i, h)
    ws.column_dimensions[L(i)].width = w
ws.row_dimensions[4].height = 32
for k in range(15):
    r = 5 + k
    lv = LEVELS[k] if k < len(LEVELS) else None
    for c, v in enumerate([f"PG-{lv['rank']}" if lv else None, f"{lv['name']} (placeholder)" if lv else None, lv["name"] if lv else None, None, None, None, None, None,
                           "Not configured – HR to enter approved range" if lv else None], start=1):
        put(ws, (r, c), v, f(10), INPUT, L_AL if c in (2, 3, 8, 9) else C_AL, BOX, '"$"#,##0.00' if c in (4, 5) else "dd-mmm-yy" if c in (6, 7) else None)
    put(ws, (r, 10), f'=TRIM(IF(AND(ISNUMBER(D{r}),ISNUMBER(E{r}),D{r}>E{r}),"Min is above max. ","")&IF(AND(OR(D{r}<>"",E{r}<>""),F{r}=""),"Effective date missing. ","")'
                     f'&IF(AND(F{r}<>"",G{r}<>"",G{r}<F{r}),"End before start. ",""))', f(9, color="C00000"), GREY, L_AL, BOX)
ws2 = hb.create_sheet("HR Skill Premiums")
ws2.sheet_view.showGridLines = False
put(ws2, "A1", "Skill Premiums – CONFIDENTIAL", f(18, True, NAVY))
put(ws2, "A2", "Optional premiums for approved qualifications. Enter only approved amounts.", f(10, italic=True, color="C00000"))
for i, (h, w) in enumerate([("Premium name", 28), ("Department (optional)", 22), ("Amount (USD)", 12), ("Unit", 12), ("Effective date", 12), ("End date", 12),
                            ("Approval authority", 22), ("Policy reference / notes", 40)], start=1):
    head(ws2, 4, i, h)
    ws2.column_dimensions[L(i)].width = w
for k in range(20):
    for c in range(1, 9):
        put(ws2, (5 + k, c), "per hour" if c == 4 and k == 0 else None, f(10), INPUT, None, BOX, '"$"#,##0.00' if c == 3 else "dd-mmm-yy" if c in (5, 6) else None)
ws3 = hb.create_sheet("HR Pay Change Log")
put(ws3, "A1", "Change log – record every change (who, when, what, approved by)", f(14, True, NAVY))
for i, (h, w) in enumerate([("Date", 12), ("Changed by", 22), ("What changed", 50), ("Old value", 18), ("New value", 18), ("Approved by", 22)], start=1):
    head(ws3, 3, i, h)
    ws3.column_dimensions[L(i)].width = w
for k in range(100):
    for c in range(1, 7):
        put(ws3, (4 + k, c), None, f(10), INPUT, None, BOX, "dd-mmm-yy" if c == 1 else None)
for n, c in (("HR Pay Grades", "C00000"), ("HR Skill Premiums", "C00000"), ("HR Pay Change Log", "C00000")):
    wb[n].sheet_properties.tabColor = c
wb.move_sheet("Lists", offset=len(wb.sheetnames))
wb.move_sheet("Calc", offset=len(wb.sheetnames))
wb.active = 0
wb.save(OUT)
print(f"Saved {OUT.name} ({len(SKILLS)} skills, {len(AREAS)} departments, {len(people)} sample people, {len(log_rows)} log rows)")
