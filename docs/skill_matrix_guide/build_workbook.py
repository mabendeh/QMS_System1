"""Builds the plant-wide Skill Matrix Excel template that matches the guide deck."""
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment

OUT = sys.argv[1] if len(sys.argv) > 1 else "Skill_Matrix_Template_CH_Guenther.xlsx"

GREEN, DARK, AMBER = "1D6F42", "0E3B27", "E39B2D"
F = "Arial"
LEVEL_FILL = ["F4C7C3", "FCE4B6", "FFF2CC", "CFE8D5", "8CCB9F"]
LEVEL_TXT = ["8A1C14", "7A4A00", "6B5A00", "1D5B35", "0E3B27"]
INPUT_FILL = PatternFill("solid", fgColor="FFFDE7")
HEAD_FILL = PatternFill("solid", fgColor=GREEN)
SUM_FILL = PatternFill("solid", fgColor="F2F2F2")
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

def font(**k):
    k.setdefault("name", F); k.setdefault("size", 10)
    return Font(**k)

def head(ws, row, values, col=1, height=30):
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=col + i, value=v)
        c.font = font(bold=True, color="FFFFFF"); c.fill = HEAD_FILL; c.border = BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = height

def title(ws, text, sub):
    ws["A1"] = text; ws["A1"].font = font(size=16, bold=True, color=DARK)
    ws["A2"] = sub; ws["A2"].font = font(size=10, italic=True, color="7A8C82")

def box(ws, rng, inp=False, center=False):
    for row in ws[rng]:
        for c in row:
            c.border = BOX
            c.font = font(color="0000FF" if inp else "000000")
            if inp: c.fill = INPUT_FILL
            if center: c.alignment = Alignment(horizontal="center", vertical="center")

DEPTS = ["Mixing", "Packaging", "Sanitation", "Warehouse", "Quality", "Maintenance"]
SKILLS = [
    ("Production", "Mixer Setup", "SOP-MX-01", 3),
    ("Production", "Line Packaging", "SOP-PK-04", 5),
    ("Food Safety", "Allergen Control", "Allergen Plan", 6),
    ("Food Safety", "Sanitation (SSOP)", "SSOP-02", 4),
    ("Food Safety", "CCP Monitoring", "HACCP Plan", 4),
    ("Quality", "Metal Detector", "SOP-QA-07", 3),
    ("Equipment", "Forklift", "Safety-12", 2),
]
PEOPLE = [
    (1040, "Maria G.", "Mixing", "Operator", "A", [4, 3, 4, 3, 3, 3, 0]),
    (1047, "James T.", "Packaging", "Lead", "B", [3, 4, 3, 2, 2, 3, 3]),
    (1054, "Ana R.", "Packaging", "Operator", "C", [2, 3, 3, 4, 1, 2, 0]),
    (1061, "David K.", "Mixing", "Supervisor", "A", [3, 2, 2, 3, 4, 4, 2]),
    (1068, "Luis M.", "Warehouse", "Driver", "B", [1, 3, 2, 2, 0, 1, 4]),
    (1075, "Sarah P.", "Sanitation", "Temp", "C", [0, 2, 3, 1, 3, 2, 0]),
    (1082, "Kevin B.", "Sanitation", "Sanitor", "A", [2, 1, 1, 3, 0, 1, 3]),
    (1089, "Nora H.", "Quality", "QA Tech", "B", [3, 2, 4, 2, 2, 3, 1]),
]
FIRST, LAST = 5, 300          # matrix data rows (as in the deck)
SKILL_COLS = "CDEFGHI"        # 7 skill columns, J = versatility

wb = Workbook()

# ---------------- Instructions ----------------
ws = wb.active; ws.title = "Instructions"
title(ws, "Plant-wide Skill Matrix – C.H. Guenther", "Template that follows the 'Building a Skill Matrix in Excel' guide")
lines = [
    ("How to use", None),
    ("1", "Employees: list every person in the plant (Emp ID, Name, Department, Role, Shift)."),
    ("2", "Skills: list the skills, their source document and the plant-wide Target (minimum qualified people)."),
    ("3", "Matrix: pick each name in column A (drop-down) – Department fills automatically. Score each skill 0–4 from the drop-down."),
    ("4", "By Department and Dashboard update automatically. GAP = fewer qualified people (level 3–4) than the target."),
    ("5", "Training Log: record the last training date per person and skill; 'Retrain' shows after 365 days."),
    ("", ""),
    ("Cell colours", None),
    ("Yellow cells, blue text", "Input cells – type or pick from the drop-down here."),
    ("White / grey cells", "Formulas – do not overwrite."),
    ("", ""),
    ("Example data", None),
    ("", "The 8 people and their scores are EXAMPLES only. Replace them with your own roster and ratings (delete example rows on Employees, Matrix and Training Log)."),
    ("", "Skill names and targets on the Skills sheet are examples too – replace with your own SOP / HACCP tasks."),
    ("", ""),
    ("Tips", None),
    ("", "Filter Matrix column B to see one department. Add rows anywhere between rows 5 and 300."),
    ("", "To protect formulas: Review › Protect Sheet (the yellow input cells are already unlocked)."),
    ("", "To add a department: type it into the Departments list on the Legend sheet and add it to By Department column A."),
]
r = 4
for a, b in lines:
    if b is None:
        ws.cell(row=r, column=1, value=a).font = font(size=12, bold=True, color=GREEN)
    else:
        ws.cell(row=r, column=1, value=a).font = font(bold=True)
        ws.cell(row=r, column=2, value=b).font = font()
        ws.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical="top")
    r += 1
ws["A12"].fill = INPUT_FILL; ws["A12"].font = font(bold=True, color="0000FF")
ws.column_dimensions["A"].width = 24; ws.column_dimensions["B"].width = 110

# ---------------- Legend ----------------
lg = wb.create_sheet("Legend")
title(lg, "Rating scale", "Only levels 3 and 4 count as 'qualified'")
head(lg, 4, ["Level", "Name", "Meaning"])
levels = [("Not trained", "Has not received training on this task"),
          ("In training", "Learning; must be supervised at all times"),
          ("Trained", "Can do the task with occasional support"),
          ("Qualified", "Works independently to the SOP standard"),
          ("Trainer", "Can teach others and sign off training")]
for i, (n, m) in enumerate(levels):
    rr = 5 + i
    for j, v in enumerate([i, n, m]):
        c = lg.cell(row=rr, column=1 + j, value=v)
        c.fill = PatternFill("solid", fgColor=LEVEL_FILL[i]); c.border = BOX
        c.font = font(bold=j < 2, color=LEVEL_TXT[i])
        c.alignment = Alignment(horizontal="center" if j == 0 else "left", vertical="center")
head(lg, 4, ["Departments"], col=5)
for i, d in enumerate(DEPTS):
    lg.cell(row=5 + i, column=5, value=d)
box(lg, "E5:E20", inp=True)
lg.column_dimensions["A"].width = 8; lg.column_dimensions["B"].width = 16
lg.column_dimensions["C"].width = 46; lg.column_dimensions["E"].width = 18

# ---------------- Employees ----------------
em = wb.create_sheet("Employees")
title(em, "Employees", "One row per person in the plant – source: HR roster")
head(em, 4, ["Emp ID", "Name", "Department", "Role", "Shift"])
for i, p in enumerate(PEOPLE):
    for j, v in enumerate(p[:5]):
        em.cell(row=5 + i, column=1 + j, value=v)
box(em, f"A5:E{LAST}", inp=True)
for c in em[f"E5:E{LAST}"]:
    c[0].alignment = Alignment(horizontal="center")
for col, w in zip("ABCDE", [10, 22, 16, 16, 8]):
    em.column_dimensions[col].width = w
em.freeze_panes = "A5"
dv = DataValidation(type="list", formula1="=Legend!$E$5:$E$20", allow_blank=True)
em.add_data_validation(dv); dv.add(f"C5:C{LAST}")
dvs = DataValidation(type="list", formula1='"A,B,C"', allow_blank=True)
em.add_data_validation(dvs); dvs.add(f"E5:E{LAST}")
em.auto_filter.ref = f"A4:E{LAST}"

# ---------------- Skills ----------------
sk = wb.create_sheet("Skills")
title(sk, "Skills", "Tasks from SOPs, job descriptions and the HACCP / food safety plan")
head(sk, 4, ["Category", "Skill", "Source", "Target"])
for i, s in enumerate(SKILLS):
    for j, v in enumerate(s):
        sk.cell(row=5 + i, column=1 + j, value=v)
box(sk, "A5:D11", inp=True)
for c in sk["D5:D11"]:
    c[0].alignment = Alignment(horizontal="center")
sk["D4"].comment = Comment("Target = plant-wide minimum number of qualified people (level 3-4) needed to cover all shifts. Example values sized for the 8 example people – set your own plant-wide targets.", "QA")
for col, w in zip("ABCD", [16, 24, 18, 10]):
    sk.column_dimensions[col].width = w

# ---------------- Matrix ----------------
mx = wb.create_sheet("Matrix")
title(mx, "Skill Matrix – whole plant", "Pick a name in column A, then score each skill 0–4. Scores start in cell C5.")
hdr = ["Name", "Department"] + [None] * 7 + ["Qualified skills"]
head(mx, 4, hdr, height=60)
for k, col in enumerate(SKILL_COLS):
    c = mx[f"{col}4"]; c.value = f"=Skills!B{5 + k}"
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True, text_rotation=90)
for i in range(FIRST, LAST + 1):
    mx[f"B{i}"] = f'=IF(A{i}="","",IFERROR(INDEX(Employees!$C$5:$C${LAST},MATCH(A{i},Employees!$B$5:$B${LAST},0)),"Not on roster"))'
    mx[f"J{i}"] = f'=IF(A{i}="","",COUNTIF(C{i}:I{i},">=3"))'
for i, p in enumerate(PEOPLE):
    rr = FIRST + i
    mx[f"A{rr}"] = p[1]
    for k, v in enumerate(p[5]):
        mx[f"{SKILL_COLS[k]}{rr}"] = v
box(mx, f"A{FIRST}:A{LAST}", inp=True)
box(mx, f"C{FIRST}:I{LAST}", inp=True, center=True)
box(mx, f"B{FIRST}:B{LAST}")
box(mx, f"J{FIRST}:J{LAST}", center=True)
# summary rows 301-305 (as in the deck)
mx["B301"] = "Summary"; mx["B301"].font = font(bold=True, color=GREEN)
labels = {302: "Target", 303: "Qualified (3–4)", 304: "Gap", 305: "Status"}
for rr, lab in labels.items():
    c = mx[f"B{rr}"]; c.value = lab; c.font = font(bold=True); c.fill = SUM_FILL; c.border = BOX
for k, col in enumerate(SKILL_COLS):
    mx[f"{col}301"] = f"={col}4"
    mx[f"{col}301"].font = font(bold=True, color="FFFFFF"); mx[f"{col}301"].fill = HEAD_FILL
    mx[f"{col}301"].alignment = Alignment(horizontal="center", wrap_text=True); mx[f"{col}301"].border = BOX
    mx[f"{col}302"] = f"=Skills!D{5 + k}"
    mx[f"{col}303"] = f'=COUNTIF({col}${FIRST}:{col}${LAST},">=3")'
    mx[f"{col}304"] = f"={col}303-{col}302"
    mx[f"{col}305"] = f'=IF({col}304<0,"GAP","OK")'
    for rr in range(302, 306):
        c = mx[f"{col}{rr}"]; c.border = BOX; c.font = font(bold=rr in (303, 305))
        c.alignment = Alignment(horizontal="center"); c.fill = SUM_FILL
mx.row_dimensions[301].height = 30
mx.freeze_panes = "C5"
mx.column_dimensions["A"].width = 22; mx.column_dimensions["B"].width = 16
for col in SKILL_COLS:
    mx.column_dimensions[col].width = 7
mx.column_dimensions["J"].width = 11
mx.auto_filter.ref = f"A4:J{LAST}"
# drop-downs
dvn = DataValidation(type="list", formula1=f"=Employees!$B$5:$B${LAST}", allow_blank=True,
                     showErrorMessage=True, errorTitle="Unknown name", error="Pick a name from the Employees sheet.")
mx.add_data_validation(dvn); dvn.add(f"A{FIRST}:A{LAST}")
dvl = DataValidation(type="list", formula1='"0,1,2,3,4"', allow_blank=True, showErrorMessage=True,
                     errorStyle="stop", errorTitle="Invalid level", error="Enter a level from 0 to 4")
mx.add_data_validation(dvl); dvl.add(f"C{FIRST}:I{LAST}")
# conditional formatting per level
rng = f"C{FIRST}:I{LAST}"
for lvl in range(5):
    mx.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND(C{FIRST}<>"",C{FIRST}={lvl})'],
        fill=PatternFill("solid", fgColor=LEVEL_FILL[lvl], bgColor=LEVEL_FILL[lvl]),
        font=Font(name=F, bold=True, color=LEVEL_TXT[lvl])))
st = "C305:I305"
mx.conditional_formatting.add(st, CellIsRule(operator="equal", formula=['"GAP"'], fill=PatternFill("solid", fgColor=LEVEL_FILL[0], bgColor=LEVEL_FILL[0]), font=Font(name=F, bold=True, color=LEVEL_TXT[0])))
mx.conditional_formatting.add(st, CellIsRule(operator="equal", formula=['"OK"'], fill=PatternFill("solid", fgColor=LEVEL_FILL[3], bgColor=LEVEL_FILL[3]), font=Font(name=F, bold=True, color=LEVEL_TXT[3])))
mx.conditional_formatting.add("C304:I304", CellIsRule(operator="lessThan", formula=["0"], font=Font(name=F, color="C0504D")))
wb.defined_names["Dept"] = DefinedName("Dept", attr_text=f"Matrix!$B${FIRST}:$B${LAST}")

# ---------------- By Department ----------------
bd = wb.create_sheet("By Department")
title(bd, "Qualified people by department", "Counts level 3–4 per department. '–' = nobody qualified; red = only 1 qualified person.")
head(bd, 4, ["Department"] + [None] * 7, height=60)
for k in range(7):
    c = bd.cell(row=4, column=2 + k, value=f"=Matrix!{SKILL_COLS[k]}4")
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True, text_rotation=90)
nd = 12  # rows 5..16 for departments (6 used + spare)
for i in range(nd):
    rr = 5 + i
    bd[f"A{rr}"] = f'=IF(Legend!E{5 + i}="","",Legend!E{5 + i})'
    for k in range(7):
        L = "BCDEFGH"[k]
        bd[f"{L}{rr}"] = f'=IF($A{rr}="","",COUNTIFS(Dept,$A{rr},Matrix!{SKILL_COLS[k]}${FIRST}:{SKILL_COLS[k]}${LAST},">=3"))'
box(bd, f"A5:H{4 + nd}")
for row in bd[f"B5:H{4 + nd}"]:
    for c in row:
        c.alignment = Alignment(horizontal="center"); c.number_format = '0;-0;"–"'
tr = 5 + nd
bd[f"A{tr}"] = "Plant total"
for k in range(7):
    L = "BCDEFGH"[k]
    bd[f"{L}{tr}"] = f"=SUM({L}5:{L}{4 + nd})"
box(bd, f"A{tr}:H{tr}", center=True)
for c in bd[f"A{tr}:H{tr}"][0]:
    c.font = font(bold=True); c.fill = SUM_FILL
bd[f"A{tr}"].alignment = Alignment(horizontal="left")
bd[f"A{tr + 1}"] = "Plant total should equal Matrix row 303 (people whose name is not on the roster are not counted by department)."
bd[f"A{tr + 1}"].font = font(italic=True, size=9, color="7A8C82")
r2 = f"B5:H{4 + nd}"
bd.conditional_formatting.add(r2, CellIsRule(operator="equal", formula=["1"], fill=PatternFill("solid", fgColor=LEVEL_FILL[0], bgColor=LEVEL_FILL[0]), font=Font(name=F, bold=True, color=LEVEL_TXT[0])))
bd.conditional_formatting.add(r2, CellIsRule(operator="equal", formula=["2"], fill=PatternFill("solid", fgColor=LEVEL_FILL[2], bgColor=LEVEL_FILL[2]), font=Font(name=F, bold=True, color=LEVEL_TXT[2])))
bd.conditional_formatting.add(r2, CellIsRule(operator="greaterThanOrEqual", formula=["3"], fill=PatternFill("solid", fgColor=LEVEL_FILL[3], bgColor=LEVEL_FILL[3]), font=Font(name=F, bold=True, color=LEVEL_TXT[3])))
bd.column_dimensions["A"].width = 18
for L in "BCDEFGH":
    bd.column_dimensions[L].width = 8
bd.freeze_panes = "B5"

# ---------------- Training Log ----------------
tl = wb.create_sheet("Training Log")
title(tl, "Training log & expiry", "One row per person per skill. 'Retrain' appears 365 days after the last training date.")
head(tl, 4, ["Name", "Skill", "Last trained", "Trainer", "Status"])
ex = [("Maria G.", "Allergen Control", "2026-03-12", "Nora H."),
      ("Sarah P.", "CCP Monitoring", "2025-08-20", "David K."),
      ("Luis M.", "Forklift", "2026-06-02", "External")]
import datetime as dt
for i, (n, s, d, t) in enumerate(ex):
    rr = 5 + i
    tl[f"A{rr}"] = n; tl[f"B{rr}"] = s; tl[f"C{rr}"] = dt.date.fromisoformat(d); tl[f"D{rr}"] = t
box(tl, "A5:D1000", inp=True)
for i in range(5, 1001):
    tl[f"E{i}"] = f'=IF(C{i}="","",IF(TODAY()-C{i}>365,"Retrain","OK"))'
    tl[f"C{i}"].number_format = "dd-mmm-yyyy"
box(tl, "E5:E1000", center=True)
tl.conditional_formatting.add("E5:E1000", CellIsRule(operator="equal", formula=['"Retrain"'], fill=PatternFill("solid", fgColor=LEVEL_FILL[0], bgColor=LEVEL_FILL[0]), font=Font(name=F, bold=True, color=LEVEL_TXT[0])))
tl.conditional_formatting.add("E5:E1000", CellIsRule(operator="equal", formula=['"OK"'], fill=PatternFill("solid", fgColor=LEVEL_FILL[3], bgColor=LEVEL_FILL[3]), font=Font(name=F, color=LEVEL_TXT[3])))
for L, w in zip("ABCDE", [22, 24, 14, 16, 10]):
    tl.column_dimensions[L].width = w
dvn2 = DataValidation(type="list", formula1=f"=Employees!$B$5:$B${LAST}", allow_blank=True)
dvk = DataValidation(type="list", formula1="=Skills!$B$5:$B$11", allow_blank=True)
dvd = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True, showErrorMessage=True, error="Enter a date")
for d_ in (dvn2, dvk, dvd):
    tl.add_data_validation(d_)
dvn2.add("A5:A1000"); dvk.add("B5:B1000"); dvd.add("C5:C1000")
tl.freeze_panes = "A5"; tl.auto_filter.ref = "A4:E1000"

# ---------------- Dashboard ----------------
db = wb.create_sheet("Dashboard", 1)
title(db, "Plant-wide skill coverage", "Qualified people (level 3–4) per skill vs target – updates automatically")
head(db, 4, ["Skill", "Target", "Qualified", "Gap", "Status"])
for k in range(7):
    rr = 5 + k; L = SKILL_COLS[k]
    db[f"A{rr}"] = f"=Matrix!{L}4"; db[f"B{rr}"] = f"=Matrix!{L}302"
    db[f"C{rr}"] = f"=Matrix!{L}303"; db[f"D{rr}"] = f"=Matrix!{L}304"; db[f"E{rr}"] = f"=Matrix!{L}305"
box(db, "A5:E11", center=True)
for c in db["A5:A11"]:
    c[0].alignment = Alignment(horizontal="left"); c[0].font = font(color="008000")
db["A13"] = "Skills with a GAP"; db["A13"].font = font(bold=True)
db["B13"] = '=COUNTIF(E5:E11,"GAP")'; db["B13"].font = font(bold=True, size=12, color="C0504D")
db["A14"] = "People on the matrix"; db["A14"].font = font(bold=True)
db["B14"] = f'=COUNTA(Employees!B5:B{LAST})'; db["B14"].font = font(bold=True, size=12)
db.conditional_formatting.add("E5:E11", CellIsRule(operator="equal", formula=['"GAP"'], fill=PatternFill("solid", fgColor=LEVEL_FILL[0], bgColor=LEVEL_FILL[0]), font=Font(name=F, bold=True, color=LEVEL_TXT[0])))
db.conditional_formatting.add("E5:E11", CellIsRule(operator="equal", formula=['"OK"'], fill=PatternFill("solid", fgColor=LEVEL_FILL[3], bgColor=LEVEL_FILL[3]), font=Font(name=F, bold=True, color=LEVEL_TXT[3])))
db.column_dimensions["A"].width = 22
for L in "BCDE":
    db.column_dimensions[L].width = 11
ch = BarChart(); ch.type = "col"; ch.grouping = "clustered"
ch.title = "Plant-wide: qualified people per skill vs target"
ch.y_axis.title = "People"; ch.y_axis.majorGridlines = None
data = Reference(db, min_col=2, max_col=3, min_row=4, max_row=11)
ch.add_data(data, titles_from_data=True)
ch.set_categories(Reference(db, min_col=1, min_row=5, max_row=11))
ch.series[0].graphicalProperties.solidFill = "C9D3CD"; ch.series[0].graphicalProperties.line.solidFill = "C9D3CD"
ch.series[1].graphicalProperties.solidFill = GREEN; ch.series[1].graphicalProperties.line.solidFill = GREEN
ch.dataLabels = DataLabelList(); ch.dataLabels.showVal = True
for _a in ("showSerName", "showCatName", "showLegendKey", "showPercent"):
    setattr(ch.dataLabels, _a, False)
ch.legend.position = "b"; ch.height = 9; ch.width = 20
ch.x_axis.delete = False; ch.y_axis.delete = False
db.add_chart(ch, "G4")

# Unlock input cells so Review > Protect Sheet keeps them editable
from openpyxl.styles import Protection
for sheet, ranges in {em: [f"A5:E{LAST}"], sk: ["A5:D11"], mx: [f"A{FIRST}:A{LAST}", f"C{FIRST}:I{LAST}"],
                      tl: ["A5:D1000"], lg: ["E5:E20"]}.items():
    for rg in ranges:
        for row in sheet[rg]:
            for c in row:
                c.protection = Protection(locked=False)

for s in wb.worksheets:
    s.sheet_view.showGridLines = s.title not in ("Instructions", "Dashboard")
wb.save(OUT)
print("saved", OUT)
