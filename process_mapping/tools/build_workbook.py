"""Build the CHG process mapping Excel workbook (primary data-management format).

Usage:  python build_workbook.py project_dump.json out.xlsx [--demo]

Data sheets use exactly the same headers as the app's Excel export, so the workbook can be
imported into the CHG Process Mapper. Calculation sheets (Step Metrics, KPI Calculations,
PAM Summary, route formula columns) are live Excel formulas.
"""
import json, sys, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter as L
from openpyxl.comments import Comment

src, out = sys.argv[1], sys.argv[2]
DEMO = "--demo" in sys.argv
D = json.load(open(src))
P = D["demo"] if DEMO else D["template"]
CHECK = D["checklistDemo"] if DEMO else D["checklistTemplate"]
FIELDS, SETTINGS, FORMULAS, WASTES = D["fields"], D["settings"], D["formulas"], D["wastes"]
NYC = "Not Yet Collected"
S2T = {"VER": "Verified", "EST": "Estimate", "NYC": NYC}

FONT = "Arial"
HDR = PatternFill("solid", fgColor="13315C")
INP = PatternFill("solid", fgColor="FFF8D6")   # light yellow = cells the team fills in
CALC = PatternFill("solid", fgColor="EEF2F6")
AMB = PatternFill("solid", fgColor="FBE7C2")
thin = Side(style="thin", color="D6DEE8")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = Workbook()
wb.remove(wb.active)

def sheet(name, headers, widths=None, inputs=()):
    ws = wb.create_sheet(name)
    ws.append(headers)
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=i)
        c.font = Font(name=FONT, bold=True, color="FFFFFF"); c.fill = HDR
        c.alignment = Alignment(vertical="center", wrap_text=True)
        ws.column_dimensions[L(i)].width = (widths[i - 1] if widths else max(11, min(40, len(h) + 4)))
    ws.freeze_panes = "A2"; ws.row_dimensions[1].height = 30
    ws._inputs = set(inputs)
    return ws

def style_rows(ws, first=2):
    for row in ws.iter_rows(min_row=first, max_row=ws.max_row):
        for c in row:
            c.font = Font(name=FONT, size=10, color="0000FF" if c.column in ws._inputs else "000000")
            c.border = BOX
            c.alignment = Alignment(vertical="top", wrap_text=isinstance(c.value, str) and len(c.value) > 40)
            if c.column in ws._inputs: c.fill = INP
    if ws.max_row > 1: ws.auto_filter.ref = f"A1:{L(ws.max_column)}{ws.max_row}"

def dv_list(ws, col, values, first=2, last=2000):
    dv = DataValidation(type="list", formula1='"' + ",".join(values) + '"', allow_blank=True, showErrorMessage=True)
    ws.add_data_validation(dv); dv.add(f"{L(col)}{first}:{L(col)}{last}")

def nyc_rule(ws, rng, col_letter, first=2):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'OR(${col_letter}{first}="",${col_letter}{first}="{NYC}")'], fill=AMB))

def nv(v): return "" if v is None or v == "" else v
def step_name(sid):
    for l in P["lines"]:
        for k in ("current", "future"):
            S = l.get(k)
            if S:
                for s in S["steps"]:
                    if s["id"] == sid: return s["name"]
    return ""

# ---------------- README ----------------
ws = wb.create_sheet("README")
ws.column_dimensions["A"].width = 26; ws.column_dimensions["B"].width = 120
rows = [
    ("CHG Process Mapping Workbook", ""),
    ("Site", "C.H. Guenther & Son — Duncanville, TX"),
    ("First deployment", "Nugget Line (model line). Rolled Dough Lines #3 and #4 are included as unverified templates."),
    ("Built", datetime.date.today().isoformat()),
    ("Contents", "DEMONSTRATION DATA — every value marked DEMO is illustrative and is NOT a plant measurement." if DEMO else "Blank template. No measurements, equipment arrangements or savings have been entered. Blank = Not Yet Collected."),
    ("", ""),
    ("How to use", "1) Walk the floor with the Checklist sheet.  2) Enter values in the yellow cells of Settings and Measurements, with Status and Source.  3) Review Step Metrics and KPI Calculations (formulas).  4) Import this workbook into the CHG Process Mapper app to draw and update the maps, or keep working here."),
    ("Cell colors", "Yellow fill + blue text = cells the team fills in.  Grey fill = calculated by formula, do not overwrite.  Amber = Not Yet Collected."),
    ("Status values", "Verified = measured on the floor or taken from a system of record, with the source recorded.  Estimate = assumption, planning value or engineering estimate.  Not Yet Collected = no value yet."),
    ("Never", "Never type an estimate into a value cell without setting Status = Estimate and recording the basis in Source. Savings are reported only after Finance validates them (Finance Validation sheet)."),
    ("IDs", "Keep the ID columns (Line ID, Step ID, Route ID, Opp ID…) unchanged so rows stay linked when imported into the app."),
    ("", ""),
    ("EXAMPLE ROW (format only)", "Measurements: Line ID = (unchanged) · State = Current · Step = Packaging · Field key = cycleTime · Value = 6.3 · Status = Verified · Source = \"J. Smith, 11/18/2026 1st shift, stopwatch 30 cycles\". This example is for format only and is not a measurement."),
    ("", ""),
    ("Sheets", "Settings · Steps · Measurements · Activities · Layouts · Zones · Routes · Flow Nodes · Flow Edges · Swimlane · Opportunities · Actions · Finance Validation · Checklist — data (import into the app).  Step Metrics · KPI Calculations · PAM Summary — formulas.  Formulas — calculation reference."),
]
for r in rows: ws.append(r)
ws["A1"].font = Font(name=FONT, size=16, bold=True, color="13315C")
for r in range(2, ws.max_row + 1):
    ws.cell(row=r, column=1).font = Font(name=FONT, bold=True, size=10)
    ws.cell(row=r, column=2).font = Font(name=FONT, size=10, color="B9770E" if (r == 5 and DEMO) or r == 13 else "000000", bold=(r == 5))
    ws.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical="top")

# ---------------- data sheets (same headers as app export) ----------------
ws = sheet("Lines", ["Line ID", "Line", "Unit", "Has future state", "Notes"], [14, 24, 10, 16, 90], inputs=(2, 3, 5))
for l in P["lines"]: ws.append([l["id"], l["name"], l["unit"], "Yes" if l.get("future") else "No", l["notes"]])
style_rows(ws)

ws = sheet("Settings", ["Line ID", "Key", "Setting", "Unit", "Value", "Status", "Source", "How to collect"], [14, 16, 30, 12, 12, 18, 46, 70], inputs=(5, 6, 7))
for l in P["lines"]:
    for sd in SETTINGS:
        f = l["settings"][sd["k"]]; st = "NYC" if f["v"] is None else f["s"]
        ws.append([l["id"], sd["k"], sd["label"], sd["unit"], nv(f["v"]), S2T[st], f["src"], sd["how"]])
style_rows(ws); dv_list(ws, 6, ["Verified", "Estimate", NYC]); nyc_rule(ws, f"E2:E{ws.max_row}", "E")

def step_rows():
    for l in P["lines"]:
        for k in ("current", "future"):
            S = l.get(k)
            if not S: continue
            yield l, k.capitalize(), S

ws = sheet("Steps", ["Line ID", "State", "Step ID", "Order", "Step", "Department", "Process owner", "Equipment", "Applies to line", "VA class", "Type", "Food safety", "Allergen", "Sanitation", "Traceability", "Quality control", "Notes"],
           [14, 9, 12, 7, 24, 13, 18, 18, 15, 10, 11, 40, 34, 30, 30, 34, 30], inputs=(5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17))
for l, st, S in step_rows():
    for i, s in enumerate(S["steps"], 1):
        q = s["fsq"]
        ws.append([l["id"], st, s["id"], i, s["name"], s["dept"], s["owner"], s["equipment"], s["applic"], s["vaClass"], s["kind"], q["foodSafety"], q["allergen"], q["sanitation"], q["trace"], q["qc"], s["notes"]])
style_rows(ws)
dv_list(ws, 6, ["Production", "Quality", "Maintenance", "Sanitation", "Warehouse", "Planning", "Supervisors"]); dv_list(ws, 9, ["To verify", "Confirmed", "Not applicable"]); dv_list(ws, 10, ["VA", "NNVA", "NVA"]); dv_list(ws, 11, ["operation", "storage", "inspection"])
nyc_rule(ws, f"G2:H{ws.max_row}", "G")

ws = sheet("Measurements", ["Line ID", "State", "Step ID", "Step", "Field key", "Measure", "Unit", "Value", "Status", "Source", "How to measure"], [14, 9, 12, 24, 18, 28, 12, 11, 18, 44, 80], inputs=(8, 9, 10))
for l, st, S in step_rows():
    for s in S["steps"]:
        for fd in FIELDS:
            f = s["f"][fd["k"]]; stt = "NYC" if f["v"] is None else f["s"]
            ws.append([l["id"], st, s["id"], s["name"], fd["k"], fd["label"], fd["unit"].replace("units", l["unit"] + "s"), nv(f["v"]), S2T[stt], f["src"], fd["how"]])
style_rows(ws); dv_list(ws, 9, ["Verified", "Estimate", NYC], last=ws.max_row + 500); nyc_rule(ws, f"H2:H{ws.max_row}", "H")
MEAS_LAST = ws.max_row

ws = sheet("Activities", ["Line ID", "State", "Activity ID", "Step ID", "Step", "Activity", "Type", "Time (min)", "Distance (ft)", "People", "Class", "Status", "Source", "Improvement opportunity"], [14, 9, 12, 12, 22, 40, 7, 11, 12, 9, 8, 18, 30, 40], inputs=(6, 7, 8, 9, 10, 11, 12, 13, 14))
for l, st, S in step_rows():
    for a in S["acts"]:
        t = a["time"]; ws.append([l["id"], st, a["id"], a["stepId"], step_name(a["stepId"]), a["desc"], a["type"], nv(t), nv(a["dist"]), nv(a["people"]), a["va"], S2T["NYC" if t is None else a["status"]], a["src"], a["opportunity"]])
style_rows(ws); dv_list(ws, 7, ["O", "T", "I", "D", "S"]); dv_list(ws, 11, ["VA", "NNVA", "NVA"]); dv_list(ws, 12, ["Verified", "Estimate", NYC]); nyc_rule(ws, f"H2:H{ws.max_row}", "H")
ws.cell(row=1, column=7).comment = Comment("O = Operation, T = Transportation, I = Inspection, D = Delay, S = Storage", "CHG")

ws = sheet("Layouts", ["Line ID", "State", "Width (ft)", "Depth (ft)", "Note"], [14, 9, 11, 11, 110], inputs=(3, 4, 5))
for l, st, S in step_rows(): ws.append([l["id"], st, S["layout"]["width"], S["layout"]["height"], S["layout"]["note"]])
style_rows(ws)
ws = sheet("Zones", ["Line ID", "State", "Zone ID", "Zone", "X (ft)", "Y (ft)", "Width (ft)", "Depth (ft)", "Color"], [14, 9, 12, 30, 9, 9, 11, 11, 10], inputs=(4, 5, 6, 7, 8, 9))
for l, st, S in step_rows():
    for z in S["layout"]["zones"]: ws.append([l["id"], st, z["id"], z["name"], z["x"], z["y"], z["w"], z["h"], z["color"]])
style_rows(ws)

# Routes: distance from the app (drawn) or measured; formula columns compute shift/annual values
ws = sheet("Routes", ["Line ID", "State", "Route ID", "Category", "Route", "Points (x,y ft)", "Trips per shift", "People", "Status", "Distance per trip (ft)", "Distance per shift (ft)", "Time per shift (min)", "Labor hr/yr", "Labor cost $/yr"],
           [14, 9, 12, 11, 44, 34, 13, 9, 12, 14, 15, 14, 12, 14], inputs=(4, 5, 6, 7, 8, 9, 10))
kpi = D["demoKpi"] if DEMO else D["templateKpi"]
for li, l in enumerate(P["lines"]):
    for k in ("current", "future"):
        S = l.get(k)
        if not S: continue
        for r in S["routes"]:
            dist = kpi[li][k]["routes"].get(r["id"])
            ws.append([l["id"], k.capitalize(), r["id"], r["cat"], r["name"], "; ".join(f"{p[0]},{p[1]}" for p in r["points"]), nv(r["trips"]["v"]), nv(r["people"]["v"]), "Verified" if r["status"] == "VER" else "Estimate", round(dist, 1) if dist else ""])
SET = "Settings"
def setting(line_cell, key):
    return f'SUMIFS({SET}!$E:$E,{SET}!$A:$A,{line_cell},{SET}!$B:$B,"{key}")'
for r in range(2, max(ws.max_row, 2) + 40):
    sp = f'IF($D{r}="forklift",{setting(f"$A{r}", "forkliftSpeed")},{setting(f"$A{r}", "walkSpeed")})'
    ws[f"K{r}"] = f'=IF($A{r}="","",IF(AND(ISNUMBER($G{r}),ISNUMBER($J{r})),$G{r}*$J{r},"{NYC}"))'
    ws[f"L{r}"] = f'=IF($A{r}="","",IF(AND(ISNUMBER($K{r}),{sp}>0),$K{r}/{sp},"{NYC}"))'
    ws[f"M{r}"] = f'=IF($A{r}="","",IF(AND(ISNUMBER($L{r}),ISNUMBER($H{r}),{setting(f"$A{r}", "shiftsPerDay")}>0,{setting(f"$A{r}", "daysPerYear")}>0),$L{r}*$H{r}*{setting(f"$A{r}", "shiftsPerDay")}*{setting(f"$A{r}", "daysPerYear")}/60,"{NYC}"))'
    ws[f"N{r}"] = f'=IF($A{r}="","",IF(AND(ISNUMBER($M{r}),{setting(f"$A{r}", "laborRate")}>0),$M{r}*{setting(f"$A{r}", "laborRate")},"{NYC}"))'
style_rows(ws)
for r in range(2, ws.max_row + 1):
    for col in "KLMN": ws[f"{col}{r}"].fill = CALC
    ws[f"K{r}"].number_format = "#,##0"; ws[f"L{r}"].number_format = "#,##0.0"; ws[f"M{r}"].number_format = "#,##0.0"; ws[f"N{r}"].number_format = "$#,##0"
dv_list(ws, 4, ["walk", "forklift", "ingredient", "product"]); dv_list(ws, 9, ["Verified", "Estimate"])
ws.cell(row=1, column=10).comment = Comment("From the route drawn in the app (layout scale 1 unit = 1 ft), or a measuring-wheel distance if no route is drawn.", "CHG")
ROUTES_LAST = ws.max_row

ws = sheet("Flow Nodes", ["Line ID", "State", "Node ID", "Symbol", "Label", "X", "Y", "Department", "Step ID", "Attached to"], [14, 9, 12, 12, 30, 7, 7, 13, 12, 12], inputs=(4, 5, 8))
for l, st, S in step_rows():
    for n in S["flow"]["nodes"]: ws.append([l["id"], st, n["id"], n["kind"], n["label"], n["x"], n["y"], n.get("dept", ""), n.get("stepId", "") or "", n.get("side", "") or ""])
style_rows(ws)
ws = sheet("Flow Edges", ["Line ID", "State", "Edge ID", "From node", "To node", "Label"], [14, 9, 12, 12, 12, 12], inputs=(6,))
for l, st, S in step_rows():
    for e in S["flow"]["edges"]: ws.append([l["id"], st, e["id"], e["from"], e["to"], e["label"]])
style_rows(ws)
ws = sheet("Swimlane", ["Line ID", "State", "Item ID", "Lane", "Column", "Activity", "Type", "Flags", "Next items"], [14, 9, 12, 13, 8, 36, 10, 26, 30], inputs=(4, 5, 6, 7, 8))
for l, st, S in step_rows():
    for i in S["swim"]: ws.append([l["id"], st, i["id"], i["lane"], i["col"], i["label"], i["kind"], ", ".join(k for k, v in (i.get("flags") or {}).items() if v), ", ".join(i.get("next") or [])])
style_rows(ws)
ws.cell(row=1, column=8).comment = Comment("Comma-separated: duplicate, unnecessary, delay, gap", "CHG")

ws = sheet("Opportunities", ["Opp ID", "Line ID", "Step ID", "Step", "Waste", "Opportunity", "Priority", "Owner", "Est. annual savings ($)", "Estimate basis", "Source"], [12, 14, 12, 22, 15, 50, 9, 18, 14, 36, 16], inputs=(5, 6, 7, 8, 9, 10, 11))
for o in P["opps"]: ws.append([o["id"], o["lineId"], o["stepId"], step_name(o["stepId"]), o["waste"], o["desc"], o["priority"], o["owner"], nv(o["estAnnual"]), o["estBasis"], o["source"]])
style_rows(ws); dv_list(ws, 5, [w["k"] for w in WASTES]); dv_list(ws, 7, ["High", "Medium", "Low"])
for r in range(2, 400): ws[f"I{r}"].number_format = "$#,##0"
ws = sheet("Actions", ["Action ID", "Opp ID", "Action", "Owner", "Due date", "Status", "Notes"], [12, 12, 56, 20, 12, 13, 30], inputs=(3, 4, 5, 6, 7))
for a in P["actions"]:
    due = datetime.date.fromisoformat(a["due"]) if a.get("due") else ""
    ws.append([a["id"], a["oppId"], a["action"], a["owner"], due, a["status"], a["notes"]])
style_rows(ws); dv_list(ws, 6, ["Not started", "In progress", "Blocked", "Done"])
for r in range(2, 400): ws[f"E{r}"].number_format = "yyyy-mm-dd"
ws.conditional_formatting.add("A2:G400", FormulaRule(formula=['AND($E2<>"",$E2<TODAY(),$F2<>"Done")'], fill=PatternFill("solid", fgColor="FBE6E3")))
ws = sheet("Finance Validation", ["Opp ID", "Opportunity", "Benefit type", "Baseline", "Valuation method", "Estimated ($/yr)", "Validation status", "Validated ($/yr)", "Finance reviewer", "Date"], [12, 46, 18, 30, 30, 14, 24, 14, 18, 12], inputs=(3, 4, 5, 6, 7, 8, 9, 10))
for o in P["opps"]:
    f = o["fin"]; ws.append([o["id"], o["desc"], f["category"], f["baseline"], f["method"], nv(o["estAnnual"]), f["status"], nv(f["validated"]), f["reviewer"], datetime.date.fromisoformat(f["date"]) if f.get("date") else ""])
style_rows(ws); dv_list(ws, 3, ["Hard savings", "Cost avoidance", "Capacity", "Nonfinancial", "Not yet classified"]); dv_list(ws, 7, ["Not started", "Method agreed with Finance", "In verification", "Validated by Finance", "Rejected"])
for r in range(2, 400): ws[f"F{r}"].number_format = "$#,##0"; ws[f"H{r}"].number_format = "$#,##0"; ws[f"J{r}"].number_format = "yyyy-mm-dd"
ws.cell(row=1, column=8).comment = Comment("Enter only after Finance has validated against actual cost results. Labor counts only when paid hours fall; scrap valued at cost, not selling price; net of one-time costs.", "CHG")

ws = sheet("Checklist", ["Key", "Line", "Area", "Item", "How to measure", "Done"], [10, 16, 26, 44, 90, 8], inputs=(6,))
for l, items in zip(P["lines"], CHECK):
    for i in items: ws.append([i["k"], l["name"], i["area"], i["item"], i["how"], "Yes" if P["checks"].get(i["k"]) else ""])
style_rows(ws); dv_list(ws, 6, ["Yes"], last=ws.max_row)
ws.column_dimensions["A"].hidden = True

# ---------------- Step Metrics (formulas) ----------------
NL = P["lines"][0]
ws = wb.create_sheet("Step Metrics")
hdr = ["Line ID", "State", "Step ID", "Step", "Applies to line", "VA class"] + [f"{f['label']} ({f['unit']})" for f in FIELDS] + ["Effective C/T (s)", "Queue time (min)", "VA processing (min)", "Yield factor", "Changeover hr/week", "Bottleneck?", "ECT (selected line/state)", "Yield (selected line/state)"]
ws.append(hdr); ws.append(["", "", "", "keys →", "", ""] + [f["k"] for f in FIELDS] + [""] * 8)
fc = {f["k"]: L(7 + i) for i, f in enumerate(FIELDS)}
first = 3
rows_ = []
for l, st, S in step_rows():
    for s in S["steps"]: rows_.append((l, st, s))
last = first + len(rows_) - 1
M = "Measurements"
K = "'KPI Calculations'"
for i, (l, st, s) in enumerate(rows_):
    r = first + i
    ws[f"A{r}"] = l["id"]; ws[f"B{r}"] = st; ws[f"C{r}"] = s["id"]
    ws[f"D{r}"] = f'=IFERROR(INDEX(Steps!$E:$E,MATCH(1,INDEX((Steps!$A:$A=$A{r})*(Steps!$B:$B=$B{r})*(Steps!$C:$C=$C{r}),0),0)),"")' if False else s["name"]
    ws[f"E{r}"] = s["applic"]; ws[f"F{r}"] = s["vaClass"]
    for f in FIELDS:
        c = fc[f["k"]]
        crit = f'{M}!$A$2:$A${MEAS_LAST},$A{r},{M}!$B$2:$B${MEAS_LAST},$B{r},{M}!$C$2:$C${MEAS_LAST},$C{r},{M}!$E$2:$E${MEAS_LAST},{c}$2'
        ws[f"{c}{r}"] = f'=IF(COUNTIFS({crit},{M}!$H$2:$H${MEAS_LAST},"<>")=0,"{NYC}",SUMIFS({M}!$H$2:$H${MEAS_LAST},{crit}))'
    ct, up, pt, wip, wt, y, sc, co, cpw = (fc[k] for k in ["cycleTime", "uptime", "processTime", "wip", "waitTime", "yield", "scrap", "changeover", "changeoversPerWeek"])
    na = f'$E{r}="Not applicable"'
    dem = f'{setting(f"$A{r}", "demandPerDay")}'
    avail = f'({setting(f"$A{r}", "shiftsPerDay")}*({setting(f"$A{r}", "minPerShift")}-{setting(f"$A{r}", "breakMin")}))'
    ws[f"W{r}"] = f'=IF({na},"",IF(ISNUMBER({ct}{r}),{ct}{r}/IF(ISNUMBER({up}{r}),{up}{r}/100,1),"{NYC}"))'
    ws[f"X{r}"] = f'=IF({na},"",IF(AND(ISNUMBER({wip}{r}),{dem}>0,{avail}>0),{wip}{r}/{dem}*{avail},IF(ISNUMBER({wt}{r}),{wt}{r},"{NYC}")))'
    ws[f"Y{r}"] = f'=IF(OR({na},$F{r}<>"VA"),"",IF(ISNUMBER({pt}{r}),{pt}{r},"{NYC}"))'
    ws[f"Z{r}"] = f'=IF({na},"",IF(ISNUMBER({y}{r}),{y}{r}/100,IF(ISNUMBER({sc}{r}),1-{sc}{r}/100,"{NYC}")))'
    ws[f"AA{r}"] = f'=IF({na},"",IF(AND(ISNUMBER({co}{r}),ISNUMBER({cpw}{r})),{co}{r}*{cpw}{r}/60,""))'
    ws[f"AC{r}"] = f'=IF(AND($A{r}={K}!$C$4,$B{r}={K}!$C$5,ISNUMBER(W{r})),W{r},"")'
    ws[f"AB{r}"] = f'=IF(AND(ISNUMBER(AC{r}),AC{r}=MAX($AC${first}:$AC${last})),"BOTTLENECK","")'
    ws[f"AD{r}"] = f'=IF(AND($A{r}={K}!$C$4,$B{r}={K}!$C$5,ISNUMBER(Z{r})),Z{r},"")'
for i, h in enumerate(hdr, 1):
    c = ws.cell(row=1, column=i); c.font = Font(name=FONT, bold=True, color="FFFFFF"); c.fill = HDR; c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.cell(row=2, column=i).font = Font(name=FONT, size=8, color="8A97A8")
    ws.column_dimensions[L(i)].width = 24 if i == 4 else 13
ws.row_dimensions[1].height = 44; ws.freeze_panes = "E3"
for row in ws.iter_rows(min_row=first, max_row=last):
    for c in row:
        c.font = Font(name=FONT, size=10); c.border = BOX
        if c.column >= 7: c.fill = CALC; c.number_format = "#,##0.00"
ws.conditional_formatting.add(f"G{first}:AA{last}", FormulaRule(formula=[f'G{first}="{NYC}"'], fill=AMB))
ws.conditional_formatting.add(f"AB{first}:AB{last}", FormulaRule(formula=[f'AB{first}="BOTTLENECK"'], fill=PatternFill("solid", fgColor="F6C9C3")))
SM = f"'Step Metrics'"
rng = lambda col: f"{SM}!${col}${first}:${col}${last}"

# ---------------- KPI Calculations ----------------
ws = wb.create_sheet("KPI Calculations")
ws.column_dimensions["A"].width = 4; ws.column_dimensions["B"].width = 38; ws.column_dimensions["C"].width = 20; ws.column_dimensions["D"].width = 16; ws.column_dimensions["E"].width = 90
ws["B1"] = "KPI calculations (live formulas)"; ws["B1"].font = Font(name=FONT, size=15, bold=True, color="13315C")
ws["B2"] = "Change the Line ID / State to calculate another line. Results show Not Yet Collected until their inputs are entered." ; ws["B2"].font = Font(name=FONT, size=10, color="5E6B7A")
ws["B4"] = "Line ID"; ws["C4"] = NL["id"]; ws["D4"] = f'=IFERROR(INDEX(Lines!$B:$B,MATCH(C4,Lines!$A:$A,0)),"Unknown line")'
ws["B5"] = "State"; ws["C5"] = "Current"
for c in ("C4", "C5"): ws[c].fill = INP; ws[c].font = Font(name=FONT, color="0000FF", bold=True)
dv = DataValidation(type="list", formula1='"Current,Future"'); ws.add_data_validation(dv); dv.add("C5")
dvl = DataValidation(type="list", formula1=f"=Lines!$A$2:$A${len(P['lines']) + 1}"); ws.add_data_validation(dvl); dvl.add("C4")
row = 7
def hdr_row(t):
    global row
    ws[f"B{row}"] = t; ws[f"B{row}"].font = Font(name=FONT, bold=True, color="FFFFFF")
    for col in "BCDE": ws[f"{col}{row}"].fill = HDR
    ws[f"C{row}"] = "Value"; ws[f"D{row}"] = "Unit"; ws[f"E{row}"] = "Formula / note"
    for col in "CDE": ws[f"{col}{row}"].font = Font(name=FONT, bold=True, color="FFFFFF")
    row += 1
REF = {}
def put(key, label, formula, unit, note, fmt="#,##0.00"):
    global row
    ws[f"B{row}"] = label; ws[f"C{row}"] = formula; ws[f"D{row}"] = unit; ws[f"E{row}"] = note
    ws[f"C{row}"].number_format = fmt; ws[f"C{row}"].fill = CALC
    for col in "BCDE": ws[f"{col}{row}"].font = Font(name=FONT, size=10); ws[f"{col}{row}"].border = BOX
    ws[f"E{row}"].alignment = Alignment(wrap_text=True)
    REF[key] = f"$C${row}"; row += 1
def sget(key):
    return f'IF(COUNTIFS({SET}!$A:$A,$C$4,{SET}!$B:$B,"{key}",{SET}!$E:$E,"<>")=0,"{NYC}",SUMIFS({SET}!$E:$E,{SET}!$A:$A,$C$4,{SET}!$B:$B,"{key}"))'
hdr_row("Line settings (from Settings sheet)")
for sd in SETTINGS: put(sd["k"], sd["label"], "=" + sget(sd["k"]), sd["unit"], sd["how"], "#,##0.00")
row += 1; hdr_row("Flow, time and efficiency")
crit = f'{SM}!$A${first}:$A${last},$C$4,{SM}!$B${first}:$B${last},$C$5'
applic_crit = crit + f',{SM}!$E${first}:$E${last},"<>Not applicable"'
R = REF
N = lambda *ks: "AND(" + ",".join(f"ISNUMBER({R[k]})" for k in ks) + ")"
put("avail", "Available time per day", f'=IF({N("shiftsPerDay", "minPerShift", "breakMin")},{R["shiftsPerDay"]}*({R["minPerShift"]}-{R["breakMin"]}),"{NYC}")', "min/day", FORMULAS[0][1], "#,##0")
put("takt", "Takt time", f'=IF(AND(ISNUMBER({R["avail"]}),ISNUMBER({R["demandPerDay"]})),IF({R["demandPerDay"]}>0,{R["avail"]}*60/{R["demandPerDay"]},"{NYC}"),"{NYC}")', "s/unit", FORMULAS[1][1])
put("nsteps", "Steps that apply to the line", f"=COUNTIFS({applic_crit})", "steps", "Steps not marked Not applicable", "0")
put("missQ", "Steps missing queue time", f'=COUNTIFS({applic_crit},{rng("X")},"{NYC}")', "steps", "Needs WIP (with demand and shift settings) or observed waiting time", "0")
put("missP", "Steps missing processing time", f'=COUNTIFS({applic_crit},{rng("H")},"{NYC}")', "steps", "", "0")
put("lt", "Total lead time", f'=IF(COUNTIFS({applic_crit},{rng("X")},">=0")+COUNTIFS({applic_crit},{rng("H")},">=0")=0,"{NYC}",SUMIFS({rng("X")},{applic_crit})+SUMIFS({rng("H")},{applic_crit}))', "min", FORMULAS[6][1] + " — incomplete while any step above is missing data")
put("lth", "Total lead time (hours)", f'=IF(ISNUMBER({R["lt"]}),{R["lt"]}/60,"{NYC}")', "hr", "")
put("va", "Value-added time", f'=IF(COUNTIFS({crit},{rng("Y")},">=0")=0,"{NYC}",SUMIFS({rng("Y")},{crit}))', "min", FORMULAS[7][1])
put("nva", "Non-value-added time", f'=IF(AND(ISNUMBER({R["lt"]}),ISNUMBER({R["va"]})),{R["lt"]}-{R["va"]},"{NYC}")', "min", FORMULAS[8][1])
put("pce", "Process cycle efficiency", f'=IF(AND(ISNUMBER({R["lt"]}),ISNUMBER({R["va"]})),IF({R["lt"]}>0,{R["va"]}/{R["lt"]},"{NYC}"),"{NYC}")', "%", FORMULAS[9][1], "0.00%")
put("status", "Lead time data status", f'=IF({R["missQ"]}+{R["missP"]}>0,"Incomplete — "&({R["missQ"]}+{R["missP"]})&" inputs missing","Complete — check Status column in Measurements for estimates")', "", "")
row += 1; hdr_row("Bottleneck, capacity and quality")
put("ect", "Bottleneck effective cycle time", f'=IF(COUNT({rng("AC")})=0,"{NYC}",MAX({rng("AC")}))', "s/unit", FORMULAS[2][1] + ". Provisional until every step is timed.")
put("bn", "Bottleneck step", f'=IF(ISNUMBER({R["ect"]}),INDEX({rng("D")},MATCH({R["ect"]},{rng("AC")},0)),"{NYC}")', "", FORMULAS[3][1], "@")
put("timed", "Steps with cycle time", f'=COUNTIFS({applic_crit},{rng("W")},">=0")', "steps", "Compare with steps that apply", "0")
put("cap", "Line capacity", f'=IF(ISNUMBER({R["ect"]}),3600/{R["ect"]},"{NYC}")', "units/hr", FORMULAS[4][1], "#,##0")
put("rty", "Rolled throughput yield", f'=IF(COUNT({rng("AD")})=0,"{NYC}",PRODUCT({rng("AD")}))', "%", FORMULAS[10][1] + ". Partial while some steps lack yield/scrap.", "0.00%")
put("co", "Changeover time", f'=IF(COUNTIFS({crit},{rng("AA")},">=0")=0,"{NYC}",SUMIFS({rng("AA")},{crit}))', "hr/week", FORMULAS[11][1])
put("labor", "Total labor", f'=IF(COUNTIFS({applic_crit},{rng("M")},">=0")=0,"{NYC}",SUMIFS({rng("M")},{applic_crit}))', "people", "Σ labor over applicable steps")
put("prod", "Labor productivity", f'=IF(AND(ISNUMBER({R["cap"]}),ISNUMBER({R["labor"]})),IF({R["labor"]}>0,{R["cap"]}/{R["labor"]},"{NYC}"),"{NYC}")', "units/labor-hr", FORMULAS[20][1])
row += 1; hdr_row("Travel (Routes sheet)")
rc = f"Routes!$A$2:$A${ROUTES_LAST + 40},$C$4,Routes!$B$2:$B${ROUTES_LAST + 40},$C$5"
put("thr", "Travel labor", f'=IF(COUNTIFS({rc},Routes!$M$2:$M${ROUTES_LAST + 40},">=0")=0,"{NYC}",SUMIFS(Routes!$M$2:$M${ROUTES_LAST + 40},{rc}))', "hr/yr", FORMULAS[14][1])
put("tcost", "Travel labor cost", f'=IF(COUNTIFS({rc},Routes!$N$2:$N${ROUTES_LAST + 40},">=0")=0,"{NYC}",SUMIFS(Routes!$N$2:$N${ROUTES_LAST + 40},{rc}))', "$/yr", FORMULAS[15][1] + " — estimate until Finance validates", "$#,##0")
row += 1; hdr_row("Savings pipeline (all lines)")
FV = "'Finance Validation'"
put("est", "Estimated, not yet validated", f'=SUMIFS({FV}!$F:$F,{FV}!$G:$G,"<>Validated by Finance",{FV}!$G:$G,"<>Rejected")', "$/yr", "Not reported as savings", "$#,##0")
put("val", "Validated by Finance (all categories)", f'=SUMIFS({FV}!$H:$H,{FV}!$G:$G,"Validated by Finance")', "$/yr", "", "$#,##0")
put("hard", "Validated hard savings", f'=SUMIFS({FV}!$H:$H,{FV}!$G:$G,"Validated by Finance",{FV}!$C:$C,"Hard savings")', "$/yr", "The only figure reported as savings", "$#,##0")
put("over", "Overdue actions", f'=COUNTIFS(Actions!$E:$E,"<"&TODAY(),Actions!$F:$F,"<>Done")', "actions", "Due date before today and not Done", "0")
ws.conditional_formatting.add(f"C8:C{row}", FormulaRule(formula=[f'C8="{NYC}"'], fill=AMB))
ws.freeze_panes = "B7"

# ---------------- PAM Summary ----------------
ws = wb.create_sheet("PAM Summary")
ws.append(["Activity type", "Code", "Activities", "Time (min)", "Distance (ft)", "Not timed"])
A = "Activities"
for code, name in [("O", "Operation"), ("T", "Transportation"), ("I", "Inspection"), ("D", "Delay"), ("S", "Storage")]:
    r = ws.max_row + 1
    c = f'{A}!$A:$A,{K}!$C$4,{A}!$B:$B,{K}!$C$5,{A}!$G:$G,$B{r}'
    ws.append([name, code, f"=COUNTIFS({c})", f"=SUMIFS({A}!$H:$H,{c})", f"=SUMIFS({A}!$I:$I,{c})", f'=COUNTIFS({c},{A}!$H:$H,"")'])
r = ws.max_row + 1
ws.append(["Total", "", f"=SUM(C2:C{r - 1})", f"=SUM(D2:D{r - 1})", f"=SUM(E2:E{r - 1})", f"=SUM(F2:F{r - 1})"])
c = f'{A}!$A:$A,{K}!$C$4,{A}!$B:$B,{K}!$C$5'
ws.append(["Non-value-added time", "NVA", "", f'=SUMIFS({A}!$H:$H,{c},{A}!$K:$K,"NVA")', "", ""])
ws.append(["Value-added share", "VA", "", f'=IF(D{r}>0,SUMIFS({A}!$H:$H,{c},{A}!$K:$K,"VA")/D{r},"{NYC}")', "", ""])
ws.append(["Line / state shown", "", f"={K}!$D$4&\" — \"&{K}!$C$5", "", "", ""])
for i, w in enumerate([26, 8, 12, 12, 13, 11], 1): ws.column_dimensions[L(i)].width = w
for i in range(1, 7): c_ = ws.cell(row=1, column=i); c_.font = Font(name=FONT, bold=True, color="FFFFFF"); c_.fill = HDR
for row_ in ws.iter_rows(min_row=2, max_row=ws.max_row):
    for c_ in row_: c_.font = Font(name=FONT, size=10, bold=c_.row == r); c_.border = BOX; c_.fill = CALC if c_.column > 2 else PatternFill()
ws[f"D{r + 2}"].number_format = "0.0%"

# ---------------- Formulas ----------------
ws = sheet("Formulas", ["Calculation", "Formula", "Unit"], [38, 110, 14])
for f in FORMULAS: ws.append(f)
style_rows(ws)

order = ["README", "Checklist", "KPI Calculations", "Step Metrics", "PAM Summary", "Lines", "Settings", "Steps", "Measurements", "Activities", "Routes", "Layouts", "Zones", "Flow Nodes", "Flow Edges", "Swimlane", "Opportunities", "Actions", "Finance Validation", "Formulas"]
wb._sheets = [wb[n] for n in order]
for n, color in [("README", "13315C"), ("Checklist", "B9770E"), ("KPI Calculations", "2F9E6E"), ("Step Metrics", "2F9E6E"), ("PAM Summary", "2F9E6E")]: wb[n].sheet_properties.tabColor = color
wb.active = 0
wb.save(out)
print("saved", out)
