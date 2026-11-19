"""Builds CHG_Supervisor_90-Day_Projects.xlsx – 3-month improvement projects for Nugget Line,
Mixers & Rolled Dough and Packaging, with weekly KPI tracking, savings and evidence.

Usage: python build_90day.py <out_dir>
"""
import sys
import random
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from xlcommon import *  # noqa
from projects90 import P, GROUPS, DEFAULT_SLOTS

S_D, S_M, S_P, S_L, S_H = "Dashboard", "Project Menu", "People", "Lists", "How To Use"
GCOL = {"NL": "C55A11", "MR": "1F4E79", "PK": "2E7D32"}
SLOTS = [(g, i) for g in ["NL", "MR", "PK"] for i in range(1, 5)]
STATUS = {"On Track": "Green", "Behind": "Yellow", "Off Track": "Red", "Data Missing": "Yellow", "Not Started": "Gray",
          "Inactive": "Gray"}
WK0, NW = 25, 13           # first week row, number of weeks
AP0 = 58                   # first action-plan row


def build(demo, path):
    random.seed(5)
    AS_OF = dt.date(2026, 11, 18)
    wb = Workbook()
    wb.remove(wb.active)
    names = [S_D, S_M] + [f"{g}-{i}" for g, i in SLOTS] + [S_P, S_L, S_H]
    sh = {n: wb.create_sheet(n) for n in names}
    for n in names:
        sh[n].sheet_properties.tabColor = GCOL[n[:2]] if n[:2] in GCOL and n[2] == "-" else (NAVY if n in (S_D, S_M) else "7F7F7F")
    tag = "DEMONSTRATION DATA – fictitious" if demo else "C.H. Guenther & Son — Duncanville"

    # ---------------- Lists ----------------
    ws = sh[S_L]
    title_band(ws, "Lists & Settings", "Settings and dropdown values.", 14, legend=False)
    sets = [("Status date (blank = today)", AS_OF if demo else None, DATE, None),
            ("ƒ Status date used", "=IF(ISNUMBER(C5),C5,TODAY())", DATE, "AsOf90"),
            ("'On target' tolerance above target path", 0.05, PCT, "set_Tol"),
            ("Status: On Track if achieved ≥ expected −", 0.15, PCT, "set_Green"),
            ("Status: Behind if achieved ≥ expected −", 0.35, PCT, "set_Yellow")]
    for i, (lab, v, fmt, nm) in enumerate(sets):
        r = 5 + i
        ws.cell(r, 2, lab).font = f(9, lab.startswith("ƒ"))
        c = ws.cell(r, 3, v)
        c.border = BORDER
        c.alignment = Alignment(horizontal="center")
        c.font = f(9, False, "000000" if lab.startswith("ƒ") else "0000FF")
        c.fill = fill(LGRAY if lab.startswith("ƒ") else WHITE)
        c.number_format = fmt
        if nm:
            define(wb, nm, f"{qs(ws)}!$C${r}")
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 14
    lists = [("E", "Yes / No", ["Yes", "No"]), ("F", "No (until evidence)", ["No"]), ("G", "Shift", ["A", "B", "C", "All"]),
             ("H", "Role", ["Supervisor", "Team Leader", "Coach / CI", "Manager", "Finance", "QA"]),
             ("I", "Group", list(GROUPS.values()))]
    for g in ["NL", "MR", "PK"]:
        lists.append(({"NL": "J", "MR": "K", "PK": "L"}[g], f"{g} projects", [p[0] for p in P if p[1] == g]))
    for colL, h, vals in lists:
        col = ws[colL + "1"].column
        hdr_row(ws, 11, col, [h])
        ws.column_dimensions[colL].width = 14
        for i in range(12):
            c = ws.cell(12 + i, col, vals[i] if i < len(vals) else None)
            c.border = BORDER
            c.font = f(9)
    for nm, colL in [("l9_YesNo", "E"), ("l9_Shift", "G"), ("l9_Role", "H"), ("l9_Group", "I"),
                     ("l9_NL", "J"), ("l9_MR", "K"), ("l9_PK", "L")]:
        dyn_list(wb, nm, ws, colL, 12, 23)
    define(wb, "l9_No", f"{qs(ws)}!$F$12:$F$12")
    hdr_row(ws, 11, 14, ["People list (auto)"], calc=True)
    ws.column_dimensions["N"].width = 22
    for k in range(60):
        ws.cell(12 + k, 14, f'=IFERROR(INDEX(PP_Name,MATCH({k + 1},PP_Seq,0))&"","")').font = f(8, False, GRAY)
    dyn_list(wb, "l9_People", ws, "N", 12, 71, text_only=True)
    # chart data block (P..S)
    hdr_row(ws, 11, 16, ["Project", "Improvement achieved", "Savings to date", "Group"], calc=True)
    for k, (g, i) in enumerate(SLOTS):
        r = 12 + k
        S = qs(sh[f"{g}-{i}"])
        ws.cell(r, 16, f'="{g}-{i} "&{S}!$C$7')
        ws.cell(r, 17, f'=IF(ISNUMBER({S}!$L$7),{S}!$L$7,0)').number_format = PCT
        ws.cell(r, 18, f'=N({S}!$H$11)').number_format = "$#,##0"
        ws.cell(r, 19, GROUPS[g])
    for k, g in enumerate(["NL", "MR", "PK"]):
        r = 26 + k
        ws.cell(r, 16, GROUPS[g])
        ws.cell(r, 18, f"=SUM(R{12 + 4 * k}:R{15 + 4 * k})").number_format = "$#,##0"
    for c in range(16, 20):
        ws.column_dimensions[L(c)].width = 16

    # ---------------- People ----------------
    ws = sh[S_P]
    title_band(ws, "People", "Supervisors, coaches and Finance. Each name appears in the dropdowns once Active? = Yes.", 8)
    ps = [dict(key="Name", header="Name", width=26), dict(key="Role", header="Role", width=16, dv="l9_Role"),
          dict(key="Group", header="Group", width=22, dv="l9_Group"), dict(key="Shift", header="Shift", width=8, dv="l9_Shift"),
          dict(key="Active", header="Active?", width=9, dv="l9_YesNo", default="Yes")]
    pc = Cols(ps)
    ps.append(dict(key="Seq", header="#", width=5, calc=True, fmt="0",
                   formula=f'=IF(AND({{Name}}<>"",{{Active}}<>"No"),COUNTIFS(${pc.letter["Name"]}$6:{{Name}},"?*",${pc.letter["Active"]}$6:{{Active}},"<>No"),"")'))
    pdata = []
    if demo:
        for k, (g, i) in enumerate(SLOTS):
            pdata.append(dict(Name=f"DEMO Supervisor {g}{i}", Role="Supervisor", Group=GROUPS[g], Shift="ABC"[i % 3]))
        pdata += [dict(Name="DEMO Coach", Role="Coach / CI", Group=None, Shift="All"),
                  dict(Name="DEMO Controller", Role="Finance", Group=None, Shift="All")]
    build_table(ws, wb, Cols(ps), 5, 60, "tblPeople90", pdata, name_prefix="PP", freeze_col="C")

    # ---------------- Project Menu ----------------
    ws = sh[S_M]
    title_band(ws, "Project Menu – 90-day projects", "Small, low-cost projects that replace an old habit with a new standard. "
               "Pick one on a project sheet (cell C6). CI recommendation – adapt with leadership, QA and Finance.", 30)
    ms = [dict(key="Code", header="Code", width=7), dict(key="Group", header="Group", width=14),
          dict(key="Name", header="Project", width=30, wrap=True),
          dict(key="Old", header="OLD habit", width=28, wrap=True), dict(key="New", header="NEW habit", width=30, wrap=True),
          dict(key="Why", header="Why it saves money", width=30, wrap=True),
          dict(key="KPI", header="KPI (lower is better)", width=20, wrap=True), dict(key="Unit", header="Unit", width=7),
          dict(key="Vol", header="Volume basis (per week)", width=16, wrap=True),
          dict(key="Price", header="$ basis", width=20, wrap=True),
          dict(key="Pct", header="Default target reduction", width=9, fmt=PCT),
          dict(key="Level", header="Difficulty", width=10), dict(key="Pillar", header="GPS Pillar", width=10)]
    for m in range(1, 4):
        for j in range(1, 4):
            ms.append(dict(key=f"A{(m - 1) * 3 + j}", header=f"Month {m} – action {j}", width=28, wrap=True))
    ms += [dict(key="Routine", header="New standard (daily / shift routine)", width=30, wrap=True),
           dict(key="Evidence", header="Evidence to upload", width=26, wrap=True)]
    mdata = []
    for p in P:
        row = dict(Code=p[0], Group=GROUPS[p[1]], Name=p[2], Old=p[3], New=p[4], Why=p[5], KPI=p[6], Unit=p[7], Vol=p[8],
                   Price=p[9], Pct=p[10], Level=p[11], Pillar=p[12], Routine=p[14], Evidence=p[15])
        for k, a in enumerate(p[13], 1):
            row[f"A{k}"] = a
        mdata.append(row)
    build_table(ws, wb, Cols(ms), 5, 30, "tblMenu90", mdata, name_prefix="MN", freeze_col="D", row_height=75)
    rec = {c for v in DEFAULT_SLOTS.values() for c in v}
    ws.cell(4, 2, "Pre-selected on the project sheets: " + ", ".join(sorted(rec)) +
            ".  The 5th project in each group is a spare choice.").font = f(8, True, BURG)

    # ---------------- Project sheets ----------------
    for g, i in SLOTS:
        build_slot(wb, sh[f"{g}-{i}"], g, i, demo, AS_OF, tag)

    build_dash(wb, sh, demo)

    # ---------------- How To Use ----------------
    ws = sh[S_H]
    title_band(ws, "How To Use", "90-day projects: small, simple, measurable – change one habit and prove the savings.", 6, legend=False)
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 110
    rows = [
        ("H", "SET UP (once)"),
        ("1. People", "Add supervisors, coaches and Finance on 'People'."),
        ("2. Choose the project", "Each group has 4 project sheets (NL-1…4 Nugget Line, MR-1…4 Mixers & Rolled Dough, PK-1…4 Packaging). A recommended project is pre-selected in C6 – change it from the dropdown (menu on 'Project Menu'). Set Active? = No for slots you do not use."),
        ("3. Fill the setup", "Supervisor, coach, shift, start date (Monday of week 1), baseline (average of the 2 weeks BEFORE the start), volume per week and $ per unit (ask Finance). Target = baseline − default reduction, or type your own target."),
        ("H", "EVERY WEEK (5 minutes)"),
        ("4. Enter the KPI", "Type the week's average KPI in 'Actual KPI'. Write what you did and paste the evidence link."),
        ("5. Do the actions", "Work through the Month 1 → 2 → 3 action plan. Paste the evidence link, then mark Done = Yes (Yes is only available once a link is pasted)."),
        ("6. Review", "Dashboard in the weekly meeting: improvement achieved, savings and status for every project and group."),
        ("H", "HOW NUMBERS WORK"),
        ("KPIs", "Every KPI is 'lower is better' (scrap lbs, minutes, batches, cases, oz overweight)."),
        ("Target path", "Week 1 = baseline, then a straight line to the target by week 8; weeks 9–13 hold the target (sustain)."),
        ("Weekly savings", "(Baseline − Actual) × volume per week × $ per unit. Cumulative savings add up the weeks."),
        ("Latest (4-wk avg)", "Average of the last 4 weeks entered – smooths out one good or bad week."),
        ("Improvement achieved", "(Baseline − latest 4-wk avg) ÷ (Baseline − Target). 100% = target reached."),
        ("Annual run-rate", "(Baseline − latest 4-wk avg) × volume × $ × 52 weeks. Estimate until Finance confirms."),
        ("Status", "On Track (green) = achieved within 15 pts of what the plan expects for that week · Behind (amber) within 35 pts · Off Track (red) · Data Missing (amber) = no KPI for 2+ weeks · Not Started (gray)."),
        ("Phases", "Month 1 (weeks 1–4) measure & standardize · Month 2 (weeks 5–8) implement · Month 3 (weeks 9–13) sustain & verify."),
        ("H", "RULES THAT MAKE IT STICK"),
        ("Keep it small", "One line, one KPI, one habit. No capital spend; fixes the crew can do."),
        ("Make it visible", "Post the KPI on the line board every shift. Talk about it at the start of every shift."),
        ("Food safety first", "Weight / temperature / rework changes are agreed with QA before use. Never go below label minimums."),
        ("Savings", "Use real costs from Finance for $ per unit. Treat run-rate as an estimate until Finance confirms."),
    ]
    r = 5
    for a_, b_ in rows:
        if a_ == "H":
            section(ws, r, 2, b_, 2)
            r += 1
            continue
        ws.cell(r, 2, a_).font = f(10, True, NAVY)
        ws.cell(r, 2).alignment = Alignment(vertical="top")
        c = ws.cell(r, 3, b_)
        c.font = f(10)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 18 if len(b_) < 110 else 32
        r += 1
    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


def lk(field):
    return f'IFERROR(INDEX(MN_{field},MATCH($C$6,MN_Code,0))&"","")'


def build_slot(wb, ws, g, i, demo, AS_OF, tag):
    code = DEFAULT_SLOTS[g][i - 1]
    proj = next(p for p in P if p[0] == code)
    title_band(ws, f"{g}-{i} · {GROUPS[g]} – 90-day project", f"{tag}  ·  Pick the project in C6. White cells = input.", 13)
    widths = {"A": 2, "B": 24, "C": 12, "D": 24, "E": 11, "F": 11, "G": 10, "H": 12, "I": 12, "J": 12, "K": 30, "L": 26, "M": 12}
    for k, v in widths.items():
        ws.column_dimensions[k].width = v
    # habit banner
    ws.merge_cells("B4:M4")
    ws["B4"] = f'="OLD: "&{lk("Old")}&"     ▶     NEW: "&{lk("New")}'
    ws["B4"].font = f(10, True, WHITE)
    ws["B4"].fill = fill(GCOL[g])
    ws["B4"].alignment = Alignment(wrap_text=True, vertical="center", indent=1)
    ws.row_dimensions[4].height = 34
    section(ws, 5, 2, "Project setup", 5)
    setup = [
        (6, "Project (select)", code, None, f"l9_{g}"),
        (7, "ƒ Project", "=" + lk("Name"), None, None),
        (8, "ƒ Why it saves money", "=" + lk("Why"), None, None),
        (9, "ƒ KPI (lower is better)", f'={lk("KPI")}&"  ("&{lk("Unit")}&")"', None, None),
        (10, "ƒ New standard", "=" + lk("Routine"), None, None),
        (11, "Supervisor", f"DEMO Supervisor {g}{i}" if demo else None, None, "l9_People"),
        (12, "Coach / Manager", "DEMO Coach" if demo else None, None, "l9_People"),
        (13, "Shift", "ABC"[i % 3] if demo else None, None, "l9_Shift"),
        (14, "Start date (Monday, week 1)", dt.date(2026, 8, 24) + dt.timedelta(days=7 * (i - 1)) if demo else None, DATE, "date"),
        (15, "Baseline (avg of 2 weeks before start)", None, "#,##0.00", "num"),
        (16, "Target – your own (optional)", None, "#,##0.00", "num"),
        (17, "ƒ Target used", f'=IF(ISNUMBER(C16),C16,IF(ISNUMBER(C15),C15*(1-IFERROR(INDEX(MN_Pct,MATCH($C$6,MN_Code,0)),0)),""))', "#,##0.00", None),
        (18, "=\"Volume per week – \"&" + lk("Vol"), None, "#,##0", "num"),
        (19, "=\"$ per unit – \"&" + lk("Price"), None, "$#,##0.00", "num"),
        (20, "Active?", "Yes", None, "l9_YesNo"),
        (21, "ƒ Default target reduction", f'=IFERROR(INDEX(MN_Pct,MATCH($C$6,MN_Code,0)),"")', PCT, None),
    ]
    base_demo = {"lbs": 40, "min": 35, "oz": 0.6, "batches": 6, "cases": 30}[proj[7]]
    vol_demo = {"startups per week": 10, "shifts per week": 10, "packages per week": 150000, "changeovers per week": 6}.get(proj[8], 1)
    price_demo = {"lbs": 0.9, "min": 12.0, "oz": 0.05, "batches": 180.0, "cases": 2.5}[proj[7]]
    for r, lab, val, fmt, dv in setup:
        lc = ws.cell(r, 2, lab)
        lc.font = f(9, True, NAVY)
        lc.alignment = Alignment(wrap_text=True, vertical="center")
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
        if demo:
            if r == 15:
                val = base_demo
            elif r == 18:
                val = vol_demo
            elif r == 19:
                val = price_demo
        c = ws.cell(r, 3, val)
        calc = isinstance(lab, str) and lab.startswith("ƒ")
        c.font = f(9, r in (6, 7), "000000" if calc else "0000FF")
        c.fill = fill(LGRAY if calc else WHITE)
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")
        for cc in range(3, 7):
            ws.cell(r, cc).border = BORDER
        if fmt:
            c.number_format = fmt
        if dv:
            if dv == "date":
                d = DataValidation(type="date", operator="between", formula1="DATE(2020,1,1)", formula2="DATE(2040,12,31)", allow_blank=True)
            elif dv == "num":
                d = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True)
            else:
                d = DataValidation(type="list", formula1=f"={dv}", allow_blank=True)
            ws.add_data_validation(d)
            d.add(f"C{r}")
        ws.row_dimensions[r].height = 30 if r in (8, 10) else 18
    # helpers (hidden col R)
    F_, B_, Q_ = f"$F${WK0}:$F${WK0 + NW - 1}", f"$B${WK0}:$B${WK0 + NW - 1}", f"$Q${WK0}:$Q${WK0 + NW - 1}"
    P_ = f"$P${WK0}:$P${WK0 + NW - 1}"
    ws["R6"] = f"=MAX({Q_})"                                          # last week with data
    ws["R7"] = "=MIN(1,MAX(0,(R6-1)/7))"                               # expected share of improvement
    ws["R8"] = f'=COUNTIF($J${AP0}:$J${AP0 + 8},"Yes")'                # actions done
    ws["R9"] = f'=IF(NOT(ISNUMBER(C14)),0,MIN(13,MAX(0,INT((AsOf90-C14)/7)+1)))'  # current week
    ws.column_dimensions["R"].hidden = True
    # KPI tiles (H..M)
    tile(ws, 6, 8, "CURRENT WEEK", '=IF(R9=0,"–",R9&" of 13")', None, '=IF(R9=0,"set start date",IF(R9<=4,"Month 1 – measure & standardize",IF(R9<=8,"Month 2 – implement","Month 3 – sustain & verify")))', 2, NAVY)
    tile(ws, 6, 10, "LATEST (4-WK AVG)", f'=IF(COUNT({F_})=0,"–",AVERAGEIFS({F_},{P_},"<=4"))', "#,##0.00",
         f'="baseline "&IF(ISNUMBER(C15),TEXT(C15,"#,##0.00"),"–")&" → target "&IF(ISNUMBER(C17),TEXT(C17,"#,##0.00"),"–")&" "&{lk("Unit")}', 2, NAVY)
    tile(ws, 6, 12, "IMPROVEMENT ACHIEVED", '=IF(OR(NOT(ISNUMBER(J7)),NOT(ISNUMBER(C17)),N(C15)=N(C17)),"–",(C15-J7)/(C15-C17))', PCT,
         '="100% = target reached"', 2, "2E7D32")
    tile(ws, 10, 8, "SAVINGS TO DATE", f"=SUM($I${WK0}:$I${WK0 + NW - 1})", "$#,##0", '="(baseline − actual) × volume × $"', 2, "2E7D32")
    tile(ws, 10, 10, "ANNUAL RUN-RATE", '=IF(OR(NOT(ISNUMBER(J7)),NOT(ISNUMBER(C18)),NOT(ISNUMBER(C19))),"–",(C15-J7)*C18*C19*52)', "$#,##0",
         '="estimate – Finance to confirm"', 2, GRAY)
    tile(ws, 10, 12, "STATUS",
         f'=IF(C20="No","Inactive",IF(COUNT({F_})=0,"Not Started",IF(AND(R9<=13,R9-R6>=2),"Data Missing",IF(N(L7)>=R7-set_Green,"On Track",IF(N(L7)>=R7-set_Yellow,"Behind","Off Track")))))',
         None, '=R8&" of 9 actions done"', 2, BURG)
    text_cf(ws, "L11", STATUS)
    ws.row_dimensions[7].height = 26
    ws.row_dimensions[11].height = 26
    # weekly table
    section(ws, WK0 - 2, 2, "Weekly tracking – 13 weeks (3 months)", 11)
    hdr_row(ws, WK0 - 1, 2, ["Week", "Week Starting", "Phase", "Target Path", "Actual KPI", "vs Baseline", "On Target?",
                             "Weekly Savings", "Cumulative Savings", "What we did this week", "Evidence link (Teams)"])
    for c in (6, 11, 12):
        ws.cell(WK0 - 1, c).fill = fill(NAVY)
    for c in (3, 4, 5, 7, 8, 9, 10):
        ws.cell(WK0 - 1, c).fill = fill(CALCHDR)
    ws.row_dimensions[WK0 - 1].height = 30
    demo_actual = []
    if demo:
        start = dt.date(2026, 8, 24) + dt.timedelta(days=7 * (i - 1))
        weeks_done = min(13, (AS_OF - start).days // 7 + 1)
        tgt = base_demo * (1 - proj[10])
        speed = [1.0, 0.8, 1.1, 0.55][i - 1]
        for w in range(1, 14):
            if w > weeks_done or (g == "PK" and i == 3 and w > weeks_done - 2):
                demo_actual.append(None)
            else:
                path = base_demo - (base_demo - tgt) * min(1, max(0, (w - 1) / 7)) * speed
                demo_actual.append(round(path * random.uniform(0.95, 1.07), 2))
    for k in range(NW):
        r = WK0 + k
        cells = [
            (2, k + 1, "0"),
            (3, f'=IF(ISNUMBER($C$14),$C$14+7*(B{r}-1),"")', DATE),
            (4, f'=IF(B{r}<=4,"M1 – Measure & standardize",IF(B{r}<=8,"M2 – Implement","M3 – Sustain & verify"))', None),
            (5, f'=IF(AND(ISNUMBER($C$15),ISNUMBER($C$17)),$C$15-($C$15-$C$17)*MIN(1,MAX(0,(B{r}-1)/7)),"")', "#,##0.00"),
            (6, demo_actual[k] if demo else None, "#,##0.00"),
            (7, f'=IF(AND(ISNUMBER(F{r}),N($C$15)<>0),F{r}/$C$15-1,"")', "+0%;-0%;0%"),
            (8, f'=IF(OR(NOT(ISNUMBER(F{r})),NOT(ISNUMBER(E{r}))),"",IF(F{r}<=E{r}*(1+set_Tol),"✓ On target","✗ Off target"))', None),
            (9, f'=IF(AND(ISNUMBER(F{r}),ISNUMBER($C$15),ISNUMBER($C$18),ISNUMBER($C$19)),($C$15-F{r})*$C$18*$C$19,"")', "$#,##0;[Red]-$#,##0"),
            (10, f'=IF(ISNUMBER(I{r}),SUM($I${WK0}:I{r}),"")', "$#,##0;[Red]-$#,##0"),
            (11, "DEMO: actions this week" if demo and demo_actual[k] is not None else None, None),
            (12, f"https://example.sharepoint.com/sites/DunGPS/Shared/Supervisor/DEMO_{g}{i}_wk{k + 1}.jpg"
             if demo and demo_actual[k] is not None and k % 2 == 0 else None, None),
        ]
        for c, v, fmt in cells:
            cell = ws.cell(r, c, v)
            cell.border = BORDER
            cell.font = f(9)
            cell.alignment = Alignment(horizontal="center" if c < 11 else "left", vertical="center")
            if c not in (6, 11, 12):
                cell.fill = fill(LGRAY)
            if fmt:
                cell.number_format = fmt
        ws.cell(r, 15, f'=IF(ISNUMBER($C$15),$C$15,"")')                 # baseline series
        ws.cell(r, 16, f'=IF(ISNUMBER(F{r}),COUNT(F{r}:$F${WK0 + NW - 1}),"")')  # rank from end
        ws.cell(r, 17, f'=IF(ISNUMBER(F{r}),B{r},"")')                    # week with data
    for c in ("O", "P", "Q"):
        ws.column_dimensions[c].hidden = True
    rng = f"H{WK0}:H{WK0 + NW - 1}"
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT(H{WK0},1)="✓"'], font=Font(color="2E7D32", bold=True)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT(H{WK0},1)="✗"'], font=Font(color="C62828", bold=True)))
    for k, col in [(0, "FBE5D6"), (4, "DEEBF7"), (8, "E2F0D9")]:
        ws.conditional_formatting.add(f"D{WK0 + k}:D{WK0 + min(k + 3, 12) + (1 if k == 8 else 0)}",
                                      FormulaRule(formula=["TRUE"], fill=fill(col)))
    # dv for actual
    d = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True)
    ws.add_data_validation(d)
    d.add(f"F{WK0}:F{WK0 + NW - 1}")
    # chart
    lc = LineChart()
    lc.add_data(Reference(ws, min_col=15, min_row=WK0, max_row=WK0 + NW - 1), titles_from_data=False)
    lc.add_data(Reference(ws, min_col=5, min_row=WK0, max_row=WK0 + NW - 1), titles_from_data=False)
    lc.add_data(Reference(ws, min_col=6, min_row=WK0, max_row=WK0 + NW - 1), titles_from_data=False)
    from openpyxl.chart.series import SeriesLabel
    for s, nm, colr, dash in [(lc.series[0], "Baseline", "A6A6A6", "dash"), (lc.series[1], "Target path", "2E7D32", "sysDash"),
                              (lc.series[2], "Actual", GCOL[g], None)]:
        s.tx = SeriesLabel(v=nm)
        color_series(s, colr, True)
        if dash:
            s.graphicalProperties.line.dashStyle = dash
    lc.series[2].marker.symbol = "circle"
    lc.series[2].marker.size = 6
    lc.series[2].graphicalProperties.line.width = 38000
    lc.set_categories(Reference(ws, min_col=2, min_row=WK0, max_row=WK0 + NW - 1))
    lc.y_axis.scaling.min = 0
    style_chart(lc, "Improvement over 13 weeks (lower is better)", ytitle=None, xtitle="Week", w=24, h=7.5)
    ws.add_chart(lc, f"B{WK0 + NW + 1}")
    # action plan
    section(ws, AP0 - 2, 2, "3-month action plan – paste the evidence link, then mark Done", 11)
    hdr_row(ws, AP0 - 1, 2, ["Month", "Action", "", "", "", "Owner", "", "Due", "Done?", "Evidence link (Teams)", "Date Done"])
    for c in (3, 7, 9):
        ws.cell(AP0 - 1, c).fill = fill(CALCHDR if c in (3, 9) else NAVY)
    ws.merge_cells(start_row=AP0 - 1, start_column=3, end_row=AP0 - 1, end_column=6)
    ws.merge_cells(start_row=AP0 - 1, start_column=7, end_row=AP0 - 1, end_column=8)
    for k in range(9):
        r = AP0 + k
        m = k // 3 + 1
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
        ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=8)
        done = demo and (k < 3 or (k < 6 and i != 4) or (k == 6 and i == 1))
        vals = [(2, f"M{m}"), (3, "=" + lk(f"A{k + 1}")), (7, f"DEMO Supervisor {g}{i}" if demo else None),
                (9, f'=IF(ISNUMBER($C$14),$C$14+{28 * m - 1},"")'), (10, "Yes" if done else "No"),
                (11, f"https://example.sharepoint.com/sites/DunGPS/Shared/Supervisor/DEMO_{g}{i}_A{k + 1}.pdf" if done else None),
                (12, dt.date(2026, 9, 1) + dt.timedelta(days=9 * k) if done else None)]
        for c, v in vals:
            cell = ws.cell(r, c, v)
            cell.font = f(9)
            cell.alignment = Alignment(wrap_text=c != 11, vertical="center", horizontal="left" if c in (3, 7, 11) else "center")
            if c in (2, 3, 9):
                cell.fill = fill(LGRAY)
        for c in range(2, 13):
            ws.cell(r, c).border = BORDER
        ws.cell(r, 9).number_format = DATE
        ws.cell(r, 12).number_format = DATE
        ws.row_dimensions[r].height = 28
    d = DataValidation(type="list", formula1=f"=IF(LEN($K{AP0})>0,l9_YesNo,l9_No)", allow_blank=True, showErrorMessage=True,
                       errorTitle="Evidence required", error="Paste the evidence link (Teams) in 'Evidence link' first, then mark Done = Yes.")
    ws.add_data_validation(d)
    d.add(f"J{AP0}:J{AP0 + 8}")
    d = DataValidation(type="list", formula1="=l9_People", allow_blank=True)
    ws.add_data_validation(d)
    d.add(f"G{AP0}:G{AP0 + 8}")
    d = DataValidation(type="date", operator="between", formula1="DATE(2020,1,1)", formula2="DATE(2040,12,31)", allow_blank=True)
    ws.add_data_validation(d)
    d.add(f"L{AP0}:L{AP0 + 8}")
    text_cf(ws, f"J{AP0}:J{AP0 + 8}", {"Yes": "Green"})
    ws.conditional_formatting.add(f"I{AP0}:I{AP0 + 8}", FormulaRule(formula=[f'AND(ISNUMBER(I{AP0}),I{AP0}<AsOf90,J{AP0}<>"Yes")'],
                                                                     font=Font(color="C62828", bold=True)))
    ws.cell(AP0 + 10, 2, "Evidence: " ).font = f(8, True, NAVY)
    ws.cell(AP0 + 10, 3, "=" + lk("Evidence")).font = f(8, False, GRAY, True)
    ws.freeze_panes = "A5"
    landscape(ws, f"A1:M{AP0 + 10}", one_page=False)
    ws.page_setup.fitToHeight = 2
    ws.sheet_view.zoomScale = 90


def build_dash(wb, sh, demo):
    ws = sh[S_D]
    LS = qs(sh[S_L])
    title_band(ws, "Supervisor 90-Day Projects — Dashboard", "", 15, legend=False)
    ws["B3"] = '="Status as of "&TEXT(AsOf90,"dd-mmm-yyyy")&"  ·  Nugget Line · Mixers & Rolled Dough · Packaging  ·  savings = (baseline − actual) × volume × $"' + (
        '&"  ·  DEMONSTRATION DATA"' if demo else "")
    widths = [2, 8, 32, 18, 10, 8, 10, 10, 10, 8, 13, 12, 12, 9, 13]
    for k, w in enumerate(widths, 1):
        ws.column_dimensions[L(k)].width = w
    cols = ["Slot", "Project", "Supervisor", "Start", "Week", "Baseline", "Latest (4-wk)", "Target", "Unit",
            "Improvement Achieved", "Savings to Date", "Annual Run-Rate", "Actions", "Status"]
    # tiles
    rows_all = []
    r0 = 10
    r = r0
    group_rows = {}
    for g in ["NL", "MR", "PK"]:
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=15)
        hc = ws.cell(r, 2, GROUPS[g].upper())
        hc.font = f(10, True, WHITE)
        hc.fill = fill(GCOL[g])
        hc.alignment = Alignment(indent=1)
        hdr_row(ws, r + 1, 2, cols)
        ws.row_dimensions[r + 1].height = 30
        rr = []
        for i in range(1, 5):
            rw = r + 1 + i
            S = qs(sh[f"{g}-{i}"])
            vals = [
                (2, f'=HYPERLINK("#\'{g}-{i}\'!A1","{g}-{i}")', None, "center"),
                (3, f"={S}!$C$7&\"\"", None, "left"),
                (4, f'=IF({S}!$C$11="","(not set)",{S}!$C$11&"")', None, "left"),
                (5, f'=IF(ISNUMBER({S}!$C$14),{S}!$C$14,"")', DATE, "center"),
                (6, f'=IF({S}!$R$9=0,"",{S}!$R$9)', "0", "center"),
                (7, f'=IF(ISNUMBER({S}!$C$15),{S}!$C$15,"")', "#,##0.00", "center"),
                (8, f'=IF(ISNUMBER({S}!$J$7),{S}!$J$7,"")', "#,##0.00", "center"),
                (9, f'=IF(ISNUMBER({S}!$C$17),{S}!$C$17,"")', "#,##0.00", "center"),
                (10, f'=IFERROR(INDEX(MN_Unit,MATCH({S}!$C$6,MN_Code,0))&"","")', None, "center"),
                (11, f'=IF(ISNUMBER({S}!$L$7),{S}!$L$7,"")', PCT, "center"),
                (12, f"=N({S}!$H$11)", "$#,##0;[Red]-$#,##0", "center"),
                (13, f'=IF(ISNUMBER({S}!$J$11),{S}!$J$11,"")', "$#,##0;[Red]-$#,##0", "center"),
                (14, f'={S}!$R$8&"/9"', None, "center"),
                (15, f"={S}!$L$11&\"\"", None, "center"),
            ]
            for c, v, fmt, al in vals:
                calc_cell(ws, rw, c, v, fmt, align=al, size=8)
            ws.cell(rw, 2).font = Font(name=FONT, size=9, bold=True, color=NAVY, underline="single")
            rr.append(rw)
        tr = r + 6
        calc_cell(ws, tr, 2, "TOTAL", None, True, size=8)
        calc_cell(ws, tr, 3, f'=COUNTIF(O{rr[0]}:O{rr[-1]},"On Track")&" of "&(4-COUNTIF(O{rr[0]}:O{rr[-1]},"Inactive"))&" projects on track"', None, True, align="left", size=8)
        calc_cell(ws, tr, 11, f'=IFERROR(AVERAGE(K{rr[0]}:K{rr[-1]}),"")', PCT, True, size=8)
        calc_cell(ws, tr, 12, f"=SUM(L{rr[0]}:L{rr[-1]})", "$#,##0;[Red]-$#,##0", True, size=8)
        calc_cell(ws, tr, 13, f"=SUM(M{rr[0]}:M{rr[-1]})", "$#,##0;[Red]-$#,##0", True, size=8)
        for c in range(2, 16):
            ws.cell(tr, c).fill = fill(LGRAY)
            ws.cell(tr, c).border = BORDER
        bar_cf(ws, f"K{rr[0]}:K{rr[-1]}", "2E7D32")
        text_cf(ws, f"O{rr[0]}:O{rr[-1]}", STATUS)
        group_rows[g] = (rr, tr)
        rows_all += rr
        r = tr + 2
    last = r
    allK = ",".join(f"K{x}" for x in rows_all)
    allL = "+".join(f"L{group_rows[g][1]}" for g in group_rows)
    allM = "+".join(f"M{group_rows[g][1]}" for g in group_rows)
    stat = lambda s: "+".join(f'COUNTIF(O{group_rows[g][0][0]}:O{group_rows[g][0][-1]},"{s}")' for g in group_rows)
    acts = "+".join(f"N('{g}-{i}'!$R$8)" for g, i in SLOTS)
    tiles = [("ACTIVE PROJECTS", f"=12-({stat('Inactive')})", "0", '="across 3 groups"', NAVY, 2),
             ("ON TRACK", f"={stat('On Track')}", "0", f'=({stat("Behind")})&" behind · "&({stat("Off Track")})&" off track"', "2E7D32", 2),
             ("AVG IMPROVEMENT ACHIEVED", f'=IFERROR(AVERAGE({allK}),0)', PCT, '="vs. target (100% = reached)"', "2E7D32", 3),
             ("SAVINGS TO DATE", f"={allL}", "$#,##0", '="sum of weekly savings"', "2E7D32", 2),
             ("ANNUAL RUN-RATE", f"={allM}", "$#,##0", '="estimate – Finance to confirm"', GRAY, 2),
             ("ACTIONS DONE", f'=({acts})&" / 108"', None, '="with evidence"', NAVY, 3)]
    c = 2
    for lab, fm, fmt, sub, acc, w in tiles:
        tile(ws, 5, c, lab, fm, fmt, sub, w, acc)
        c += w
    ws.row_dimensions[6].height = 28
    # charts
    section(ws, last, 2, "Improvement achieved & savings", 14)
    L_ = sh[S_L]
    bc = BarChart()
    bc.type = "bar"
    bc.add_data(Reference(L_, min_col=17, min_row=12, max_row=23), titles_from_data=False)
    bc.set_categories(Reference(L_, min_col=16, min_row=12, max_row=23))
    color_series(bc.series[0], "2E7D32")
    bc.x_axis.scaling.orientation = "maxMin"
    bc.y_axis.number_format = "0%"
    bc.y_axis.scaling.min = 0
    bc.legend = None
    style_chart(bc, "Improvement achieved by project (100% = target)", w=16, h=9)
    bc.legend = None
    ws.add_chart(bc, f"B{last + 1}")
    sc = BarChart()
    sc.type = "col"
    sc.add_data(Reference(L_, min_col=18, min_row=26, max_row=28), titles_from_data=False)
    sc.set_categories(Reference(L_, min_col=16, min_row=26, max_row=28))
    color_series(sc.series[0], NAVY)
    sc.y_axis.number_format = "$#,##0"
    style_chart(sc, "Savings to date by group", w=11, h=9)
    sc.legend = None
    ws.add_chart(sc, f"J{last + 1}")
    landscape(ws, f"A1:O{last + 20}", one_page=True)
    ws.freeze_panes = "A4"
    protect(ws)


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "."
    build(False, f"{out}/CHG_Supervisor_90-Day_Projects.xlsx")
    build(True, f"{out}/CHG_Supervisor_90-Day_Projects_DEMO.xlsx")
    print("built")
