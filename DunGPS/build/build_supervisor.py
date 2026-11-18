"""Builds CHG_Supervisor_Productivity_Project_Tracker.xlsx (production) and a DEMO version.

Supervisors pick improvement projects (5S, standard work, SMED, problem solving, ...) from a
CI project library; progress is tracked through 5 standard milestones with coach verification.
Usage: python build_supervisor.py <out_dir>
"""
import sys
import random
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from xlcommon import *  # noqa

S_DASH = "Program Dashboard"
S_ASG = "Project Assignments"
S_LIB = "Project Library"
S_GAP = "Gap & Barrier Log"
S_5S = "5S Audit Log"
S_ROS = "Supervisor Roster"
S_CFG = "Lists & Settings"
S_INS = "Instructions"
S_CALC = "Calc"
N_ASG, N_GAP, N_5S, N_ROS, N_LIB = 300, 300, 400, 100, 60
NSUP = 24  # scorecard capacity

# (code, name, category, pillar, bronze ref, level, weeks, KPI, unit, method, M1..M5, evidence)
LIB = [
    ("5S-01", "5S Workplace Organization – Production Area", "Workplace Organization", "AM", "", "Starter", 8, "5S audit score", "%",
     "5S audit, red-tag, shadow boards, visual controls",
     "Sort: baseline 5S audit; red-tag and remove unneeded items", "Set in order: locations, labels, shadow boards, floor marking",
     "Shine: deep clean; cleaning standard & schedule", "Standardize: visual standards, OPL, 5S checklist",
     "Sustain: 4 weekly audits at or above target score", "Before/after photos, audit sheets, red-tag log, OPL"),
    ("5S-02", "5S – Lab & Sanitation Areas", "Workplace Organization", "FSQ", "B-FSQ-08", "Starter", 8, "5S audit score", "%",
     "5S audit, red-tag, visual controls",
     "Sort: baseline audit; red-tag", "Set in order: labelled storage, chemical / tool locations",
     "Shine: cleaning standard", "Standardize: visual standards & checklist", "Sustain: 4 weekly audits at target",
     "Before/after photos, audit sheets"),
    ("PS-01", "Process Improvement – 5 Why & Fishbone (A3)", "Problem Solving", "VO", "B-VO-01", "Starter", 6, "Problem KPI (defined in A3)", "",
     "A3, 5 Why, fishbone, go-see, PDCA",
     "Define problem statement & baseline data", "Root cause: fishbone + 5 Why verified at the gemba",
     "Countermeasures implemented (who/what/when)", "Results verified with ≥4 weeks of data",
     "Standardize (SOP/OPL), train, share learning", "Completed A3, data before/after, updated SOP/OPL"),
    ("SW-01", "Standard Work Instruction (1 per Supervisor)", "Standard Work", "E&T", "B-E&T-01", "Starter", 6, "SWI adherence", "%",
     "Time observation, key points & reasons, TWI job instruction",
     "Select task; observe ≥10 cycles across shifts", "Draft SWI: steps, key points, reasons, safety & quality",
     "Trial & validate with operators; adjust", "Approve & train all shifts (sign-off)",
     "Adherence audits ≥90% for 4 weeks", "Approved SWI, training sign-off, adherence audits"),
    ("CIL-01", "CIL (Clean-Inspect-Lubricate) Standard for an Asset", "Autonomous Maintenance", "AM", "B-AM-01", "Intermediate", 8, "CIL completion", "%",
     "AM step 1-3, abnormality tags, CIL sheet",
     "Select asset; deep clean & tag abnormalities", "Restore basic conditions; close tags",
     "Draft CIL (points, method, tool, time, frequency)", "Train operators; sign-off; post visual CIL",
     "CIL completion ≥90% for 4 weeks", "CIL sheet, tag log, training sign-off, compliance data"),
    ("CL-01", "Centerline Standard Deployment", "Autonomous Maintenance", "AM", "B-AM-01", "Intermediate", 8, "Centerline compliance", "%",
     "Critical parameter analysis, centerline sheet, visual marks",
     "Identify critical process / machine settings", "Set target & ranges using data",
     "Visual centerlines on equipment + check sheet", "Train operators & leads on all shifts",
     "Compliance audits ≥90%; deviations handled", "Centerline sheet, photos, check sheets"),
    ("DH-01", "Defect Identification & Handling (Tag System)", "Autonomous Maintenance", "AM", "B-AM-01", "Starter", 6, "Tag closure rate", "%",
     "Red/white tags, tag board, tag log",
     "Set up tag board & log", "Train operators to identify & tag defects", "Weekly tag review with maintenance",
     "Tag closure process & ownership defined", "Closure rate at target for 4 weeks", "Tag log, board photo, closure trend"),
    ("DC-01", "Deep Clean & Restore Selected Asset", "Autonomous Maintenance", "AM", "B-AM-02", "Starter", 4, "Abnormalities found / closed", "#",
     "AM step 1 deep clean, abnormality list",
     "Select asset; safety & LOTO plan", "Deep clean with operators", "Abnormality list & tags",
     "Restore / repair with maintenance", "Before/after review; cleaning standard", "Before/after photos, abnormality list"),
    ("LOTO-01", "LOTO & Safety Map for Equipment", "Safety", "AM", "B-AM-03", "Intermediate", 8, "Assets with approved LOTO & map", "#",
     "Energy source identification, LOTO procedure, safety map",
     "List assets & energy sources", "Draft machine-specific LOTO procedures", "Create safety maps (hazards, PPE, energy points)",
     "EHS approval; post at equipment; train", "LOTO audit passed", "Approved LOTO, posted maps, audit"),
    ("SMED-01", "SMED – Changeover Reduction", "Productivity", "VO", "B-VO-06", "Advanced", 10, "Changeover time", "min",
     "SMED (video, internal/external, parallel work, quick-release)",
     "Video & baseline changeover (3+ events)", "Separate internal vs external elements",
     "Convert / streamline elements; tooling & kits", "Trial new method; verify time & quality",
     "Standard changeover + training; sustain 4 weeks", "Baseline vs new time study, standard"),
    ("OEE-01", "Top-3 Downtime Loss Reduction (Redzone)", "Productivity", "VO", "B-VO-04", "Intermediate", 8, "Downtime", "min/shift",
     "Redzone loss reports, Pareto, 5 Why",
     "Pull Redzone loss data; Pareto top-3 losses", "Root cause on #1 loss", "Countermeasures implemented",
     "Verify reduction with ≥4 weeks data", "Standardize & move to next loss", "Redzone reports before/after, A3"),
    ("FIT-01", "Tiger Team / FIT – Loss Elimination", "Productivity", "VO", "B-VO-07", "Advanced", 12, "Target loss", "",
     "Charter, cross-functional team, DMAIC",
     "Charter approved (problem, scope, team, baseline)", "Measure & analyze losses", "Implement improvements",
     "Verify results & savings (Finance)", "Control plan & handover", "Charter, analysis, results, control plan"),
    ("PFM-01", "Process Flow Map – 1 Line", "Process Definition", "VO", "B-VO-02", "Starter", 4, "Wastes identified", "#",
     "Process mapping, value-add analysis",
     "Define line & boundaries", "Walk the process; map steps & flows", "Identify wastes / delays",
     "Validate map with team", "Approved map + improvement list", "Process flow map"),
    ("SPIOC-01", "SPIOC – Process Definition", "Process Definition", "VO", "B-VO-05", "Starter", 3, "SPIOC completed", "#",
     "Supplier-Process-Input-Output-Customer",
     "Select process", "Identify suppliers & inputs", "Define process steps & outputs", "Identify customers & requirements",
     "Review & approve", "Completed SPIOC"),
    ("YLD-01", "Scrap / Yield Improvement", "Cost", "VO", "", "Intermediate", 8, "Scrap", "%",
     "Data collection, Pareto, 5 Why, SPC",
     "Baseline scrap by type", "Pareto & root cause", "Countermeasures", "Verify ≥4 weeks", "Standardize",
     "Scrap data before/after, A3"),
    ("GA-01", "Giveaway / Overfill Reduction", "Cost", "VO", "", "Intermediate", 8, "Giveaway", "%",
     "Weight data, capability (Cpk), centerlining",
     "Baseline weight data & giveaway", "Capability study; root cause", "Adjust targets / controls",
     "Verify giveaway reduction (compliance maintained)", "Standardize checks", "Weight data, Cpk, standard"),
    ("LAB-01", "Labor Productivity – Line Balancing / Crewing", "Productivity", "VO", "", "Advanced", 10, "Lbs per labor hour", "lb/hr",
     "Time study, line balance chart, crewing standard",
     "Baseline output & crewing", "Time study & line balance", "Re-balance / crewing standard trial",
     "Verify productivity & safety/ergonomics", "Approve crewing standard", "Time study, balance chart, standard"),
    ("SAN-01", "Sanitation Time Reduction", "Food Safety & Quality", "FSQ", "B-FSQ-05", "Advanced", 10, "Sanitation hours", "hr",
     "SMED for sanitation – food-safety verification must be maintained",
     "Baseline sanitation time by step", "Separate prep vs critical steps", "Improve tools, kits, sequence",
     "Verify time AND sanitation verification (ATP/swab)", "Standardize SSOP update", "Time study, verification records, SSOP"),
    ("SKL-01", "Operator Skill Matrix for My Area", "People Development", "E&T", "B-E&T-02", "Starter", 4, "Cross-training coverage", "%",
     "Skill matrix, training plan",
     "List tasks / skills for area", "Assess each operator", "Identify gaps & training plan", "Deliver training",
     "Coverage target met; matrix posted", "Skill matrix, training plan"),
    ("LSW-01", "Leader Standard Work (Supervisor)", "Leadership", "E&T", "B-E&T-03", "Starter", 4, "LSW adherence", "%",
     "Leader standard work, gemba walk, tier meetings",
     "Document daily/weekly routine", "Draft LSW", "Pilot 2 weeks", "Manager review & approve",
     "Adherence ≥90% for 4 weeks", "LSW document, adherence checks"),
    ("FMDS-01", "FMDS / Tier Board Refresh for My Area", "Leadership", "LEAD", "B-LEAD-01", "Starter", 4, "Tier meeting adherence", "%",
     "FMDS board, KPIs, action log",
     "Review current board & KPIs", "Redesign board to standard", "Launch tier meeting cadence", "Coach team leads",
     "Audit adherence 4 weeks", "Board photo, attendance, action log"),
    ("SAFE-01", "Hazard Elimination / SafeStart Observations", "Safety", "EHS", "B-EHS-01", "Starter", 6, "Hazards closed", "#",
     "SafeStart, hazard hunt, hierarchy of controls",
     "Hazard hunt with team", "Prioritize risks", "Implement controls", "Verify effectiveness with EHS",
     "Standardize & share", "Hazard list, before/after"),
    ("GMP-01", "GMP / Foreign Material Risk Reduction", "Food Safety & Quality", "FSQ", "B-FSQ-06", "Intermediate", 6, "GMP audit findings", "#",
     "GMP audit, FM risk assessment",
     "Baseline GMP / FM audit", "Risk assessment", "Implement controls", "Re-audit & verify", "Standardize checks",
     "Audit results, risk assessment"),
    ("IDEA-01", "Idea / Kaizen Pipeline for My Team", "Engagement", "IM", "B-IM-01", "Starter", 4, "Ideas implemented", "#",
     "Idea board, quick kaizen",
     "Launch idea board with team", "Collect & screen ideas", "Implement quick wins", "Submit larger ideas to Steering Committee",
     "Recognize & sustain cadence", "Idea log, implemented list"),
]
CATS = ["Workplace Organization", "Problem Solving", "Standard Work", "Autonomous Maintenance", "Safety",
        "Productivity", "Process Definition", "Cost", "Food Safety & Quality", "People Development", "Leadership",
        "Engagement"]
LISTS = [
    ("Project Status", "Status", ["Not Started", "In Progress", "On Hold", "Complete", "Cancelled"], True),
    ("Milestone Status", "MS", ["Not Started", "In Progress", "Done"], True),
    ("Coach Verification", "Verif", ["Not Submitted", "Submitted - Pending Review", "Verified", "Rejected - Rework"], True),
    ("Yes / No", "YesNo", ["Yes", "No"], True),
    ("Active / Inactive", "Active", ["Active", "Inactive"], True),
    ("Gap Status", "GStatus", ["Open", "In Progress", "Closed", "Cancelled"], True),
    ("Severity", "Sev", ["High", "Medium", "Low"], True),
    ("Shift", "Shift", ["A", "B", "C", "All"], False),
    ("Production Area", "Area", ["Pre-weigh", "Mixing", "Dough Processing / Forming", "Gemini Equipment", "Nugget Line",
                                 "Rolled Dough Lines", "Baking", "Cooling / Freezing", "Packaging", "Sanitation",
                                 "Maintenance / Parts Room", "Laboratory / QA", "Warehouse / Shipping"], False),
    ("Project Category", "Cat", CATS, False),
    ("Difficulty", "Level", ["Starter", "Intermediate", "Advanced"], False),
    ("Program Role", "Role", ["Supervisor", "Team Leader", "Area Manager", "Coach / CI Engineer", "Plant Leadership", "Finance"], False),
    ("Barrier Category", "BarCat", ["Time / staffing", "Training / skills", "Maintenance support", "Engineering support",
                                    "Data / system (Redzone, CMMS)", "Materials / supplier", "Equipment condition",
                                    "Management support / decision", "Communication", "Other"], False),
    ("Support Needed From", "Support", ["Maintenance", "Engineering", "Quality / FSQ", "EHS", "HR / Training",
                                        "Plant Leadership", "Supply Chain", "IT / Systems", "Finance"], False),
    ("Pillar", "Pillar", ["AM", "PM", "FSQ", "EHS", "LEAD", "E&T", "VO", "IM", "SC"], False),
]
R0, RN = 6, 55


def build(demo, path):
    random.seed(7)
    AS_OF = dt.date(2026, 11, 18)
    wb = Workbook()
    wb.remove(wb.active)
    sh = {n: wb.create_sheet(n) for n in [S_DASH, S_ASG, S_LIB, S_GAP, S_5S, S_ROS, S_CFG, S_INS, S_CALC]}
    for n, c in {S_DASH: NAVY, S_ASG: BURG, S_LIB: BURG, S_GAP: BURG, S_5S: BURG, S_ROS: "7F7F7F", S_CFG: "7F7F7F",
                 S_INS: "7F7F7F", S_CALC: "BFBFBF"}.items():
        sh[n].sheet_properties.tabColor = c
    tag = ("DEMONSTRATION DATA – fictitious records for testing only." if demo
           else "C.H. Guenther & Son — Duncanville · Supervisor Productivity Program")

    # ---------- settings ----------
    ws = sh[S_CFG]
    title_band(ws, "Lists & Settings", "Program parameters and dropdown lists. Gray-header lists are system values "
               "used by formulas – add at the end, do not rename.", 30)
    section(ws, 5, 2, "Program parameters", 3)
    params = [
        ("sp_AsOfOverride", "Status date override", AS_OF if demo else None, DATE, "Blank = TODAY()."),
        ("sp_AsOf", "ƒ Effective status date", "=IF(ISNUMBER(sp_AsOfOverride),sp_AsOfOverride,TODAY())", DATE, ""),
        ("sp_PeriodStart", "Program period start", None, DATE, "Blank = 1-Jan of the status-date year."),
        ("sp_Start", "ƒ Effective period start", "=IF(ISNUMBER(sp_PeriodStart),sp_PeriodStart,DATE(YEAR(sp_AsOf),1,1))", DATE, ""),
        ("sp_Target", "Verified projects required per supervisor (period)", 2, "0",
         "Program expectation per participating supervisor. Confirm with plant leadership."),
        ("sp_W1", "Milestone weight M1", 0.15, PCT, "Weights must total 100%."),
        ("sp_W2", "Milestone weight M2", 0.20, PCT, ""),
        ("sp_W3", "Milestone weight M3", 0.30, PCT, ""),
        ("sp_W4", "Milestone weight M4", 0.20, PCT, ""),
        ("sp_W5", "Milestone weight M5", 0.15, PCT, ""),
        ("sp_WSum", "ƒ Weight check", "=IF(ABS(sp_W1+sp_W2+sp_W3+sp_W4+sp_W5-1)<0.0001,\"OK – 100%\",\"ERROR – weights must total 100%\")", None, ""),
        ("sp_IPCredit", "Credit for a milestone 'In Progress'", 0.5, PCT, "Share of the milestone weight earned while in progress."),
        ("sp_Cap", "Max credit before coach verification", 0.9, PCT, "Unverified projects cannot exceed this."),
        ("sp_GreenTol", "Green tolerance (pts behind plan)", 0.10, PCT, ""),
        ("sp_YellowTol", "Yellow tolerance (pts behind plan)", 0.25, PCT, ""),
        ("sp_DueSoon", "Due-soon window (days)", 14, "0", ""),
        ("sp_5STarget", "5S audit target score", 0.80, PCT, "Score = total points ÷ 25."),
    ]
    r = 6
    for nm, lab, val, fmt, note in params:
        ws.cell(r, 2, lab).font = f(9, lab.startswith("ƒ"))
        c = ws.cell(r, 3, val)
        c.border = BORDER
        c.alignment = Alignment(horizontal="center")
        c.font = f(9, False, "000000" if lab.startswith("ƒ") else "0000FF")
        c.fill = fill(LGRAY if lab.startswith("ƒ") else WHITE)
        if fmt:
            c.number_format = fmt
        ws.cell(r, 4, note).font = f(8, False, GRAY, True)
        define(wb, nm, f"{qs(ws)}!$C${r}")
        r += 1
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 44
    col = 6
    lc = {}
    for hdr, key, vals, system in LISTS:
        hdr_row(ws, 5, col, [hdr], calc=system)
        ws.column_dimensions[L(col)].width = max(14, min(30, max(len(v) for v in vals) + 2))
        for i in range(RN - R0 + 1):
            c = ws.cell(R0 + i, col, vals[i] if i < len(vals) else None)
            c.font = f(9)
            c.border = BORDER
        dyn_list(wb, f"sl_{key}", ws, L(col), R0, RN)
        lc[key] = L(col)
        col += 1

    # ---------- roster ----------
    ws = sh[S_ROS]
    title_band(ws, "Supervisor Roster", "Add people here once – all dropdowns update automatically. "
               "'In Program = Yes' makes a supervisor appear in the scorecard and the Supervisor dropdown.", 16)
    rs = [dict(key="Name", header="Name", width=24), dict(key="Role", header="Program Role", width=18, dv="sl_Role"),
          dict(key="Area", header="Primary Area", width=20, dv="sl_Area"), dict(key="Shift", header="Shift", width=7, dv="sl_Shift"),
          dict(key="Coach", header="Manager / Coach", width=22, dv="sl_People"),
          dict(key="InProg", header="In Program?", width=9, dv="sl_YesNo", default="Yes"),
          dict(key="Active", header="Active / Inactive", width=10, dv="sl_Active", default="Active"),
          dict(key="Email", header="Email", width=26)]
    tmp = Cols(rs)
    nL, aL, pL = tmp.letter["Name"], tmp.letter["Active"], tmp.letter["InProg"]
    rs += [dict(key="SeqAll", header="Seq (People)", width=7, calc=True, fmt="0",
                formula=f'=IF(AND({{Name}}<>"",{{Active}}="Active"),COUNTIFS(${nL}$6:{{Name}},"?*",${aL}$6:{{Active}},"Active"),"")'),
           dict(key="SeqSup", header="Seq (Supervisors)", width=7, calc=True, fmt="0",
                formula=f'=IF(AND({{Name}}<>"",{{Active}}="Active",{{InProg}}="Yes"),COUNTIFS(${nL}$6:{{Name}},"?*",${aL}$6:{{Active}},"Active",${pL}$6:{{InProg}},"Yes"),"")'),
           dict(key="Assigned", header="Projects (in scope)", width=8, calc=True, fmt="0",
                formula='=IF({Name}="","",COUNTIFS(A_Supervisor,{Name},A_InScope,1))'),
           dict(key="Verified", header="Verified Complete", width=8, calc=True, fmt="0",
                formula='=IF({Name}="","",COUNTIFS(A_Supervisor,{Name},A_Verified,1))')]
    rdata = []
    if demo:
        for i in range(1, 10):
            rdata.append(dict(Name=f"DEMO Supervisor {i:02d}", Role="Supervisor",
                              Area=["Nugget Line", "Rolled Dough Lines", "Baking", "Packaging", "Mixing", "Sanitation",
                                    "Pre-weigh", "Cooling / Freezing", "Gemini Equipment"][i - 1],
                              Shift="ABC"[(i - 1) % 3], Coach="DEMO Coach 01" if i < 6 else "DEMO Coach 02",
                              InProg="Yes", Active="Active", Email=f"demo.sup{i:02d}@example.com"))
        rdata += [dict(Name="DEMO Coach 01", Role="Coach / CI Engineer", Shift="All", InProg="No", Active="Active"),
                  dict(Name="DEMO Coach 02", Role="Area Manager", Shift="All", InProg="No", Active="Active"),
                  dict(Name="DEMO Controller", Role="Finance", Shift="All", InProg="No", Active="Active")]
    build_table(ws, wb, Cols(rs), 5, N_ROS, "tblRoster", rdata, name_prefix="R", freeze_col="C")

    # ---------- library ----------
    ws = sh[S_LIB]
    title_band(ws, "Project Library", "Menu of standard improvement projects supervisors choose from. CI-recommended "
               "and fully editable – not an official CHG standard. Add a row to offer a new project type.", 20)
    ls = [dict(key="Code", header="Project Code", width=10), dict(key="Name", header="Project", width=34, wrap=True),
          dict(key="Cat", header="Category", width=18, dv="sl_Cat"), dict(key="Pillar", header="DunGPS Pillar", width=8, dv="sl_Pillar"),
          dict(key="Bronze", header="Bronze Req Supported", width=11),
          dict(key="Level", header="Difficulty", width=11, dv="sl_Level"),
          dict(key="Weeks", header="Typical Duration (weeks)", width=9, fmt="0", dv="num"),
          dict(key="KPI", header="Primary KPI", width=18, wrap=True), dict(key="Unit", header="Unit", width=8),
          dict(key="Method", header="Method / Tools", width=26, wrap=True)]
    for j in range(1, 6):
        ls.append(dict(key=f"M{j}", header=f"Milestone {j} – definition of done", width=26, wrap=True))
    ls += [dict(key="Evidence", header="Required Evidence", width=28, wrap=True),
           dict(key="Active", header="Offered?", width=8, dv="sl_YesNo", default="Yes")]
    ldata = [dict(Code=x[0], Name=x[1], Cat=x[2], Pillar=x[3], Bronze=x[4] or None, Level=x[5], Weeks=x[6], KPI=x[7],
                  Unit=x[8] or None, Method=x[9], M1=x[10], M2=x[11], M3=x[12], M4=x[13], M5=x[14], Evidence=x[15],
                  Active="Yes") for x in LIB]
    lcols = Cols(ls)
    build_table(ws, wb, lcols, 5, N_LIB, "tblLibrary", ldata, name_prefix="LIB", freeze_col="D", row_height=48)
    dyn_list(wb, "sl_ProjCodes", ws, lcols.letter["Code"], 6, 5 + N_LIB)

    # ---------- assignments ----------
    ws = sh[S_ASG]
    title_band(ws, "Project Assignments", tag + "  ·  Pick a Project Code – name, category, pillar, KPI and milestone "
               "definitions fill in automatically. Verified = all 5 milestones Done + evidence + coach Verified.", 50)
    lk = lambda field: f'IFERROR(INDEX(LIB_{field},MATCH({{Code}},LIB_Code,0))&"","")'
    asp = [
        dict(key="ID", header="Assignment ID", width=10),
        dict(key="Supervisor", header="Supervisor", width=20, dv="sl_Supervisors"),
        dict(key="Area", header="Area", width=16, dv="sl_Area"),
        dict(key="Shift", header="Shift", width=6, dv="sl_Shift"),
        dict(key="Code", header="Project Code (select)", width=10, dv="sl_ProjCodes"),
        dict(key="PName", header="Project", width=30, calc=True, wrap=True, formula="=" + lk("Name")),
        dict(key="Cat", header="Category", width=16, calc=True, formula="=" + lk("Cat")),
        dict(key="Pillar", header="Pillar", width=7, calc=True, formula="=" + lk("Pillar")),
        dict(key="Scope", header="Specific Scope / Title", width=28, wrap=True),
        dict(key="Coach", header="Coach / Sponsor", width=18, dv="sl_People"),
        dict(key="Team", header="Team Members", width=18, wrap=True),
        dict(key="Start", header="Start Date", width=10, dv="date", fmt=DATE),
        dict(key="Weeks", header="Std Weeks", width=6, calc=True, fmt="0",
             formula='=IFERROR(INDEX(LIB_Weeks,MATCH({Code},LIB_Code,0)),"")'),
        dict(key="Target", header="Target Date", width=10, dv="date", fmt=DATE,
             note="Leave blank to use Start + library duration."),
        dict(key="EffTarget", header="Effective Target", width=10, calc=True, fmt=DATE,
             formula='=IF(ISNUMBER({Target}),{Target},IF(AND(ISNUMBER({Start}),ISNUMBER({Weeks})),{Start}+7*{Weeks},""))'),
        dict(key="Actual", header="Actual Completion", width=10, dv="date", fmt=DATE),
        dict(key="Status", header="Status", width=11, dv="sl_Status", default="Not Started"),
    ]
    for j in range(1, 6):
        asp.append(dict(key=f"M{j}", header=f"M{j}", width=9, dv="sl_MS", default="Not Started", align="center"))
    asp += [
        dict(key="NextMS", header="Current Milestone (definition of done)", width=30, calc=True, wrap=True,
             formula=('=IF({Code}="","",IF({M1}<>"Done","M1: "&' + lk("M1") + ',IF({M2}<>"Done","M2: "&' + lk("M2") +
                      ',IF({M3}<>"Done","M3: "&' + lk("M3") + ',IF({M4}<>"Done","M4: "&' + lk("M4") +
                      ',IF({M5}<>"Done","M5: "&' + lk("M5") + ',"All milestones done – submit for verification"))))))')),
        dict(key="KPI", header="KPI", width=14, calc=True, wrap=True, formula="=" + lk("KPI")),
        dict(key="Unit", header="Unit", width=6, calc=True, formula="=" + lk("Unit")),
        dict(key="Baseline", header="Baseline", width=8, fmt="#,##0.00"),
        dict(key="KTarget", header="KPI Target", width=8, fmt="#,##0.00"),
        dict(key="KActual", header="KPI Actual", width=8, fmt="#,##0.00"),
        dict(key="KProg", header="Progress to KPI Target", width=8, calc=True, fmt=PCT,
             formula='=IF(AND(ISNUMBER({Baseline}),ISNUMBER({KTarget}),ISNUMBER({KActual})),IF({KTarget}={Baseline},"",MAX(0,({KActual}-{Baseline})/({KTarget}-{Baseline}))),"")'),
        dict(key="EstSav", header="Est. Savings ($/yr)", width=10, fmt="$#,##0;($#,##0);-", dv="num"),
        dict(key="VerSav", header="Verified Savings ($/yr)", width=10, fmt="$#,##0;($#,##0);-", dv="num"),
        dict(key="FinOK", header="Finance Approved?", width=8, dv="sl_YesNo", default="No"),
        dict(key="Evid", header="Evidence Link", width=20),
        dict(key="Open", header="Open", width=7, calc=True, formula='=IF({Evid}="","",HYPERLINK({Evid},"Open ↗"))'),
        dict(key="Verif", header="Coach Verification", width=14, dv="sl_Verif", default="Not Submitted"),
        dict(key="VerBy", header="Verified By", width=16, dv="sl_People"),
        dict(key="VerDate", header="Verification Date", width=10, dv="date", fmt=DATE),
        dict(key="BarCat", header="Barrier Category", width=16, dv="sl_BarCat"),
        dict(key="Gap", header="Gap / Barrier", width=24, wrap=True),
        dict(key="Next", header="Next Action", width=24, wrap=True),
        dict(key="Updated", header="Last Updated", width=10, dv="date", fmt=DATE),
        # calc
        dict(key="InScope", header="In Scope", width=6, calc=True, fmt="0",
             formula='=IF(AND({Supervisor}<>"",{Code}<>"",{Status}<>"Cancelled"),1,0)'),
        dict(key="MSPct", header="Milestone %", width=8, calc=True, fmt=PCT,
             formula=('=IF({InScope}=0,"",' + "+".join(f'sp_W{j}*IF({{M{j}}}="Done",1,IF({{M{j}}}="In Progress",sp_IPCredit,0))' for j in range(1, 6)) + ')')),
        dict(key="Verified", header="Verified (1/0)", width=7, calc=True, fmt="0",
             formula='=IF(AND({InScope}=1,{Status}="Complete",{M1}="Done",{M2}="Done",{M3}="Done",{M4}="Done",{M5}="Done",{Verif}="Verified",{Evid}<>""),1,0)'),
        dict(key="Credit", header="Completion %", width=8, calc=True, fmt=PCT,
             formula='=IF({InScope}=0,"",IF({Verified}=1,1,MIN(sp_Cap,N({MSPct}))))'),
        dict(key="Planned", header="Planned % to Date", width=8, calc=True, fmt=PCT,
             formula='=IF({InScope}=0,"",IF(NOT(ISNUMBER({EffTarget})),0,IF(NOT(ISNUMBER({Start})),IF(sp_AsOf>={EffTarget},1,0),MAX(0,MIN(1,(sp_AsOf-{Start})/MAX(1,{EffTarget}-{Start}))))))'),
        dict(key="Overdue", header="Overdue (1/0)", width=7, calc=True, fmt="0",
             formula='=IF(AND({InScope}=1,{Verified}=0,ISNUMBER({EffTarget})),IF({EffTarget}<sp_AsOf,1,0),0)'),
        dict(key="Days", header="Days Late", width=7, calc=True, fmt="0;-0;0",
             formula='=IF(OR({InScope}=0,{Verified}=1,NOT(ISNUMBER({EffTarget}))),"",sp_AsOf-{EffTarget})'),
        dict(key="DashStatus", header="Dashboard Status", width=15, calc=True,
             formula='=IF({InScope}=0,IF({Status}="Cancelled","Cancelled",""),IF({Verified}=1,"Verified Complete",IF({Overdue}=1,"Overdue",IF({Status}="Complete","Complete - Pending Verification",IF(OR({Status}="In Progress",{Status}="On Hold",N({MSPct})>0),"In Progress","Not Started")))))'),
        dict(key="RAG", header="RAG", width=8, calc=True,
             formula='=IF({InScope}=0,IF({Status}="Cancelled","N/A",""),IF({Verified}=1,"Green",IF({Overdue}=1,"Red",IF(OR({Status}="On Hold",{Verif}="Rejected - Rework"),"Yellow",IF(AND({DashStatus}="Not Started",N({Planned})=0),"Gray",IF(N({Credit})>=N({Planned})-sp_GreenTol,"Green",IF(N({Credit})>=N({Planned})-sp_YellowTol,"Yellow","Red")))))))'),
        dict(key="Issues", header="Data Quality Flags", width=26, calc=True, wrap=True,
             formula=('=IF(AND({Supervisor}="",{Code}=""),"",IF({ID}="","Missing ID; ","")&IF({Supervisor}="","No supervisor; ","")'
                      '&IF({Code}="","No project selected; ",IF({PName}="","Code not in library; ",""))&IF({Coach}="","No coach; ","")'
                      '&IF(NOT(ISNUMBER({Start})),"No start date; ","")&IF(AND({Status}="Complete",{Evid}=""),"Complete w/o evidence; ","")'
                      '&IF(AND({Status}="Complete",NOT(ISNUMBER({Actual}))),"Complete w/o date; ","")'
                      '&IF(AND({Status}="Complete",OR({M1}<>"Done",{M2}<>"Done",{M3}<>"Done",{M4}<>"Done",{M5}<>"Done")),"Complete but milestones open; ","")'
                      '&IF(AND(N({VerSav})>0,{FinOK}<>"Yes"),"Savings not Finance-approved; ","")'
                      '&IF(AND({ID}<>"",COUNTIF(A_ID,{ID})>1),"Duplicate ID; ",""))')),
        dict(key="FinSav", header="Finance-Verified Savings", width=9, calc=True, fmt="$#,##0;($#,##0);-",
             formula='=IF(AND({InScope}=1,{FinOK}="Yes"),N({VerSav}),0)'),
        dict(key="Filter", header="Dashboard Filter", width=6, calc=True, fmt="0",
             formula='=IF({InScope}=0,0,IF(AND(OR(sf_Area="(All)",sf_Area="",{Area}=sf_Area),OR(sf_Shift="(All)",sf_Shift="",{Shift}=sf_Shift),OR(sf_Cat="(All)",sf_Cat="",{Cat}=sf_Cat)),1,0))'),
        dict(key="Key", header="Overdue Key", width=6, calc=True, fmt="0.0",
             formula='=IF({Overdue}=1,N({Days})+ROW()/100000,0)*{Filter}'),
    ]
    acols = Cols(asp)
    adata = demo_asg(AS_OF) if demo else []
    first, last = build_table(ws, wb, acols, 5, N_ASG, "tblAssignments", adata, name_prefix="A", freeze_col="G")
    AL = acols.letter
    rag_cf(ws, f"{AL['RAG']}{first}:{AL['RAG']}{last}")
    rng = f"{AL['M1']}{first}:{AL['M5']}{last}"
    text_cf(ws, rng, {"Done": "Green", "In Progress": "Yellow", "Not Started": "Gray"})
    text_cf(ws, f"{AL['DashStatus']}{first}:{AL['DashStatus']}{last}",
            {"Verified Complete": "Green", "Complete - Pending Verification": "Yellow", "In Progress": "Yellow",
             "Overdue": "Red", "Not Started": "Gray"})
    bar_cf(ws, f"{AL['Credit']}{first}:{AL['Credit']}{last}")
    ws.conditional_formatting.add(f"{AL['Issues']}{first}:{AL['Issues']}{last}",
                                  FormulaRule(formula=[f'LEN(${AL["Issues"]}{first})>0'], font=Font(color="C55A11")))

    # ---------- gap & barrier log ----------
    ws = sh[S_GAP]
    title_band(ws, "Gap & Barrier Log", tag + "  ·  Log anything blocking a supervisor's project. Link it to the "
               "Assignment ID so supervisor and area fill in automatically.", 26)
    gs = [dict(key="ID", header="Gap ID", width=9), dict(key="Date", header="Date Raised", width=10, dv="date", fmt=DATE),
          dict(key="AsgID", header="Assignment ID", width=10),
          dict(key="Sup", header="Supervisor", width=18, calc=True,
               formula='=IFERROR(INDEX(A_Supervisor,MATCH({AsgID},A_ID,0))&"","")'),
          dict(key="Area", header="Area", width=14, calc=True, formula='=IFERROR(INDEX(A_Area,MATCH({AsgID},A_ID,0))&"","")'),
          dict(key="CatA", header="Project Category", width=14, calc=True,
               formula='=IFERROR(INDEX(A_Cat,MATCH({AsgID},A_ID,0))&"","")'),
          dict(key="Cat", header="Barrier Category", width=18, dv="sl_BarCat"),
          dict(key="Desc", header="Gap / Barrier Description", width=32, wrap=True),
          dict(key="Sev", header="Impact", width=8, dv="sl_Sev"),
          dict(key="Support", header="Support Needed From", width=16, dv="sl_Support"),
          dict(key="Owner", header="Action Owner", width=18, dv="sl_People"),
          dict(key="Due", header="Due Date", width=10, dv="date", fmt=DATE),
          dict(key="Status", header="Status", width=11, dv="sl_GStatus", default=None),
          dict(key="Closure", header="Closure Date", width=10, dv="date", fmt=DATE),
          dict(key="Resolution", header="Resolution", width=26, wrap=True),
          dict(key="Counted", header="Counted", width=6, calc=True, fmt="0",
               formula='=IF(AND({Desc}<>"",{Status}<>"Cancelled"),1,0)'),
          dict(key="Open", header="Open", width=6, calc=True, fmt="0",
               formula='=IF(AND({Counted}=1,OR({Status}<>"Closed",NOT(ISNUMBER({Closure})))),1,0)'),
          dict(key="Overdue", header="Overdue", width=6, calc=True, fmt="0",
               formula='=IF(AND({Open}=1,ISNUMBER({Due})),IF({Due}<sp_AsOf,1,0),0)'),
          dict(key="Age", header="Age (days)", width=7, calc=True, fmt="0",
               formula='=IF(OR({Counted}=0,NOT(ISNUMBER({Date}))),"",IF({Open}=0,{Closure},sp_AsOf)-{Date})'),
          dict(key="RAG", header="RAG", width=7, calc=True,
               formula='=IF({Desc}="","",IF({Status}="Cancelled","N/A",IF({Open}=0,"Green",IF(OR({Overdue}=1,{Sev}="High"),"Red",IF(OR({Owner}="",NOT(ISNUMBER({Due}))),"Yellow","Green")))))'),
          dict(key="Issues", header="Data Quality Flags", width=22, calc=True, wrap=True,
               formula='=IF({Desc}="","",IF({AsgID}="","No assignment link; ",IF({Sup}="","Assignment ID not found; ",""))&IF({Owner}="","No owner; ","")&IF(NOT(ISNUMBER({Due})),"No due date; ","")&IF(AND({Status}="Closed",NOT(ISNUMBER({Closure}))),"Closed w/o date; ",""))'),
          dict(key="Filter", header="Dashboard Filter", width=6, calc=True, fmt="0",
               formula='=IF({Counted}=0,0,IF(AND(OR(sf_Area="(All)",sf_Area="",{Area}=sf_Area),OR(sf_Shift="(All)",sf_Shift="",IFERROR(INDEX(A_Shift,MATCH({AsgID},A_ID,0)),"")=sf_Shift),OR(sf_Cat="(All)",sf_Cat="",IFERROR(INDEX(A_Cat,MATCH({AsgID},A_ID,0)),"")=sf_Cat)),1,0))')]
    gdata = demo_gaps(AS_OF, adata) if demo else []
    first, last = build_table(ws, wb, Cols(gs), 5, N_GAP, "tblBarriers", gdata, name_prefix="X", freeze_col="E")
    GL = Cols(gs).letter
    rag_cf(ws, f"{GL['RAG']}{first}:{GL['RAG']}{last}")

    # ---------- 5S audit log ----------
    ws = sh[S_5S]
    title_band(ws, "5S Audit Log", tag + "  ·  Score each S from 0 (not started) to 5 (world class). Total ÷ 25 = 5S score.", 20)
    fs = [dict(key="ID", header="Audit ID", width=9), dict(key="Date", header="Audit Date", width=10, dv="date", fmt=DATE),
          dict(key="Area", header="Area", width=18, dv="sl_Area"), dict(key="Zone", header="Zone / Equipment", width=18),
          dict(key="Sup", header="Area Supervisor", width=18, dv="sl_Supervisors"),
          dict(key="Auditor", header="Auditor", width=18, dv="sl_People")]
    for s in ["Sort", "Set in Order", "Shine", "Standardize", "Sustain"]:
        fs.append(dict(key=s.replace(" ", ""), header=f"{s} (0-5)", width=8, dv="int15", fmt="0"))
    fs += [dict(key="Link", header="Audit Sheet / Photos Link", width=22),
           dict(key="Notes", header="Top Findings", width=26, wrap=True),
           dict(key="Score", header="5S Score", width=8, calc=True, fmt=PCT,
                formula='=IF(OR({Area}="",COUNT({Sort}:{Sustain})<5),"",SUM({Sort}:{Sustain})/25)'),
           dict(key="Meets", header="Meets Target", width=8, calc=True,
                formula='=IF({Score}="","",IF({Score}>=sp_5STarget,"Green",IF({Score}>=sp_5STarget-0.2,"Yellow","Red")))'),
           dict(key="DateNum", header="Date Key", width=6, calc=True, fmt="0",
                formula='=IF(OR({Score}="",NOT(ISNUMBER({Date}))),0,{Date}+ROW()/100000)')]
    fcols = Cols(fs)
    for c in fcols.spec:
        if c.get("calc"):
            c["formula"] = c["formula"].replace("{Sort}:{Sustain}", "{Sort}:{Sustain}")
    fdata = demo_5s(AS_OF) if demo else []
    first, last = build_table(ws, wb, fcols, 5, N_5S, "tbl5S", fdata, name_prefix="F5", freeze_col="E")
    rag_cf(ws, f"{fcols.letter['Meets']}{first}:{fcols.letter['Meets']}{last}")
    bar_cf(ws, f"{fcols.letter['Score']}{first}:{fcols.letter['Score']}{last}")

    build_calc(wb, sh, lc)
    build_dash(wb, sh, demo, lc)
    build_ins(wb, sh, demo)
    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


def build_calc(wb, sh, lc):
    ws = sh[S_CALC]
    CF = qs(sh[S_CFG])
    title_band(ws, "Calc (system – do not edit)", "Helper lists and roll-ups for the Program Dashboard.", 30, legend=False)
    hdr_row(ws, 5, 2, ["Supervisors", "People", "F: Area", "F: Shift", "F: Category"], calc=True)
    for k in range(N_ROS + 1):
        r = 6 + k
        if k < N_ROS:
            ws.cell(r, 2, f'=IFERROR(INDEX(R_Name,MATCH({k + 1},R_SeqSup,0))&"","")')
            ws.cell(r, 3, f'=IFERROR(INDEX(R_Name,MATCH({k + 1},R_SeqAll,0))&"","")')
        if k == 0:
            for c in (4, 5, 6):
                ws.cell(r, c, "(All)")
        elif R0 + k - 1 <= RN:
            lr = R0 + k - 1
            ws.cell(r, 4, f'=IF({CF}!${lc["Area"]}${lr}="","",{CF}!${lc["Area"]}${lr}&"")')
            ws.cell(r, 5, f'=IF({CF}!${lc["Shift"]}${lr}="","",{CF}!${lc["Shift"]}${lr}&"")')
            ws.cell(r, 6, f'=IF({CF}!${lc["Cat"]}${lr}="","",{CF}!${lc["Cat"]}${lr}&"")')
    for c, nm in zip("BCDEF", ["sl_Supervisors", "sl_People", "sl_fArea", "sl_fShift", "sl_fCat"]):
        dyn_list(wb, nm, ws, c, 6, 6 + N_ROS, text_only=True)
        ws.column_dimensions[c].width = 18
    protect(ws)


def build_dash(wb, sh, demo, lc):
    ws = sh[S_DASH]
    CF = qs(sh[S_CFG])
    CA = qs(sh[S_CALC])
    title_band(ws, "Supervisor Productivity Project Tracker — Dashboard", "", 18, legend=False)
    ws["B3"] = ('="Status as of "&TEXT(sp_AsOf,"dd-mmm-yyyy")&"  ·  Period from "&TEXT(sp_Start,"dd-mmm-yy")&"  ·  Target: "&sp_Target&" verified project(s) per supervisor  ·  Completion rate = verified projects ÷ assigned projects"'
                + ('&"  ·  DEMONSTRATION DATA"' if demo else ""))
    for i, w in enumerate([2, 22, 15, 7, 9, 9, 9, 8, 10, 11, 10, 10, 9, 10, 10, 10, 12, 12], 1):
        ws.column_dimensions[L(i)].width = w
    # filters
    for lab, nm, lst, c in [("Area ▼", "sf_Area", "sl_fArea", 2), ("Shift ▼", "sf_Shift", "sl_fShift", 4),
                            ("Project category ▼", "sf_Cat", "sl_fCat", 7)]:
        ws.cell(5, c, lab).font = f(8, True, NAVY)
        v = ws.cell(6, c, "(All)")
        v.font = f(10, True, "0000FF")
        v.fill = fill("FFF9E5")
        v.border = BORDER
        v.alignment = Alignment(horizontal="center")
        if c != 2:
            ws.merge_cells(start_row=6, start_column=c, end_row=6, end_column=c + 1)
        dv = DataValidation(type="list", formula1=f"={lst}", allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(ws.cell(6, c).coordinate)
        define(wb, nm, f"{qs(ws)}!${L(c)}$6")
    ws.cell(6, 10, "Filters apply to every number on this page.").font = f(8, False, GRAY, True)
    # tiles
    tiles = [
        ("PARTICIPATING SUPERVISORS", '=COUNTIF(sc_Name,"?*")', "0", '=COUNTIFS(sc_Name,"?*",sc_Assigned,0)&" without a project"', NAVY),
        ("PROJECTS ASSIGNED", "=SUMPRODUCT(A_InScope,A_Filter)", "0", '=SUMPRODUCT(--(A_DashStatus="In Progress"),A_Filter)&" in progress"', NAVY),
        ("VERIFIED COMPLETE", "=SUMPRODUCT(A_Verified,A_Filter)", "0", '=SUMPRODUCT(--(A_DashStatus="Complete - Pending Verification"),A_Filter)&" awaiting coach verification"', "2E7D32"),
        ("COMPLETION RATE", '=IF(SUMPRODUCT(A_InScope,A_Filter)=0,0,SUMPRODUCT(A_Verified,A_Filter)/SUMPRODUCT(A_InScope,A_Filter))', PCT,
         '="Avg milestone progress "&TEXT(IF(SUMPRODUCT(A_InScope,A_Filter)=0,0,SUMPRODUCT(A_Credit,A_Filter)/SUMPRODUCT(A_InScope,A_Filter)),"0%")', NAVY),
        ("PROGRAM TARGET ATTAINMENT", '=IF(COUNTIF(sc_Name,"?*")*sp_Target=0,0,SUMPRODUCT(sc_TargetMet)/(COUNTIF(sc_Name,"?*")*sp_Target))', PCT,
         '="verified vs "&sp_Target&" per supervisor"', NAVY),
        ("OVERDUE PROJECTS", "=SUMPRODUCT(A_Overdue,A_Filter)", "0", '="past target, not verified"', BURG),
        ("OPEN GAPS / BARRIERS", "=SUMPRODUCT(X_Open,X_Filter)", "0", '=SUMPRODUCT(X_Overdue,X_Filter)&" overdue"', BURG),
        ("FINANCE-VERIFIED SAVINGS", "=SUMPRODUCT(A_FinSav,A_Filter)", "$#,##0", '="Est.: "&TEXT(SUMPRODUCT(A_EstSav,A_InScope,A_Filter),"$#,##0")', "2E7D32"),
    ]
    c = 2
    for i, (lab, fm, fmt, sub, acc) in enumerate(tiles):
        w = 1 if i == 0 else 2
        if i == 0:
            tile(ws, 8, 2, lab, fm, fmt, sub, 1, acc)
            c = 3
        else:
            tile(ws, 8, c, lab, fm, fmt, sub, 2, acc)
            c += 2
    ws.row_dimensions[9].height = 30
    ws.row_dimensions[10].height = 24
    # scorecard
    r0 = 12
    section(ws, r0, 2, "Supervisor scorecard", 17)
    heads = ["Supervisor", "Area", "Shift", "Assigned", "In Progress", "Verified", "Target", "Target Attainment",
             "Completion Rate", "Avg Progress", "Planned %", "Overdue", "Open Gaps", "Verified Savings", "RAG",
             "Current Focus", "Score Key"]
    hdr_row(ws, r0 + 1, 2, heads[:16])
    ws.row_dimensions[r0 + 1].height = 30
    CAc = qs(sh[S_CALC])
    for i in range(NSUP):
        r = r0 + 2 + i
        S = f"$B{r}"
        calc_cell(ws, r, 2, f'=INDEX({CAc}!$B$6:$B${5 + N_ROS},{i + 1})&""', None, True, align="left", size=8)
        g = f'{S}=""'
        def put(cn, fm, fmt="0", al="center"):
            calc_cell(ws, r, 2 + heads.index(cn), f'=IF({g},"",{fm})', fmt, align=al, size=8)
        put("Area", f'IFERROR(INDEX(R_Area,MATCH({S},R_Name,0))&"","")', None)
        put("Shift", f'IFERROR(INDEX(R_Shift,MATCH({S},R_Name,0))&"","")', None)
        put("Assigned", f"SUMPRODUCT(--(A_Supervisor={S}),A_InScope,A_Filter)")
        put("In Progress", f'SUMPRODUCT(--(A_Supervisor={S}),--(A_DashStatus="In Progress"),A_Filter)')
        put("Verified", f"SUMPRODUCT(--(A_Supervisor={S}),A_Verified,A_Filter)")
        put("Target", "sp_Target")
        put("Target Attainment", f"IF(N(sp_Target)=0,0,MIN(1,{L(7)}{r}/sp_Target))", PCT)
        put("Completion Rate", f'IF(N({L(5)}{r})=0,"",{L(7)}{r}/{L(5)}{r})', PCT)
        put("Avg Progress", f'IF(N({L(5)}{r})=0,"",SUMPRODUCT(--(A_Supervisor={S}),A_Credit,A_Filter)/{L(5)}{r})', PCT)
        put("Planned %", f'IF(N({L(5)}{r})=0,"",SUMPRODUCT(--(A_Supervisor={S}),A_Planned,A_Filter)/{L(5)}{r})', PCT)
        put("Overdue", f"SUMPRODUCT(--(A_Supervisor={S}),A_Overdue,A_Filter)")
        put("Open Gaps", f"SUMPRODUCT(--(X_Sup={S}),X_Open,X_Filter)")
        put("Verified Savings", f"SUMPRODUCT(--(A_Supervisor={S}),A_FinSav,A_Filter)", "$#,##0;($#,##0);-")
        put("RAG", (f'IF(N({L(5)}{r})=0,"No project",IF(N({L(13)}{r})>0,"Red",IF(N({L(7)}{r})>=sp_Target,"Green",'
                    f'IF(AND(N({L(11)}{r})=0,N({L(12)}{r})=0),"Gray",IF(N({L(11)}{r})>=N({L(12)}{r})-sp_GreenTol,"Green",'
                    f'IF(N({L(11)}{r})>=N({L(12)}{r})-sp_YellowTol,"Yellow","Red"))))))'), None)
        ws.row_dimensions[r].height = 22
        put("Current Focus", f'IFERROR(INDEX(A_NextMS,MATCH(1,INDEX((A_Supervisor={S})*(A_InScope=1)*(A_Verified=0),0),0)),"–")', None, "left")
    sc0, sc1 = r0 + 2, r0 + 1 + NSUP
    ws.merge_cells(start_row=r0 + 1, start_column=17, end_row=r0 + 1, end_column=18)
    for r in range(sc0, sc1 + 1):
        ws.merge_cells(start_row=r, start_column=17, end_row=r, end_column=18)
    define(wb, "sc_Name", f"{qs(ws)}!$B${sc0}:$B${sc1}")
    define(wb, "sc_Assigned", f"{qs(ws)}!$E${sc0}:$E${sc1}")
    define(wb, "sc_TargetMet", f"{qs(ws)}!$G${sc0}:$G${sc1}")
    # sc_TargetMet must be capped per supervisor -> use helper column S (hidden)
    for r in range(sc0, sc1 + 1):
        c = ws.cell(r, 20, f'=IF($B{r}="",0,MIN(N(sp_Target),N(G{r})))')
        c.font = f(7, False, "BFBFBF")
    define(wb, "sc_TargetMet", f"{qs(ws)}!$T${sc0}:$T${sc1}")
    ws.column_dimensions["T"].hidden = True
    bar_cf(ws, f"I{sc0}:I{sc1}", "2E7D32")
    bar_cf(ws, f"K{sc0}:K{sc1}")
    rag_cf(ws, f"P{sc0}:P{sc1}")
    text_cf(ws, f"P{sc0}:P{sc1}", {"No project": "Red"})
    ws.conditional_formatting.add(f"M{sc0}:M{sc1}", CellIsRule(operator="greaterThan", formula=["0"], font=Font(color="C00000", bold=True)))
    # category & area tables
    r1 = sc1 + 2
    section(ws, r1, 2, "Completion by project category", 8)
    hdr_row(ws, r1 + 1, 2, ["Category", "Assigned", "Verified", "Completion Rate", "Avg Progress", "Overdue", "Open Gaps"])
    for i in range(len(CATS)):
        r = r1 + 2 + i
        lr = R0 + i
        C = f"$B{r}"
        calc_cell(ws, r, 2, f'=IF({CF}!${lc["Cat"]}${lr}="","",{CF}!${lc["Cat"]}${lr}&"")', None, True, align="left", size=8)
        calc_cell(ws, r, 3, f'=IF({C}="","",SUMPRODUCT(--(A_Cat={C}),A_InScope,A_Filter))', "0", size=8)
        calc_cell(ws, r, 4, f'=IF({C}="","",SUMPRODUCT(--(A_Cat={C}),A_Verified,A_Filter))', "0", size=8)
        calc_cell(ws, r, 5, f'=IF(N(C{r})=0,"",D{r}/C{r})', PCT, size=8)
        calc_cell(ws, r, 6, f'=IF(N(C{r})=0,"",SUMPRODUCT(--(A_Cat={C}),A_Credit,A_Filter)/C{r})', PCT, size=8)
        calc_cell(ws, r, 7, f'=IF({C}="","",SUMPRODUCT(--(A_Cat={C}),A_Overdue,A_Filter))', "0", size=8)
        calc_cell(ws, r, 8, f'=IF({C}="","",SUMPRODUCT(--(X_CatA={C}),X_Open,X_Filter))', "0", size=8)
    ce = r1 + 1 + len(CATS)
    bar_cf(ws, f"E{r1 + 2}:E{ce}", "2E7D32")
    bar_cf(ws, f"F{r1 + 2}:F{ce}")
    section(ws, r1, 10, "Completion by area & shift", 9)
    hdr_row(ws, r1 + 1, 10, ["Area", "", "Assigned", "Verified", "Completion Rate", "Avg Progress", "Latest 5S Score", "5S RAG", "Open Gaps"])
    ws.merge_cells(start_row=r1 + 1, start_column=10, end_row=r1 + 1, end_column=11)
    nA = 13
    for i in range(nA):
        r = r1 + 2 + i
        lr = R0 + i
        A = f"$J{r}"
        ws.merge_cells(start_row=r, start_column=10, end_row=r, end_column=11)
        calc_cell(ws, r, 10, f'=IF({CF}!${lc["Area"]}${lr}="","",{CF}!${lc["Area"]}${lr}&"")', None, True, align="left", size=8)
        calc_cell(ws, r, 12, f'=IF({A}="","",SUMPRODUCT(--(A_Area={A}),A_InScope,A_Filter))', "0", size=8)
        calc_cell(ws, r, 13, f'=IF({A}="","",SUMPRODUCT(--(A_Area={A}),A_Verified,A_Filter))', "0", size=8)
        calc_cell(ws, r, 14, f'=IF(N(L{r})=0,"",M{r}/L{r})', PCT, size=8)
        calc_cell(ws, r, 15, f'=IF(N(L{r})=0,"",SUMPRODUCT(--(A_Area={A}),A_Credit,A_Filter)/L{r})', PCT, size=8)
        key = f"SUMPRODUCT(MAX((F5_Area={A})*F5_DateNum))"
        calc_cell(ws, r, 16, f'=IF(OR({A}="",{key}=0),"",INDEX(F5_Score,MATCH({key},F5_DateNum,0)))', PCT, size=8)
        calc_cell(ws, r, 17, f'=IF(P{r}="","",IF(P{r}>=sp_5STarget,"Green",IF(P{r}>=sp_5STarget-0.2,"Yellow","Red")))', None, size=8)
        calc_cell(ws, r, 18, f'=IF({A}="","",COUNTIFS(X_Area,{A},X_Open,1,X_Filter,1))', "0", size=8)
    ae = r1 + 1 + nA
    bar_cf(ws, f"N{r1 + 2}:N{ae}", "2E7D32")
    bar_cf(ws, f"P{r1 + 2}:P{ae}", BURG)
    rag_cf(ws, f"Q{r1 + 2}:Q{ae}")
    # gaps panel
    r2 = max(ce, ae) + 2
    section(ws, r2, 2, "Gaps dashboard – what is blocking completion", 17)
    gt = [("SUPERVISORS W/O PROJECT", '=COUNTIFS(sc_Name,"?*",sc_Assigned,0)', "0"),
          ("BEHIND PLAN (RED)", '=SUMPRODUCT(--(A_RAG="Red"),A_Filter)', "0"),
          ("AWAITING VERIFICATION", '=SUMPRODUCT(--(A_DashStatus="Complete - Pending Verification"),A_Filter)', "0"),
          ("MISSING EVIDENCE (COMPLETE)", '=SUMPRODUCT(--(A_Status="Complete"),--(A_Evid=""),A_Filter)', "0"),
          ("OPEN BARRIERS", "=SUMPRODUCT(X_Open,X_Filter)", "0"),
          ("OVERDUE BARRIERS", "=SUMPRODUCT(X_Overdue,X_Filter)", "0"),
          ("AVG BARRIER AGE (DAYS)", '=IFERROR(SUMIFS(X_Age,X_Open,1,X_Filter,1)/SUMPRODUCT(X_Open,X_Filter),0)', "0"),
          ("DATA-QUALITY FLAGS", '=COUNTIF(A_Issues,"?*")', "0")]
    for i, (lab, fm, fmt) in enumerate(gt):
        c0 = 2 + 2 * i if i else 2
        if i == 0:
            tile(ws, r2 + 1, 2, lab, fm, fmt, "", 1, BURG)
        else:
            tile(ws, r2 + 1, 1 + 2 * i, lab, fm, fmt, "", 2, BURG if i < 6 else GRAY)
    # barrier pareto
    r3 = r2 + 5
    hdr_row(ws, r3, 2, ["Barrier category", "Open", "Total", "Sort key", "Sorted category", "", "", "Count", "Cum. %"])
    nb = 10
    for i in range(nb):
        r = r3 + 1 + i
        lr = R0 + i
        calc_cell(ws, r, 2, f'=IF({CF}!${lc["BarCat"]}${lr}="","",{CF}!${lc["BarCat"]}${lr}&"")', None, align="left", size=8)
        calc_cell(ws, r, 3, f'=IF($B{r}="",0,COUNTIFS(X_Cat,$B{r},X_Open,1,X_Filter,1)+COUNTIFS(A_BarCat,$B{r},A_Filter,1,A_Verified,0))', "0", size=8)
        calc_cell(ws, r, 4, f'=IF($B{r}="",0,COUNTIFS(X_Cat,$B{r},X_Counted,1,X_Filter,1)+COUNTIFS(A_BarCat,$B{r},A_Filter,1))', "0", size=8)
        calc_cell(ws, r, 5, f'=IF($B{r}="",-1,D{r}+({nb}-{i})/1000)', "0.000", size=7, color="BFBFBF")
        rng = f"$E${r3 + 1}:$E${r3 + nb}"
        ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=8)
        calc_cell(ws, r, 6, f'=IF(LARGE({rng},{i + 1})<0,"",INDEX($B${r3 + 1}:$B${r3 + nb},MATCH(LARGE({rng},{i + 1}),{rng},0)))', None, align="left", size=8)
        calc_cell(ws, r, 9, f'=IF(F{r}="",0,INDEX($D${r3 + 1}:$D${r3 + nb},MATCH(F{r},$B${r3 + 1}:$B${r3 + nb},0)))', "0", size=8)
        calc_cell(ws, r, 10, f'=IF(SUM($I${r3 + 1}:$I${r3 + nb})=0,0,SUM($I${r3 + 1}:I{r})/SUM($I${r3 + 1}:$I${r3 + nb}))', PCT, size=8)
    ws.merge_cells(start_row=r3, start_column=6, end_row=r3, end_column=8)
    # overdue list
    hdr_row(ws, r3, 12, ["Overdue project", "", "", "Supervisor", "", "Days Late", "Current milestone"])
    ws.merge_cells(start_row=r3, start_column=12, end_row=r3, end_column=14)
    ws.merge_cells(start_row=r3, start_column=15, end_row=r3, end_column=16)
    for k in range(nb):
        r = r3 + 1 + k
        m = f"MATCH(LARGE(A_Key,{k + 1}),A_Key,0)"
        g = f"LARGE(A_Key,{k + 1})<=0"
        ws.merge_cells(start_row=r, start_column=12, end_row=r, end_column=14)
        ws.merge_cells(start_row=r, start_column=15, end_row=r, end_column=16)
        calc_cell(ws, r, 12, f'=IF({g},"",INDEX(A_ID,{m})&" – "&INDEX(A_PName,{m}))', None, align="left", size=8)
        calc_cell(ws, r, 15, f'=IF({g},"",INDEX(A_Supervisor,{m})&"")', None, align="left", size=8)
        calc_cell(ws, r, 17, f'=IF({g},"",IF(INDEX(A_Days,{m})="","",INDEX(A_Days,{m})))', "0", size=8)
        calc_cell(ws, r, 18, f'=IF({g},"",IF(LEFT(INDEX(A_NextMS,{m}),2)="Al","Verify",LEFT(INDEX(A_NextMS,{m}),2)))', None, size=8)
        ws.row_dimensions[r].height = 22
    # charts
    r4 = r3 + nb + 3
    section(ws, r4, 2, "Charts", 17)
    ch = BarChart()
    ch.type = "bar"
    ch.add_data(Reference(ws, min_col=11, min_row=r0 + 1, max_row=r0 + 1 + 15), titles_from_data=True)
    ch.add_data(Reference(ws, min_col=12, min_row=r0 + 1, max_row=r0 + 1 + 15), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=sc0, max_row=sc0 + 14))
    color_series(ch.series[0], NAVY)
    color_series(ch.series[1], "BFBFBF")
    ch.x_axis.scaling.orientation = "maxMin"
    ch.y_axis.scaling.min = 0
    ch.y_axis.scaling.max = 1
    ch.y_axis.number_format = "0%"
    style_chart(ch, "Supervisor progress vs plan (first 15)", w=13, h=9)
    ws.add_chart(ch, f"B{r4 + 1}")
    cc = BarChart()
    cc.type = "col"
    cc.add_data(Reference(ws, min_col=5, min_row=r1 + 1, max_row=ce), titles_from_data=True)
    cc.add_data(Reference(ws, min_col=6, min_row=r1 + 1, max_row=ce), titles_from_data=True)
    cc.set_categories(Reference(ws, min_col=2, min_row=r1 + 2, max_row=ce))
    color_series(cc.series[0], "2E7D32")
    color_series(cc.series[1], NAVY)
    cc.y_axis.number_format = "0%"
    cc.y_axis.scaling.min = 0
    style_chart(cc, "Completion rate & progress by project category", w=13, h=9)
    ws.add_chart(cc, f"H{r4 + 1}")
    pc = BarChart()
    pc.add_data(Reference(ws, min_col=9, min_row=r3, max_row=r3 + nb), titles_from_data=True)
    pc.set_categories(Reference(ws, min_col=6, min_row=r3 + 1, max_row=r3 + nb))
    color_series(pc.series[0], BURG)
    ln = LineChart()
    ln.add_data(Reference(ws, min_col=10, min_row=r3, max_row=r3 + nb), titles_from_data=True)
    color_series(ln.series[0], NAVY, True)
    ln.y_axis.axId = 200
    ln.y_axis.crosses = "max"
    ln.y_axis.number_format = "0%"
    ln.y_axis.delete = False
    pc += ln
    style_chart(pc, "Barrier Pareto", w=11, h=9)
    ws.add_chart(pc, f"N{r4 + 1}")
    ws.freeze_panes = "A7"
    landscape(ws, f"A1:R{r4 + 20}", one_page=False)
    ws.page_setup.fitToHeight = 2
    protect(ws, ["B6:H6"])


def build_ins(wb, sh, demo):
    ws = sh[S_INS]
    title_band(ws, "Instructions", "How the Supervisor Productivity Program works.", 6, legend=False)
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 110
    rows = [
        ("H", "PURPOSE"),
        ("What it is", "A structured program in which every production supervisor leads improvement projects chosen from a standard library (5S, standard work, problem solving, SMED, CIL, centerlines …). Each project follows 5 milestones with clear definitions of done and is verified by a coach. Projects are linked to DunGPS pillars and, where relevant, to Bronze requirements (e.g. 'one standard work instruction per supervisor')."),
        ("File status", "DEMONSTRATION DATA – fictitious people and results." if demo else "PRODUCTION – library pre-loaded; no names, dates or savings entered."),
        ("H", "HOW TO USE"),
        ("1. Roster", "Enter supervisors, coaches/managers and Finance on 'Supervisor Roster'. Set In Program? = Yes for supervisors who participate. Dropdowns update automatically."),
        ("2. Choose projects", "On 'Project Assignments' pick the Supervisor, Area, Shift and a Project Code from the dropdown. Project name, category, pillar, KPI, standard duration and milestone definitions fill in automatically. Add a specific scope/title."),
        ("3. Plan", "Enter Start Date and Coach. Target Date defaults to Start + library duration (override if needed). Enter the KPI baseline and target."),
        ("4. Update weekly", "Set each milestone (M1–M5) to In Progress / Done, update Status, KPI Actual, Gap/Barrier and Next Action. 'Current Milestone' shows the next definition of done."),
        ("5. Verify", "When all milestones are Done: Status = Complete, paste the Evidence Link, coach sets Coach Verification = Verified. Only then does the project count as complete."),
        ("6. Barriers", "Log blockers on 'Gap & Barrier Log' with the Assignment ID, owner and due date. Review at the weekly tier/production meeting."),
        ("7. 5S audits", "Record 5S audits (0–5 per S) on '5S Audit Log'. The dashboard shows the latest score per area vs. target."),
        ("8. Review", "Use the Program Dashboard filters (Area, Shift, Category). Scorecard = supervisor accountability; gaps panel = what blocks completion."),
        ("H", "DEFINITIONS"),
        ("Milestone %", "Σ milestone weight × (Done = 100%, In Progress = configured credit). Default weights M1 15% · M2 20% · M3 30% · M4 20% · M5 15%."),
        ("Verified complete", "Status Complete AND all 5 milestones Done AND evidence link AND Coach Verification = Verified."),
        ("Completion %", "Verified = 100%; otherwise milestone % capped at 'Max credit before coach verification' (default 90%)."),
        ("Completion rate", "Verified projects ÷ assigned (in-scope, not cancelled) projects."),
        ("Target attainment", "Σ min(verified projects, target) per supervisor ÷ (participating supervisors × target)."),
        ("Planned % / RAG", "Linear plan from start to target date. Green within 10 pts of plan, Yellow within 25 pts, Red beyond or overdue; Gray = not started and not yet due. Supervisor with no project = 'No project' (red)."),
        ("Savings", "Estimated vs Finance-verified are separate; verified savings count only when Finance Approved? = Yes."),
        ("5S score", "Sum of 5 scores (0–5) ÷ 25. Green ≥ target, Yellow within 20 pts, Red below."),
        ("H", "ADMINISTRATION"),
        ("Project Library", "Edit or add project types (row = one project). Set Offered? = No to retire one. The library is a CI recommendation, not a CHG standard – adapt with plant leadership."),
        ("Settings", "Change weights, credit, tolerances and the projects-per-supervisor target on 'Lists & Settings'."),
        ("Excel", "No macros. Excel 2010+ functions only; works in Excel for the web / Teams."),
    ]
    r = 5
    for a, b in rows:
        if a == "H":
            section(ws, r, 2, b, 2)
            r += 1
            continue
        ws.cell(r, 2, a).font = f(9, True, NAVY)
        ws.cell(r, 2).alignment = Alignment(vertical="top")
        c = ws.cell(r, 3, b)
        c.font = f(9)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = max(15, 13 * (len(b) // 120 + 1))
        r += 1


# ---------------- demo data ----------------
def demo_asg(as_of):
    codes = [x[0] for x in LIB]
    weeks = {x[0]: x[6] for x in LIB}
    out = []
    n = 0
    for s in range(1, 10):
        for k in range(random.choice([1, 2, 3])):
            if s == 9:
                break  # supervisor 9 has no project (tests 'No project')
            n += 1
            code = "SW-01" if k == 0 else random.choice(codes)
            start = as_of - dt.timedelta(days=random.randint(10, 150))
            u = random.random()
            ms = ["Not Started"] * 5
            row = dict(ID=f"SP-{n:03d}", Supervisor=f"DEMO Supervisor {s:02d}",
                       Area=["Nugget Line", "Rolled Dough Lines", "Baking", "Packaging", "Mixing", "Sanitation",
                             "Pre-weigh", "Cooling / Freezing", "Gemini Equipment"][s - 1],
                       Shift="ABC"[(s - 1) % 3], Code=code, Scope="DEMO scope", Coach="DEMO Coach 01" if s < 6 else "DEMO Coach 02",
                       Start=start, Baseline=random.choice([None, 50, 60]), KTarget=80,
                       Updated=as_of - dt.timedelta(days=random.randint(0, 10)))
            if u < 0.3:
                ms = ["Done"] * 5
                row.update(Status="Complete", Actual=start + dt.timedelta(days=weeks[code] * 7 - 3), Verif="Verified",
                           VerBy="DEMO Coach 01", VerDate=as_of - dt.timedelta(days=4), KActual=82,
                           Evid="https://example.sharepoint.com/sites/DunGPS/Shared/DEMO_SP.pdf",
                           VerSav=random.choice([0, 15000, 30000]), EstSav=40000, FinOK=random.choice(["Yes", "No"]))
            elif u < 0.4:
                ms = ["Done"] * 5
                row.update(Status="Complete", Actual=as_of - dt.timedelta(days=2), Verif="Submitted - Pending Review",
                           Evid="https://example.sharepoint.com/sites/DunGPS/Shared/DEMO_SP2.pdf")
            elif u < 0.85:
                d = random.randint(1, 4)
                ms = ["Done"] * d + ["In Progress"] + ["Not Started"] * (4 - d)
                row.update(Status="In Progress", KActual=65,
                           BarCat=random.choice(["Time / staffing", "Maintenance support", "Training / skills"]),
                           Gap="DEMO barrier", Next="DEMO next step")
            else:
                row.update(Status="Not Started")
            for j in range(5):
                row[f"M{j + 1}"] = ms[j]
            out.append(row)
    return out


def demo_gaps(as_of, adata):
    out = []
    for i in range(1, 13):
        a = random.choice(adata)
        d0 = as_of - dt.timedelta(days=random.randint(3, 90))
        st = random.choice(["Open", "In Progress", "Closed", "Open"])
        out.append(dict(ID=f"BR-{i:03d}", Date=d0, AsgID=a["ID"],
                        Cat=random.choice(["Time / staffing", "Maintenance support", "Training / skills",
                                           "Data / system (Redzone, CMMS)", "Materials / supplier"]),
                        Desc=f"DEMO barrier {i:03d}", Sev=random.choice(["High", "Medium", "Low"]),
                        Support=random.choice(["Maintenance", "Engineering", "HR / Training"]),
                        Owner="DEMO Coach 02", Due=d0 + dt.timedelta(days=random.randint(10, 40)), Status=st,
                        Closure=(as_of - dt.timedelta(days=1)) if st == "Closed" else None))
    return out


def demo_5s(as_of):
    out = []
    n = 0
    for a in ["Nugget Line", "Packaging", "Baking", "Sanitation", "Maintenance / Parts Room"]:
        base = random.randint(1, 3)
        for m in range(3):
            n += 1
            sc = [min(5, base + m + random.randint(0, 1)) for _ in range(5)]
            out.append(dict(ID=f"5S-{n:03d}", Date=as_of - dt.timedelta(days=90 - 30 * m), Area=a, Zone="DEMO zone",
                            Sup="DEMO Supervisor 01", Auditor="DEMO Coach 01", Sort=sc[0], SetinOrder=sc[1],
                            Shine=sc[2], Standardize=sc[3], Sustain=sc[4]))
    return out


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "."
    build(False, f"{out}/CHG_Supervisor_Productivity_Project_Tracker.xlsx")
    build(True, f"{out}/CHG_Supervisor_Productivity_Project_Tracker_DEMO.xlsx")
    print("built")
