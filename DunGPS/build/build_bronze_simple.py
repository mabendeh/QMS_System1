"""Builds CHG_Bronze_Level_Tracker.xlsx – Bronze Level tracker with one sheet per GPS pillar and a gauge dashboard.

Structure
  Dashboard  – overall + 9 pillar completion gauges (grouped by the 3 GPS layers), pillar table, needs-attention, owners
  Overview   – read-only roll-up of every Bronze requirement from the pillar sheets
  LEAD … SC  – one sheet per pillar: (1) the official Bronze requirements with owner, dates, scheduled duration/hours,
               evidence and a weekly schedule bar; (2) a working log built for that pillar
               (e.g. AM: equipment, equipment type, area/line, scheduled start, due date, scheduled hours, evidence)
  Owners, Lists, How To Use

Usage: python build_bronze_simple.py <out_dir>
"""
import sys
import random
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.chart import DoughnutChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation
from xlcommon import *  # noqa
from build_master import BRONZE, PILLARS, LAYER, PURPOSE

S_D, S_OV, S_O, S_L, S_H = "Dashboard", "Overview", "Owners", "Lists", "How To Use"
ORDER = ["LEAD", "EHS", "FSQ", "AM", "PM", "VO", "E&T", "IM", "SC"]
NAME = {p[0]: p[1] for p in PILLARS}
CODE = {p: p.replace("&", "") for p in ORDER}          # safe for defined names (E&T -> ET)
SPARE = 2                                               # spare requirement rows per pillar
NLOG = 80                                               # working-log rows per pillar
NOWN = 60
WEEKS = 16
LAYER_COLOR = {"Foundation": "8C1D40", "Equipment Focus": "1F4E79", "Capability & Governance": "C55A11"}
HEALTH = {"Evidence Missing": "Red", "Done": "Green", "On Track": "Green", "Needs Verification": "Yellow",
          "At Risk": "Yellow", "Overdue": "Red", "Not Started": "Gray", "N/A": "Gray"}

# Pillar-specific working-log columns (inserted after 'Activity / Task')
EXTRA = {
    "LEAD": ("FMDS boards & tier meetings", [
        dict(key="X1", header="Area / Board", width=18),
        dict(key="X2", header="Tier Level", width=10, dv="bl_Tier"),
        dict(key="X3", header="Meeting Frequency", width=11, dv="bl_Freq"),
        dict(key="X4", header="Attendance %", width=10, dv="pct", fmt=PCT)]),
    "EHS": ("SafeStart roll-out & 21 Elements", [
        dict(key="X1", header="Element / Program (e.g. Element 4 – …)", width=24, wrap=True),
        dict(key="X2", header="Area", width=16, dv="bl_Area"),
        dict(key="X3", header="People Trained", width=9, fmt="0", dv="num"),
        dict(key="X4", header="People Required", width=9, fmt="0", dv="num"),
        dict(key="X5", header="Coverage", width=8, calc=True, fmt=PCT,
             formula='=IF(N({X4})=0,"",MIN(1,N({X3})/{X4}))')]),
    "FSQ": ("FSQ program elements, audits & records", [
        dict(key="X1", header="Program / Element", width=22, wrap=True),
        dict(key="X2", header="Area", width=16, dv="bl_Area"),
        dict(key="X3", header="Audit / Check Score %", width=10, dv="pct", fmt=PCT),
        dict(key="X4", header="Check Frequency", width=11, dv="bl_Freq")]),
    "AM": ("equipment – CIL, centerlines, defect handling, operator training, deep clean, LOTO & safety maps", [
        dict(key="X1", header="Equipment / Asset", width=20),
        dict(key="X2", header="Equipment Type", width=16, dv="bl_EqType"),
        dict(key="X3", header="Area / Line", width=16, dv="bl_Area"),
        dict(key="X4", header="Operators Trained", width=9, fmt="0", dv="num"),
        dict(key="X5", header="Open Abnormality Tags", width=9, fmt="0", dv="num")]),
    "PM": ("parts room 2S, Limble / CMMS assets, maintenance organisation", [
        dict(key="X1", header="Asset / Zone / Role", width=20),
        dict(key="X2", header="Equipment Type", width=16, dv="bl_EqType"),
        dict(key="X3", header="Limble / CMMS Ref", width=12),
        dict(key="X4", header="PM Compliance %", width=10, dv="pct", fmt=PCT)]),
    "VO": ("loss elimination – lines, tools and results", [
        dict(key="X1", header="Line / Asset", width=18),
        dict(key="X2", header="Area", width=16, dv="bl_Area"),
        dict(key="X3", header="Loss Type", width=12, dv="bl_Loss"),
        dict(key="X4", header="Baseline", width=8, fmt="#,##0.0"),
        dict(key="X5", header="Target", width=8, fmt="#,##0.0"),
        dict(key="X6", header="Actual", width=8, fmt="#,##0.0")]),
    "E&T": ("standard work, skill matrices & leader standard work – by person", [
        dict(key="X1", header="Supervisor / Employee", width=20, dv="bl_Owners"),
        dict(key="X2", header="Area", width=16, dv="bl_Area"),
        dict(key="X3", header="Task / Skill / Document", width=22, wrap=True),
        dict(key="X4", header="Qualified?", width=9, dv="bl_YesNo")]),
    "IM": ("monthly Plant Steering Committee", [
        dict(key="X1", header="Meeting Month", width=10, dv="date", fmt="mmm-yy"),
        dict(key="X2", header="Ideas Reviewed", width=9, fmt="0", dv="num"),
        dict(key="X3", header="Projects Approved", width=9, fmt="0", dv="num"),
        dict(key="X4", header="Attendees", width=9, fmt="0", dv="num")]),
    "SC": ("process governance & gap analysis", [
        dict(key="X1", header="Process", width=22, wrap=True),
        dict(key="X2", header="Gap Assessment Score %", width=10, dv="pct", fmt=PCT),
        dict(key="X3", header="Gaps Found", width=8, fmt="0", dv="num"),
        dict(key="X4", header="Gaps Closed", width=8, fmt="0", dv="num")]),
}

DEMO_LOG = {
    "AM": lambda i: dict(X1=f"DEMO {['Mixer 1', 'Sheeter 2', 'Nugget Former', 'Spiral Freezer', 'Bagger 3', 'Oven 1'][i % 6]}",
                         X2=["Mixer", "Sheeter / Former", "Sheeter / Former", "Spiral Freezer", "Packaging Machine", "Oven"][i % 6],
                         X3=["Mixing", "Rolled Dough Lines", "Nugget Line", "Cooling / Freezing", "Packaging", "Baking"][i % 6],
                         X4=random.randint(2, 8), X5=random.randint(0, 6)),
    "PM": lambda i: dict(X1=["DEMO Parts room zone A", "DEMO Parts room zone B", "DEMO Mixer 1", "DEMO Oven 1",
                             "DEMO Planner/Scheduler role"][i % 5], X2=["Other", "Other", "Mixer", "Oven", "Other"][i % 5],
                         X3=f"LMB-{100 + i}", X4=random.choice([0.6, 0.8, 0.95])),
    "LEAD": lambda i: dict(X1=f"DEMO {['Nugget Line board', 'Packaging board', 'Plant Tier 3'][i % 3]}",
                           X2=["Tier 1", "Tier 1", "Tier 3"][i % 3], X3=["Daily", "Daily", "Weekly"][i % 3],
                           X4=random.choice([0.7, 0.85, 1.0])),
    "EHS": lambda i: dict(X1=f"DEMO {'SafeStart wave ' + str(i + 1) if i < 3 else 'Element ' + str(i)}",
                          X2="Site-wide", X3=random.randint(10, 40), X4=40),
    "FSQ": lambda i: dict(X1=f"DEMO {['GMP audit', 'Allergen changeover', 'Metal detection checks', 'Hold & release log', 'SPC chart – weight'][i % 5]}",
                          X2=["Packaging", "Mixing", "Packaging", "Laboratory / QA", "Packaging"][i % 5],
                          X3=random.choice([0.78, 0.86, 0.93]), X4=random.choice(["Weekly", "Monthly"])),
    "VO": lambda i: dict(X1=f"DEMO {['Nugget Line', 'Bagger 3', 'Oven 1'][i % 3]}", X2=["Nugget Line", "Packaging", "Baking"][i % 3],
                         X3=["Changeover", "Availability", "Quality"][i % 3], X4=95, X5=60, X6=random.choice([70, 65, 88])),
    "E&T": lambda i: dict(X1=f"DEMO Owner {random.randint(1, 9):02d}", X2="Packaging", X3="DEMO: standard work – film change",
                          X4=random.choice(["Yes", "No"])),
    "IM": lambda i: dict(X1=dt.date(2026, 8 + i, 1) if i < 4 else None, X2=random.randint(4, 12), X3=random.randint(0, 3),
                         X4=random.randint(6, 10)),
    "SC": lambda i: dict(X1=["DEMO PONC tracking", "DEMO SC governance review", "DEMO Procurement approvals"][i % 3],
                         X2=random.choice([0.55, 0.7, 0.82]), X3=random.randint(3, 9), X4=random.randint(1, 5)),
}


def health_formula(appl, done, overdue, status, link, due, prog):
    return (f'IF({appl}=0,IF({status}="N/A","N/A",""),IF({done}=1,"Done",IF(AND({status}="Complete",{link}=""),"Evidence Missing",'
            f'IF({overdue}=1,"Overdue",IF({status}="Complete","Needs Verification",IF(AND(ISNUMBER({due}),{due}-AsOfB<=set_RiskDays,N({prog})<0.75),"At Risk",'
            f'IF(AND({status}="Not Started",N({prog})=0),"Not Started","On Track")))))))')


def evidence_dv(ws, rng_status, rng_verified, link_cell):
    msg = ("Upload the evidence first: click 'Upload ↗', save the file in the Teams folder, copy its link into "
           "'Evidence Link'. Then 'Complete' / 'Yes' become available.")
    for rng, ok, ne in [(rng_status, "bl_Status", "bl_StatusNE"), (rng_verified, "bl_YesNo", "bl_NoOnly")]:
        d = DataValidation(type="list", formula1=f'=IF(LEN({link_cell})>0,{ok},{ne})', allow_blank=True,
                           showErrorMessage=True, errorTitle="Evidence required", error=msg)
        ws.add_data_validation(d)
        d.add(rng)


def upload_formula(guard, pillar):
    return (f'=IF({guard},"",IF(IFERROR(INDEX(bl_Folder,MATCH("{pillar}",bl_PillarCode,0))&"","")<>"",'
            f'HYPERLINK(INDEX(bl_Folder,MATCH("{pillar}",bl_PillarCode,0)),"Upload ↗"),'
            f'IF(set_EvidRoot<>"",HYPERLINK(set_EvidRoot,"Upload ↗"),"Set folder in Lists")))')


def req_spec(p, cd):
    return [
        dict(key="Ref", header="Ref", width=8),
        dict(key="Req", header="Bronze Requirement (official)", width=38, wrap=True),
        dict(key="Owner", header="Owner", width=18, dv="bl_Owners"),
        dict(key="Start", header="Start Date", width=10, dv="date", fmt=DATE),
        dict(key="Due", header="Due Date", width=10, dv="date", fmt=DATE),
        dict(key="Dur", header="Scheduled Duration (days)", width=9, calc=True, fmt="0",
             formula='=IF(AND(ISNUMBER({Start}),ISNUMBER({Due})),{Due}-{Start},"")'),
        dict(key="Hours", header="Scheduled Hours (from log)", width=9, calc=True, fmt="0.0",
             formula=f'=IF({{Req}}="","",SUMIFS(p{cd}_D_Hours,p{cd}_D_Act,{{Ref}}))'),
        dict(key="DaysLeft", header="Days Left (− = late)", width=8, calc=True, fmt="0;[Red]-0;0",
             formula='=IF(OR({Req}="",{Done}=1,NOT(ISNUMBER({Due}))),"",{Due}-AsOfB)'),
        dict(key="Status", header="Status", width=12, default="Not Started",
             note="'Complete' only appears after an Evidence Link is pasted in this row."),
        dict(key="Pct", header="% Complete (if no log rows)", width=9, dv="bl_Pct", fmt=PCT, default=0),
        dict(key="LogN", header="Log Activities", width=7, calc=True, fmt="0",
             formula=f'=IF({{Req}}="","",COUNTIFS(p{cd}_D_Act,{{Ref}},p{cd}_D_Appl,1))'),
        dict(key="LogDone", header="Log Done", width=7, calc=True, fmt="0",
             formula=f'=IF({{Req}}="","",COUNTIFS(p{cd}_D_Act,{{Ref}},p{cd}_D_Done,1))'),
        dict(key="Prog", header="Progress", width=8, calc=True, fmt=PCT,
             formula='=IF(OR({Req}="",{Status}="N/A"),"",IF({Done}=1,1,IF(N({LogN})>0,{LogDone}/{LogN},IF({Status}="Complete",MAX(0.75,N({Pct})),N({Pct})))))'),
        dict(key="Upload", header="Upload Evidence", width=10, calc=True, formula=upload_formula('{Req}=""', p)),
        dict(key="Link", header="Evidence Link (paste Teams link)", width=22,
             note="1) Click Upload ↗  2) save the file in Teams  3) … → Copy link  4) paste here."),
        dict(key="EvDesc", header="Evidence Description", width=22, wrap=True),
        dict(key="Verified", header="Verified?", width=9, default="No",
             note="'Yes' is only available when an Evidence Link is present."),
        dict(key="VerBy", header="Verified By", width=15, dv="bl_Owners"),
        dict(key="Gap", header="Gap / What's Missing", width=24, wrap=True),
        dict(key="Next", header="Next Step", width=24, wrap=True),
        dict(key="Health", header="Health", width=15, calc=True,
             formula="=" + health_formula("{Appl}", "{Done}", "{Overdue}", "{Status}", "{Link}", "{Due}", "{Prog}")),
        dict(key="Missing", header="Missing Info", width=18, calc=True, wrap=True,
             formula=('=IF(OR({Appl}=0,{Done}=1),"",IF({Owner}="","Owner; ","")&IF(NOT(ISNUMBER({Due})),"Due date; ","")'
                      '&IF(AND({Status}="Complete",{Link}=""),"Evidence link; ","")&IF(AND({Status}="Complete",{Verified}<>"Yes"),"Verification; ",""))')),
        dict(key="Appl", header="Counts?", width=5, calc=True, fmt="0",
             formula='=IF(AND({Req}<>"",{Status}<>"N/A"),1,0)'),
        dict(key="Done", header="Done", width=5, calc=True, fmt="0",
             formula='=IF(AND({Appl}=1,{Status}="Complete",{Verified}="Yes",{Link}<>""),1,0)'),
        dict(key="Overdue", header="Overdue", width=5, calc=True, fmt="0",
             formula='=IF(AND({Appl}=1,{Done}=0,ISNUMBER({Due})),IF({Due}<AsOfB,1,0),0)'),
        dict(key="OpenDue", header="Open Due", width=5, calc=True, fmt=DATE,
             formula='=IF(AND({Appl}=1,{Done}=0,ISNUMBER({Due})),{Due},"")'),
    ]


def log_spec(p, cd):
    ls = [dict(key="ID", header="Log ID", width=8),
          dict(key="Act", header="Bronze Ref (select)", width=10, dv=f"p{cd}_Refs"),
          dict(key="ActName", header="Bronze Requirement", width=30, calc=True, wrap=True,
               formula=f'=IFERROR(INDEX(p{cd}_R_Req,MATCH({{Act}},p{cd}_R_Ref,0))&"","")'),
          dict(key="Task", header="Activity / Task", width=26, wrap=True)]
    ls += [dict(x) for x in EXTRA[p][1]]
    ls += [dict(key="Owner", header="Owner", width=16, dv="bl_Owners"),
           dict(key="Start", header="Scheduled Start", width=10, dv="date", fmt=DATE),
           dict(key="Due", header="Due Date", width=10, dv="date", fmt=DATE),
           dict(key="Hours", header="Scheduled Time (hrs)", width=9, dv="num", fmt="0.0"),
           dict(key="Status", header="Status", width=12,
                note="'Complete' only appears after an Evidence Link is pasted in this row."),
           dict(key="Upload", header="Upload Evidence", width=10, calc=True, formula=upload_formula('{Act}=""', p)),
           dict(key="Link", header="Evidence Link (paste Teams link)", width=22),
           dict(key="Verified", header="Verified?", width=9),
           dict(key="Notes", header="Notes", width=22, wrap=True),
           dict(key="DaysLeft", header="Days Left (− = late)", width=8, calc=True, fmt="0;[Red]-0;0",
                formula='=IF(OR({Act}="",{Done}=1,NOT(ISNUMBER({Due}))),"",{Due}-AsOfB)'),
           dict(key="Health", header="Health", width=15, calc=True,
                formula="=" + health_formula("{Appl}", "{Done}", "{Overdue}", "{Status}", "{Link}", "{Due}",
                                             'IF({Status}="Complete",1,0)')),
           dict(key="Appl", header=" ", width=2, calc=True, fmt=";;;",
                formula='=IF(AND({Act}<>"",{Status}<>"N/A"),1,0)'),
           dict(key="Done", header="Done", width=6, calc=True, fmt='"✓";;;',
                formula='=IF(AND({Appl}=1,{Status}="Complete",{Verified}="Yes",{Link}<>""),1,0)'),
           dict(key="Overdue", header="Late", width=6, calc=True, fmt='"⚠";;;',
                formula='=IF(AND({Appl}=1,{Done}=0,ISNUMBER({Due})),IF({Due}<AsOfB,1,0),0)')]
    return ls


def build(demo, path):
    random.seed(11)
    AS_OF = dt.date(2026, 11, 18)
    wb = Workbook()
    wb.remove(wb.active)
    names = [S_D, S_OV] + ORDER + [S_O, S_L, S_H]
    sh = {n: wb.create_sheet(n) for n in names}
    for n in names:
        sh[n].sheet_properties.tabColor = NAVY if n in (S_D, S_OV) else (
            LAYER_COLOR[LAYER[n]] if n in ORDER else "7F7F7F")
    tag = "DEMONSTRATION DATA – fictitious" if demo else "C.H. Guenther & Son — Duncanville"

    # ======================= LISTS =======================
    ws = sh[S_L]
    title_band(ws, "Lists & Settings", "Settings, dropdown values and gauge data. Add new values at the bottom of a list.", 24,
               legend=False)
    settings = [
        ("Status date (blank = today)", AS_OF if demo else None, DATE, None),
        ("ƒ Status date used", "=IF(ISNUMBER(C5),C5,TODAY())", DATE, "AsOfB"),
        ("'At Risk' warning window (days before due)", 14, "0", "set_RiskDays"),
        ("Evidence folder – Teams/SharePoint URL (all pillars)",
         "https://example.sharepoint.com/sites/DunGPS/Shared%20Documents/General/Shared" if demo else None, None, "set_EvidRoot"),
        ("Gauge: red below", 0.5, PCT, "set_GaugeRed"),
        ("Gauge: green at or above", 0.9, PCT, "set_GaugeGreen"),
    ]
    for i, (lab, val, fmt, nm) in enumerate(settings):
        r = 5 + i
        calc = lab.startswith("ƒ")
        ws.cell(r, 2, lab).font = f(9, calc)
        c = ws.cell(r, 3, val)
        c.border = BORDER
        c.font = f(9, False, "000000" if calc else "0000FF")
        c.fill = fill(LGRAY if calc else WHITE)
        c.alignment = Alignment(horizontal="center")
        if fmt:
            c.number_format = fmt
        if nm:
            define(wb, nm, f"{qs(ws)}!$C${r}")
    ws["E8"] = "Teams → DunGPS → General → Shared. Per-pillar folders (column N below) override it."
    ws["E8"].font = f(8, False, GRAY, True)
    ws.column_dimensions["B"].width = 42
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 3
    LR0 = 13
    lists = [
        ("E", "Pillar", ORDER), ("F", "Pillar Name", [NAME[p] for p in ORDER]),
        ("G", "Status", ["Not Started", "In Progress", "Complete", "N/A"]), ("H", "Yes / No", ["Yes", "No"]),
        ("I", "% Complete", [0, 0.25, 0.5, 0.75, 1]),
        ("L", "Layer", [LAYER[p] for p in ORDER]), ("M", "Pillar Purpose (GPS poster)", [PURPOSE[p] for p in ORDER]),
        ("N", "Pillar Upload Folder URL",
         [f"https://example.sharepoint.com/sites/DunGPS/Shared%20Documents/General/Shared/{CODE[p]}" for p in ORDER] if demo else []),
        ("O", "Status (no evidence yet)", ["Not Started", "In Progress", "N/A"]), ("P", "Verified (no evidence yet)", ["No"]),
        ("Q", "Equipment Type", ["Mixer", "Sheeter / Former", "Gemini Equipment", "Nugget Line Equipment", "Oven",
                                 "Spiral Freezer", "Conveyor", "Packaging Machine", "Metal Detector / X-ray",
                                 "Depositor / Filler", "Slicer", "Utilities", "Other"]),
        ("R", "Area / Line", ["Pre-weigh", "Mixing", "Dough Processing / Forming", "Gemini Equipment", "Nugget Line",
                              "Rolled Dough Lines", "Baking", "Cooling / Freezing", "Packaging", "Sanitation",
                              "Maintenance / Parts Room", "Laboratory / QA", "Warehouse / Shipping", "Site-wide"]),
        ("S", "Tier Level", ["Tier 1", "Tier 2", "Tier 3", "Plant"]),
        ("T", "Frequency", ["Per shift", "Daily", "Weekly", "Monthly", "Quarterly"]),
        ("U", "Loss Type", ["Availability", "Performance", "Quality", "Changeover", "Other"]),
    ]
    for colL, h, vals in lists:
        col = ws[colL + "1"].column
        hdr_row(ws, LR0, col, [h])
        ws.column_dimensions[colL].width = {"Pillar Name": 28, "Layer": 22, "Pillar Purpose (GPS poster)": 46,
                                            "Pillar Upload Folder URL": 36, "Equipment Type": 20, "Area / Line": 22}.get(h, 13)
        for i in range(20):
            c = ws.cell(LR0 + 1 + i, col, vals[i] if i < len(vals) else None)
            c.border = BORDER
            c.font = f(9)
            if h == "% Complete":
                c.number_format = PCT
    a, b = LR0 + 1, LR0 + 20
    for nm, colL in [("bl_Pillar", "E"), ("bl_Status", "G"), ("bl_YesNo", "H"), ("bl_Pct", "I"), ("bl_EqType", "Q"),
                     ("bl_Area", "R"), ("bl_Tier", "S"), ("bl_Freq", "T"), ("bl_Loss", "U")]:
        dyn_list(wb, nm, ws, colL, a, b)
    define(wb, "bl_PillarCode", f"{qs(ws)}!$E${a}:$E${b}")
    define(wb, "bl_Folder", f"{qs(ws)}!$N${a}:$N${b}")
    define(wb, "bl_StatusNE", f"{qs(ws)}!$O${a}:$O${a + 2}")
    define(wb, "bl_NoOnly", f"{qs(ws)}!$P${a}:$P${a}")
    ws.cell(b + 1, 6, "Pillar names, order, layers and purposes: CHG GPS poster 'How the 9 GPS (TPM) pillars work together'.").font = f(8, False, GRAY, True)
    hdr_row(ws, LR0, 11, ["Owner list (auto)"], calc=True)
    ws.column_dimensions["K"].width = 22
    for k in range(NOWN):
        ws.cell(LR0 + 1 + k, 11, f'=IFERROR(INDEX(OW_Name,MATCH({k + 1},OW_Seq,0))&"","")').font = f(8, False, GRAY)
    dyn_list(wb, "bl_Owners", ws, "K", LR0 + 1, LR0 + NOWN, text_only=True)
    define(wb, "bl_OwnerAll", f"{qs(ws)}!$K${LR0 + 1}:$K${LR0 + NOWN}")
    # gauge data (W..AE)
    G0 = 23
    hdr_row(ws, LR0, G0, ["Gauge", "Done", "Counted", "Rate", "Red", "Amber", "Green", "Rest", "Hidden"], calc=True)
    gauge_rows = {}
    for i, p in enumerate(["ALL"] + ORDER):
        r = LR0 + 1 + i
        gauge_rows[p] = r
        ws.cell(r, G0, p)
        if p == "ALL":
            ws.cell(r, G0 + 1, "=SUM(OV_Done)")
            ws.cell(r, G0 + 2, "=SUM(OV_Appl)")
        else:
            ws.cell(r, G0 + 1, f'=COUNTIFS(OV_Pillar,"{p}",OV_Done,1)')
            ws.cell(r, G0 + 2, f'=COUNTIFS(OV_Pillar,"{p}",OV_Appl,1)')
        X, Y, R_ = L(G0 + 1), L(G0 + 2), f"{L(G0 + 3)}{r}"
        ws.cell(r, G0 + 3, f"=IF({Y}{r}=0,0,{X}{r}/{Y}{r})").number_format = PCT
        ws.cell(r, G0 + 4, f"=IF({R_}<set_GaugeRed,{R_},0)")
        ws.cell(r, G0 + 5, f"=IF(AND({R_}>=set_GaugeRed,{R_}<set_GaugeGreen),{R_},0)")
        ws.cell(r, G0 + 6, f"=IF({R_}>=set_GaugeGreen,{R_},0)")
        ws.cell(r, G0 + 7, f"=1-{R_}")
        ws.cell(r, G0 + 8, 1)
        for c in range(G0, G0 + 9):
            ws.cell(r, c).font = f(8, False, GRAY)
    for c in range(G0, G0 + 9):
        ws.column_dimensions[L(c)].width = 8

    # ======================= OWNERS =======================
    ws = sh[S_O]
    title_band(ws, "Owners", "Type each person once. They appear in every Owner dropdown. Set Active? = No to remove them from dropdowns.", 8)
    os_ = [dict(key="Name", header="Name", width=26), dict(key="Role", header="Role / Position", width=24),
           dict(key="Dept", header="Department", width=18), dict(key="Pillar", header="Pillar", width=8, dv="bl_Pillar"),
           dict(key="Email", header="Email", width=28),
           dict(key="Active", header="Active?", width=9, dv="bl_YesNo", default="Yes")]
    oc = Cols(os_)
    os_.append(dict(key="Seq", header="#", width=5, calc=True, fmt="0",
                    formula=f'=IF(AND({{Name}}<>"",{{Active}}<>"No"),COUNTIFS(${oc.letter["Name"]}$6:{{Name}},"?*",${oc.letter["Active"]}$6:{{Active}},"<>No"),"")'))
    odata = []
    if demo:
        roles = [("Plant Manager", "Plant Leadership", "LEAD"), ("AM Pillar Lead", "Production", "AM"),
                 ("Maintenance Manager", "Maintenance", "PM"), ("FSQ Manager", "Quality / FSQ", "FSQ"),
                 ("EHS Manager", "EHS", "EHS"), ("Training Coordinator", "HR / Training", "E&T"),
                 ("CI Engineer", "Engineering", "VO"), ("Supply Chain Manager", "Supply Chain", "SC"),
                 ("QA Auditor (verifier)", "Quality / FSQ", "FSQ")]
        for i, (role, dept, p) in enumerate(roles, 1):
            odata.append(dict(Name=f"DEMO Owner {i:02d}", Role=role, Dept=dept, Pillar=p,
                              Email=f"demo{i:02d}@example.com", Active="Yes"))
    build_table(ws, wb, Cols(os_), 5, NOWN, "tblOwnersB", odata, name_prefix="OW", freeze_col="C")

    # ======================= PILLAR SHEETS =======================
    owner_of = {"LEAD": 1, "EHS": 5, "FSQ": 4, "AM": 2, "PM": 3, "VO": 7, "E&T": 6, "IM": 1, "SC": 8}
    req_slots = []
    week0 = "(AsOfB-WEEKDAY(AsOfB,3)-28)"  # Monday, 4 weeks before the status week
    HR = 10
    for p in ORDER:
        ws = sh[p]
        cd = CODE[p]
        reqs = BRONZE[p]
        n = len(reqs) + SPARE
        title_band(ws, f"{p} — {NAME[p]}", f"{tag}  ·  Layer: {LAYER[p]}  ·  Purpose: {PURPOSE[p]}", 30)
        ws.cell(4, 2, "Pillar Lead ▸").font = f(9, True, NAVY)
        lc = ws.cell(4, 3, f"DEMO Owner {owner_of[p]:02d}" if demo else None)
        lc.border = BORDER
        lc.font = f(9, True, "0000FF")
        d = DataValidation(type="list", formula1="=bl_Owners", allow_blank=True)
        ws.add_data_validation(d)
        d.add("C4")
        ws.cell(4, 5, "Evidence is mandatory: Upload ↗ → save in Teams → paste link → then Complete.   "
                      "Done = Complete + Evidence Link + Verified? Yes.").font = f(8, True, BURG)
        # ---- section 1: official requirements
        rs = req_spec(p, cd)
        rdata = [dict(Ref=f"{p}-{i}", Req=rq, Status="Not Started", Pct=0, Verified="No") for i, rq in enumerate(reqs, 1)]
        if demo:
            for i, row in enumerate(rdata):
                row["Owner"] = f"DEMO Owner {owner_of[p]:02d}" if not (p == "FSQ" and i == 1) else None
                row["Start"] = AS_OF - dt.timedelta(days=random.randint(30, 110))
                row["Due"] = AS_OF + dt.timedelta(days=random.randint(-35, 110))
                u = random.random()
                if u < 0.5:
                    row.update(Status="Complete", Pct=1, Verified="Yes", VerBy="DEMO Owner 09",
                               Link=f"https://example.sharepoint.com/sites/DunGPS/Shared/{cd}/DEMO_{p}-{i + 1}.pdf",
                               EvDesc="DEMO: signed checklist + photos")
                elif u < 0.58:
                    row.update(Status="Complete", Pct=1, Verified="No",
                               Link=f"https://example.sharepoint.com/sites/DunGPS/Shared/{cd}/DEMO_{p}-{i + 1}.pdf",
                               EvDesc="DEMO: evidence uploaded")
                elif u < 0.8:
                    row.update(Status="In Progress", Pct=random.choice([0.25, 0.5, 0.75]),
                               Gap="DEMO: what is missing", Next="DEMO: next step")
            if p == "FSQ":
                rdata[2].update(Status="Complete", Pct=1, Verified="No", Link=None, EvDesc=None)  # -> Evidence Missing
            if p == "SC":
                rdata[-1]["Status"] = "N/A"
        rcols = Cols(rs)
        first, last = build_table(ws, wb, rcols, HR, n, f"tblReq_{cd}", rdata, name_prefix=f"p{cd}_R", row_height=30)
        RL = rcols.letter
        for r in range(first, last + 1):
            req_slots.append((p, r))
        evidence_dv(ws, f"{RL['Status']}{first}:{RL['Status']}{last}", f"{RL['Verified']}{first}:{RL['Verified']}{last}",
                    f"${RL['Link']}{first}")
        text_cf(ws, f"{RL['Health']}{first}:{RL['Health']}{last}", HEALTH)
        text_cf(ws, f"{RL['Status']}{first}:{RL['Status']}{last}", {"Complete": "Green", "In Progress": "Yellow"})
        bar_cf(ws, f"{RL['Prog']}{first}:{RL['Prog']}{last}")
        ws.conditional_formatting.add(f"{RL['Missing']}{first}:{RL['Missing']}{last}",
                                      FormulaRule(formula=[f'LEN(${RL["Missing"]}{first})>0'], font=Font(color="C55A11", bold=True)))
        section(ws, HR - 1, 2, "1. Bronze requirements (official – CHG GPS Pillar Bronze Level Requirements)", 12)
        dyn_list(wb, f"p{cd}_Refs", ws, RL["Ref"], first, last)
        # ---- weekly schedule bar
        g0 = rcols.first_col + len(rs) + 1
        ws.cell(HR - 1, g0, "Schedule by week (Monday) – blue planned · green done · red overdue · ▲ this week").font = f(8, True, NAVY)
        for w in range(WEEKS):
            c = g0 + w
            ws.column_dimensions[L(c)].width = 4.5
            hc = ws.cell(HR, c, f"={week0}+{7 * w}")
            hc.number_format = "dd-mmm"
            hc.font = f(7, True, WHITE)
            hc.fill = fill(NAVY)
            hc.alignment = Alignment(text_rotation=90, horizontal="center", vertical="center")
            wk = f"{L(c)}${HR}"
            for r in range(first, last + 1):
                S_, D_ = f"${RL['Start']}{r}", f"${RL['Due']}{r}"
                DN, OD = f"${RL['Done']}{r}", f"${RL['Overdue']}{r}"
                fm = (f'=IF(OR(NOT(ISNUMBER({D_})),${RL["Req"]}{r}=""),"",IF(AND(IF(ISNUMBER({S_}),{S_},{D_})<={wk}+6,{D_}>={wk}),'
                      f'IF({DN}=1,2,IF({OD}=1,3,1)),IF(AND({OD}=1,{wk}>{D_},{wk}<=AsOfB),3,"")))')
                cc = ws.cell(r, c, fm)
                cc.number_format = ";;;"
                hair = Side(style="hair", color="D9D9D9")
                cc.border = Border(left=hair, right=hair, top=hair, bottom=hair)
            mc = ws.cell(last + 1, c, f'=IF(AND(AsOfB>={wk},AsOfB<{wk}+7),"▲","")')
            mc.font = f(9, True, BURG)
            mc.alignment = Alignment(horizontal="center")
        grng = f"{L(g0)}{first}:{L(g0 + WEEKS - 1)}{last}"
        for v, colr in [(1, "9DC3E6"), (2, "70AD47"), (3, "E06666")]:
            ws.conditional_formatting.add(grng, CellIsRule(operator="equal", formula=[str(v)], fill=fill(colr)))
        # ---- section 2: pillar-specific working log
        LH = last + 5
        section(ws, LH - 1, 2, f"2. {p} working log – {EXTRA[p][0]}  (one row per activity, linked to a Bronze Ref)", 12)
        ls = log_spec(p, cd)
        ldata = []
        if demo:
            nlog = {"AM": 10, "PM": 6, "FSQ": 8, "VO": 6, "E&T": 6, "EHS": 5, "LEAD": 3, "IM": 4, "SC": 3}[p]
            for i in range(nlog):
                row = dict(ID=f"{cd}-L{i + 1:02d}", Act=f"{p}-{1 + (i % len(reqs))}", Task=f"DEMO task {i + 1}",
                           Owner=f"DEMO Owner {owner_of[p]:02d}", Start=AS_OF - dt.timedelta(days=random.randint(5, 60)),
                           Due=AS_OF + dt.timedelta(days=random.randint(-20, 60)), Hours=random.choice([2, 4, 6, 8, 12]))
                row.update(DEMO_LOG[p](i))
                st = random.random()
                if st < 0.45:
                    row.update(Status="Complete", Verified="Yes",
                               Link=f"https://example.sharepoint.com/sites/DunGPS/Shared/{cd}/DEMO_{cd}_L{i + 1}.jpg")
                elif st < 0.75:
                    row.update(Status="In Progress", Verified="No")
                else:
                    row.update(Status="Not Started", Verified="No")
                ldata.append(row)
        lcols = Cols(ls)
        lf, ll = build_table(ws, wb, lcols, LH, NLOG, f"tblLog_{cd}", ldata, name_prefix=f"p{cd}_D", row_height=18)
        LL = lcols.letter
        evidence_dv(ws, f"{LL['Status']}{lf}:{LL['Status']}{ll}", f"{LL['Verified']}{lf}:{LL['Verified']}{ll}",
                    f"${LL['Link']}{lf}")
        text_cf(ws, f"{LL['Health']}{lf}:{LL['Health']}{ll}", HEALTH)
        # widths = max of both tables; hide helper columns
        wmax = {}
        for spec, cols in [(rs, rcols), (ls, lcols)]:
            for c in spec:
                Lc = cols.letter[c["key"]]
                wmax[Lc] = max(wmax.get(Lc, 0), c.get("width", 12))
        for Lc, w in wmax.items():
            ws.column_dimensions[Lc].width = w
        for k in ("Appl", "Done", "Overdue", "OpenDue"):
            ws.column_dimensions[RL[k]].hidden = True
        # ---- KPI tiles (rows 5-7)
        tl = [("COMPLETION RATE", f'=IF(SUM(p{cd}_R_Appl)=0,0,SUM(p{cd}_R_Done)/SUM(p{cd}_R_Appl))', PCT,
               f'=SUM(p{cd}_R_Done)&" of "&SUM(p{cd}_R_Appl)&" requirements done"', "2E7D32", 3),
              ("AVERAGE PROGRESS", f'=IF(SUM(p{cd}_R_Appl)=0,0,SUMPRODUCT(p{cd}_R_Appl,p{cd}_R_Prog)/SUM(p{cd}_R_Appl))', PCT,
               '="all requirements"', NAVY, 2),
              ("OVERDUE", f"=SUM(p{cd}_R_Overdue)+SUM(p{cd}_D_Overdue)", "0",
               f'=SUM(p{cd}_R_Overdue)&" req. · "&SUM(p{cd}_D_Overdue)&" log activities"', BURG, 2),
              ("EVIDENCE UPLOADED", f'=IF(SUM(p{cd}_R_Appl)=0,0,SUMPRODUCT(p{cd}_R_Appl,--(p{cd}_R_Link<>""))/SUM(p{cd}_R_Appl))', PCT,
               f'=COUNTIF(p{cd}_R_Health,"Evidence Missing")&" complete w/o evidence"', NAVY, 2),
              ("LOG ACTIVITIES DONE", f"=SUM(p{cd}_D_Done)", "0", f'="of "&SUM(p{cd}_D_Appl)&" logged"', NAVY, 2),
              ("SCHEDULED HOURS", f"=SUMIFS(p{cd}_D_Hours,p{cd}_D_Appl,1)", "0",
               f'=TEXT(SUMIFS(p{cd}_D_Hours,p{cd}_D_Appl,1,p{cd}_D_Done,0),"0")&" hrs still to do"', GRAY, 2),
              ("NEXT OPEN DUE DATE", f'=IF(MIN(p{cd}_R_OpenDue)=0,"–",MIN(p{cd}_R_OpenDue))', DATE,
               '="earliest open requirement"', GRAY, 2)]
        c = 2
        for lab, fm, fmt, sub, acc, w in tl:
            tile(ws, 5, c, lab, fm, fmt, sub, w, acc)
            c += w
        ws.row_dimensions[6].height = 28
        ws.freeze_panes = "C5"
        ws.sheet_view.zoomScale = 90

    # ======================= OVERVIEW =======================
    ws = sh[S_OV]
    title_band(ws, "Overview – all Bronze requirements", "Read-only roll-up of the pillar sheets. Update the pillar sheets, not this one.", 18,
               legend=False)
    ov = ["Ref", "Pillar", "Layer", "Requirement", "Owner", "Start", "Due", "Status", "Progress", "Evidence Link",
          "Verified?", "Health", "Missing", "Gap", "Next", "Log Activities", "Log Done", "Sched. Hours",
          "Appl", "Done", "Overdue", "DaysLate", "Key"]
    col = {h: L(2 + j) for j, h in enumerate(ov)}
    src = {"Ref": "Ref", "Requirement": "Req", "Owner": "Owner", "Start": "Start", "Due": "Due", "Status": "Status",
           "Progress": "Prog", "Evidence Link": "Link", "Verified?": "Verified", "Health": "Health", "Missing": "Missing",
           "Gap": "Gap", "Next": "Next", "Log Activities": "LogN", "Log Done": "LogDone", "Sched. Hours": "Hours",
           "Appl": "Appl", "Done": "Done", "Overdue": "Overdue"}
    RLc = Cols(req_spec("AM", "AM")).letter   # identical letters on every pillar sheet
    hdr_row(ws, 5, 2, ov, calc=True)
    ws.row_dimensions[5].height = 30
    widths = {"Ref": 8, "Pillar": 7, "Layer": 14, "Requirement": 40, "Owner": 16, "Start": 10, "Due": 10, "Status": 12,
              "Progress": 9, "Evidence Link": 20, "Verified?": 8, "Health": 15, "Missing": 18, "Gap": 22, "Next": 22}
    for h in ov:
        ws.column_dimensions[col[h]].width = widths.get(h, 7)
    r0 = 6
    for k, (p, sr) in enumerate(req_slots):
        r = r0 + k
        P = qs(sh[p])
        for h in ov:
            ref = f"{P}!${RLc[src[h]]}${sr}" if h in src else None
            if h == "Pillar":
                v = f'=IF({P}!${RLc["Req"]}${sr}="","","{p}")'
            elif h == "Layer":
                v = f'=IF({P}!${RLc["Req"]}${sr}="","","{LAYER[p]}")'
            elif h in ("Start", "Due", "Progress", "Log Activities", "Log Done", "Sched. Hours"):
                v = f'=IF(ISNUMBER({ref}),{ref},"")'
            elif h in ("Appl", "Done", "Overdue"):
                v = f"=N({ref})"
            elif h == "DaysLate":
                v = f'=IF(${col["Overdue"]}{r}=1,AsOfB-${col["Due"]}{r},"")'
            elif h == "Key":
                H_, D_, DL = f'${col["Health"]}{r}', f'${col["Due"]}{r}', f'${col["DaysLate"]}{r}'
                v = (f'=IF({H_}="Overdue",100000+N({DL}),IF({H_}="Evidence Missing",50000,IF(OR({H_}="At Risk",{H_}="Needs Verification"),'
                     f'IF(ISNUMBER({D_}),5000-MIN(4999,MAX(0,{D_}-AsOfB)),100),0)))'
                     f'+IF(OR({H_}="Overdue",{H_}="Evidence Missing",{H_}="At Risk",{H_}="Needs Verification"),ROW()/100000,0)')
            else:
                v = f"={ref}&\"\""
            c = ws[f"{col[h]}{r}"]
            c.value = v
            c.font = f(8)
            c.border = BORDER
            c.fill = fill(LGRAY)
            c.alignment = Alignment(vertical="top", wrap_text=h in ("Requirement", "Gap", "Next", "Missing"))
            if h in ("Start", "Due"):
                c.number_format = DATE
            if h == "Progress":
                c.number_format = PCT
    ovl = r0 + len(req_slots) - 1
    for h, nm in {"Ref": "Ref", "Pillar": "Pillar", "Layer": "Layer", "Requirement": "Req", "Owner": "Owner", "Due": "Due",
                  "Status": "Status", "Progress": "Prog", "Evidence Link": "Link", "Health": "Health", "Missing": "Missing",
                  "Gap": "Gap", "Next": "Next", "Appl": "Appl", "Done": "Done", "Overdue": "Overdue", "Key": "Key",
                  "Sched. Hours": "Hours"}.items():
        define(wb, f"OV_{nm}", f"{qs(ws)}!${col[h]}${r0}:${col[h]}${ovl}")
    text_cf(ws, f"{col['Health']}{r0}:{col['Health']}{ovl}", HEALTH)
    bar_cf(ws, f"{col['Progress']}{r0}:{col['Progress']}{ovl}")
    for h in ("Appl", "Done", "Overdue", "DaysLate", "Key"):
        ws.column_dimensions[col[h]].hidden = True
    ws.freeze_panes = "F6"
    ws.auto_filter.ref = f"B5:{col['Key']}{ovl}"
    protect(ws)

    build_dashboard(wb, sh, demo, gauge_rows, G0)

    # ======================= HOW TO USE =======================
    ws = sh[S_H]
    title_band(ws, "How To Use", "About 10 minutes a week per owner.", 6, legend=False)
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 110
    steps = [
        ("H", "WORKBOOK MAP"),
        ("Dashboard", "One gauge per pillar (completion rate) grouped by the 3 GPS layers, an overall gauge, pillar table, needs-attention list, owner summary and links to every pillar sheet."),
        ("Pillar sheets (9)", "LEAD, EHS, FSQ, AM, PM, VO, E&T, IM, SC. Each has (1) its official Bronze requirements with owner, start/due dates, scheduled duration, scheduled hours, a weekly schedule bar and evidence, and (2) a working log built for that pillar."),
        ("Working logs", "AM: equipment, equipment type, area/line, operators trained, open tags · PM: asset/zone/role, equipment type, Limble ref, PM compliance · LEAD: board, tier, frequency, attendance · EHS: element/program, people trained vs required · FSQ: program element, area, audit score · VO: line, loss type, baseline/target/actual · E&T: person, task/skill, qualified · IM: meeting month, ideas, projects · SC: process, gap score, gaps found/closed."),
        ("Overview", "Read-only list of all requirements pulled from the pillar sheets – handy for filtering."),
        ("Owners / Lists", "People for the dropdowns; settings, Teams folder links and dropdown values (equipment types, areas, tiers …)."),
        ("H", "WEEKLY STEPS"),
        ("1. Assign", "On each pillar sheet pick the Pillar Lead (row 4), then an Owner, Start Date and Due Date for every requirement."),
        ("2. Plan the work", "In section 2 add one row per activity: pick the Bronze Ref, describe the task, fill the pillar columns (e.g. AM: equipment + equipment type), then Scheduled Start, Due Date and Scheduled Time (hrs)."),
        ("3. Update", "Update Status, Gap / What's Missing and Next Step. % Complete is only used when a requirement has no log rows."),
        ("4. Upload evidence", "Click 'Upload ↗' (opens the pillar's Teams folder), upload the file, … → Copy link, paste it in 'Evidence Link'. 'Complete' and Verified 'Yes' only become selectable after a link is pasted."),
        ("5. Verify", "The verifier opens the evidence and sets Verified? = Yes. Only then is the activity / requirement Done."),
        ("6. Review", "Open the Dashboard in the weekly meeting. It prints landscape."),
        ("H", "HOW NUMBERS WORK"),
        ("Completion rate (gauges)", "Requirements Done ÷ requirements that count (N/A excluded). Gauge colour: red below 50%, amber 50–89%, green from 90% (change on Lists)."),
        ("Requirement progress", "With working-log rows: log rows Done ÷ log rows. Otherwise the % Complete selected."),
        ("Scheduled duration / hours", "Duration = Due − Start (days). Scheduled hours = sum of 'Scheduled Time (hrs)' of its log rows."),
        ("Schedule bar", "Blue = planned weeks (start → due), green = done, red = overdue weeks, ▲ = current week."),
        ("Health", "Done · On Track (green) · Needs Verification · At Risk – due within 14 days and under 75% (amber) · Evidence Missing · Overdue (red) · Not Started · N/A (gray, not counted)."),
        ("Bronze Ready", "The dashboard banner turns green only when every requirement that counts is Done."),
        ("Source", "Requirements are copied word for word from CHG 'GPS Pillar Bronze Level Requirements'. Pillar names, order, layers and purposes come from the CHG GPS poster."),
    ]
    r = 5
    for a_, b_ in steps:
        if a_ == "H":
            section(ws, r, 2, b_, 2)
            r += 1
            continue
        ws.cell(r, 2, a_).font = f(10, True, NAVY)
        ws.cell(r, 2).alignment = Alignment(vertical="top")
        c = ws.cell(r, 3, b_)
        c.font = f(10)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 18 if len(b_) < 110 else (32 if len(b_) < 230 else 46)
        r += 1
    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


def gauge(ws, lists_ws, row, G0, anchor, w, h):
    dc = DoughnutChart()
    dc.add_data(Reference(lists_ws, min_col=G0 + 4, max_col=G0 + 8, min_row=row), from_rows=True)
    for i, cl in enumerate(["C62828", "E0A800", "2E7D32", "D9D9D9", None]):
        pt = DataPoint(idx=i)
        if cl:
            pt.graphicalProperties.solidFill = cl
            pt.graphicalProperties.line.solidFill = "FFFFFF"
        else:
            pt.graphicalProperties.noFill = True
            pt.graphicalProperties.line.noFill = True
        dc.series[0].dPt.append(pt)
    dc.firstSliceAng = 270
    dc.holeSize = 58
    dc.legend = None
    dc.width, dc.height = w, h
    dc.roundedCorners = False
    dc.visible_cells_only = False
    gp = GraphicalProperties()
    gp.noFill = True
    gp.line.noFill = True
    dc.graphical_properties = gp
    ws.add_chart(dc, anchor)


def build_dashboard(wb, sh, demo, gauge_rows, G0):
    ws = sh[S_D]
    LS = sh[S_L]
    LSq = qs(LS)
    title_band(ws, "Bronze Level Readiness — Dashboard", "", 14, legend=False)
    ws["B3"] = '="CHG Duncanville  ·  Status as of "&TEXT(AsOfB,"dd-mmm-yyyy")' + (
        '&"  ·  DEMONSTRATION DATA"' if demo else '&"  ·  CHG GPS Pillar Bronze Level Requirements (32)"')
    ws.column_dimensions["A"].width = 2
    for i in range(2, 16):
        ws.column_dimensions[L(i)].width = 11
    ws.merge_cells("B5:N5")
    ws["B5"] = ('=IF(AND(SUM(OV_Appl)>0,SUM(OV_Done)=SUM(OV_Appl)),"✔  BRONZE READY – all "&SUM(OV_Appl)&" requirements done and verified",'
                '"NOT YET BRONZE READY  –  "&SUM(OV_Done)&" of "&SUM(OV_Appl)&" requirements done  ·  "&SUM(OV_Overdue)&" overdue  ·  "&COUNTIF(OV_Health,"Evidence Missing")&" missing evidence")')
    ws["B5"].font = f(13, True, WHITE)
    ws["B5"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[5].height = 28
    ws.conditional_formatting.add("B5:N5", FormulaRule(formula=['LEFT($B$5,1)="✔"'], fill=fill("2E7D32")))
    ws.conditional_formatting.add("B5:N5", FormulaRule(formula=['LEFT($B$5,1)<>"✔"'], fill=fill(BURG)))
    tiles = [
        ("% DONE (VERIFIED)", '=IF(SUM(OV_Appl)=0,0,SUM(OV_Done)/SUM(OV_Appl))', PCT, '=SUM(OV_Done)&" of "&SUM(OV_Appl)&" requirements"', "2E7D32", 2),
        ("EVIDENCE UPLOADED", '=IF(SUM(OV_Appl)=0,0,SUMPRODUCT(OV_Appl,--(OV_Link<>""))/SUM(OV_Appl))', PCT, '=SUMPRODUCT(OV_Appl,--(OV_Link<>""))&" have evidence"', NAVY, 2),
        ("IN PROGRESS", '=COUNTIF(OV_Health,"On Track")+COUNTIF(OV_Health,"At Risk")', "0", '=COUNTIF(OV_Health,"At Risk")&" at risk"', NAVY, 1),
        ("NEEDS VERIFICATION", '=COUNTIF(OV_Health,"Needs Verification")', "0", '="evidence in, not verified"', "C9A227", 2),
        ("EVIDENCE MISSING", '=COUNTIF(OV_Health,"Evidence Missing")', "0", '="complete, no evidence"', BURG, 2),
        ("NOT STARTED", '=COUNTIF(OV_Health,"Not Started")', "0", '="no progress"', GRAY, 1),
        ("OVERDUE", "=SUM(OV_Overdue)", "0", '="past due date"', BURG, 2),
        ("NO OWNER", '=COUNTIFS(OV_Appl,1,OV_Done,0,OV_Owner,"")', "0", '="unassigned"', BURG, 1),
    ]
    c = 2
    for lab, fm, fmt, sub, acc, w in tiles:
        tile(ws, 7, c, lab, fm, fmt, sub, w, acc)
        c += w
    ws.row_dimensions[8].height = 30
    # ---- gauges
    r0 = 11
    section(ws, r0, 2, "Completion rate by pillar  (requirements done & verified ÷ requirements that count)", 13)
    bands = [("1 · FOUNDATION – creates stable & safe conditions", ["LEAD", "EHS", "FSQ"], "Foundation"),
             ("2 · EQUIPMENT FOCUS – improves equipment health", ["AM", "PM", "VO"], "Equipment Focus"),
             ("3 · CAPABILITY & GOVERNANCE – builds capability & drives execution", ["E&T", "IM", "SC"], "Capability & Governance")]
    BAND = 11
    for bi, (lab, pils, layer) in enumerate(bands):
        top = r0 + 1 + bi * BAND
        ws.merge_cells(start_row=top, start_column=2, end_row=top, end_column=10)
        hc = ws.cell(top, 2, lab)
        hc.font = f(9, True, WHITE)
        hc.fill = fill(LAYER_COLOR[layer])
        hc.alignment = Alignment(indent=1, vertical="center")
        for k, p in enumerate(pils):
            c0 = 2 + 3 * k
            gr = gauge_rows[p]
            gauge(ws, LS, gr, G0, f"{L(c0)}{top + 1}", 6.2, 4.4)
            vr = top + 6
            ws.merge_cells(start_row=vr, start_column=c0, end_row=vr, end_column=c0 + 2)
            v = ws.cell(vr, c0, f"={LSq}!${L(G0 + 3)}${gr}")
            v.number_format = PCT
            v.font = f(16, True, NAVY)
            v.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[vr].height = 22
            ws.merge_cells(start_row=vr + 1, start_column=c0, end_row=vr + 1, end_column=c0 + 2)
            s = ws.cell(vr + 1, c0, f'=HYPERLINK("#\'{p}\'!A1","{p} – {NAME[p]}")')
            s.font = Font(name=FONT, size=9, bold=True, color=NAVY, underline="single")
            s.alignment = Alignment(horizontal="center")
            ws.merge_cells(start_row=vr + 2, start_column=c0, end_row=vr + 2, end_column=c0 + 2)
            s2 = ws.cell(vr + 2, c0, f'={LSq}!${L(G0 + 1)}${gr}&" of "&{LSq}!${L(G0 + 2)}${gr}&" done · "&COUNTIFS(OV_Pillar,"{p}",OV_Overdue,1)&" overdue"')
            s2.font = f(8, False, GRAY)
            s2.alignment = Alignment(horizontal="center")
            cell = f"{L(c0)}{vr}"
            ws.conditional_formatting.add(cell, FormulaRule(formula=[f"{cell}<set_GaugeRed"], font=Font(color="C62828", bold=True)))
            ws.conditional_formatting.add(cell, FormulaRule(formula=[f"{cell}>=set_GaugeGreen"], font=Font(color="2E7D32", bold=True)))
            ws.conditional_formatting.add(cell, FormulaRule(formula=[f"AND({cell}>=set_GaugeRed,{cell}<set_GaugeGreen)"], font=Font(color="B8860B", bold=True)))
    # overall gauge
    top = r0 + 1
    ws.merge_cells(start_row=top, start_column=11, end_row=top, end_column=14)
    hc = ws.cell(top, 11, "OVERALL BRONZE COMPLETION")
    hc.font = f(9, True, WHITE)
    hc.fill = fill(NAVY)
    hc.alignment = Alignment(horizontal="center")
    gauge(ws, LS, gauge_rows["ALL"], G0, f"K{top + 1}", 8.6, 6.0)
    vr = top + 7
    ws.merge_cells(start_row=vr, start_column=11, end_row=vr + 1, end_column=14)
    v = ws.cell(vr, 11, f"={LSq}!${L(G0 + 3)}${gauge_rows['ALL']}")
    v.number_format = PCT
    v.font = f(26, True, NAVY)
    v.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=vr + 2, start_column=11, end_row=vr + 2, end_column=14)
    s = ws.cell(vr + 2, 11, f'={LSq}!${L(G0 + 1)}${gauge_rows["ALL"]}&" of "&{LSq}!${L(G0 + 2)}${gauge_rows["ALL"]}&" requirements done & verified"')
    s.font = f(9, True, NAVY)
    s.alignment = Alignment(horizontal="center")
    lg = vr + 5
    for i, (txt, colr) in enumerate([("■ Red  below 50%", "C62828"), ("■ Amber  50–89%", "E0A800"), ("■ Green  90% +", "2E7D32")]):
        ws.cell(lg + i, 12, txt).font = f(8, True, colr)
    ws.cell(lg + 3, 12, "Thresholds on the Lists sheet").font = f(7, False, GRAY, True)
    ls0 = lg + 6
    hdr_row(ws, ls0, 11, ["GPS Layer", "", "Done", "Rate"])
    ws.merge_cells(start_row=ls0, start_column=11, end_row=ls0, end_column=12)
    for k, (lab, pils, layer) in enumerate(bands):
        r = ls0 + 1 + k
        ws.merge_cells(start_row=r, start_column=11, end_row=r, end_column=12)
        calc_cell(ws, r, 11, layer, None, True, align="left", size=8)
        dn = "+".join(f'COUNTIFS(OV_Pillar,"{p}",OV_Done,1)' for p in pils)
        ap = "+".join(f'COUNTIFS(OV_Pillar,"{p}",OV_Appl,1)' for p in pils)
        calc_cell(ws, r, 13, f'=({dn})&"/"&({ap})', None, size=8)
        calc_cell(ws, r, 14, f'=IF(({ap})=0,"",({dn})/({ap}))', PCT, True, size=8)
    bar_cf(ws, f"N{ls0 + 1}:N{ls0 + 3}", "2E7D32")
    # ---- pillar table
    rt0 = r0 + 1 + 3 * BAND + 1
    section(ws, rt0, 2, "Pillar detail", 13)
    hdr_row(ws, rt0 + 1, 2, ["Pillar", "Pillar Name", "Requirements", "Done", "In Progress", "Needs Verif. / Evid.",
                             "Not Started", "Overdue", "Completion Rate", "Avg Progress", "Pillar Status", "Next Open Due",
                             "Sched. Hours"])
    ws.row_dimensions[rt0 + 1].height = 30
    for i, p in enumerate(ORDER):
        r = rt0 + 2 + i
        P = f'"{p}"'
        calc_cell(ws, r, 2, f'=HYPERLINK("#\'{p}\'!A1","{p}")', None, True)
        calc_cell(ws, r, 3, NAME[p], None, align="left", size=8)
        calc_cell(ws, r, 4, f'=COUNTIFS(OV_Pillar,{P},OV_Appl,1)', "0")
        calc_cell(ws, r, 5, f'=COUNTIFS(OV_Pillar,{P},OV_Done,1)', "0", True)
        calc_cell(ws, r, 6, f'=COUNTIFS(OV_Pillar,{P},OV_Health,"On Track")+COUNTIFS(OV_Pillar,{P},OV_Health,"At Risk")', "0")
        calc_cell(ws, r, 7, f'=COUNTIFS(OV_Pillar,{P},OV_Health,"Needs Verification")+COUNTIFS(OV_Pillar,{P},OV_Health,"Evidence Missing")', "0")
        calc_cell(ws, r, 8, f'=COUNTIFS(OV_Pillar,{P},OV_Health,"Not Started")', "0")
        calc_cell(ws, r, 9, f'=COUNTIFS(OV_Pillar,{P},OV_Overdue,1)', "0")
        calc_cell(ws, r, 10, f'=IF(N(D{r})=0,"",E{r}/D{r})', PCT)
        calc_cell(ws, r, 11, f'=IF(N(D{r})=0,"",SUMPRODUCT(--(OV_Pillar={P}),OV_Appl,OV_Prog)/D{r})', PCT)
        calc_cell(ws, r, 12, f'=IF(N(D{r})=0,"N/A",IF(E{r}=D{r},"Complete",IF(I{r}>0,"Overdue items",IF(H{r}=D{r},"Not Started","In Progress"))))')
        calc_cell(ws, r, 13, f'=IF(MIN(p{CODE[p]}_R_OpenDue)=0,"",MIN(p{CODE[p]}_R_OpenDue))', DATE)
        calc_cell(ws, r, 14, f'=SUMIFS(OV_Hours,OV_Pillar,{P})', "0")
    rt = rt0 + 11
    calc_cell(ws, rt, 2, "TOTAL", None, True)
    calc_cell(ws, rt, 3, "All pillars", None, True, align="left")
    for c_ in range(4, 10):
        calc_cell(ws, rt, c_, f"=SUM({L(c_)}{rt0 + 2}:{L(c_)}{rt - 1})", "0", True)
    calc_cell(ws, rt, 10, f'=IF(N(D{rt})=0,"",E{rt}/D{rt})', PCT, True)
    calc_cell(ws, rt, 11, f'=IF(N(D{rt})=0,"",SUMPRODUCT(OV_Appl,OV_Prog)/D{rt})', PCT, True)
    calc_cell(ws, rt, 12, f'=IF(N(D{rt})=0,"N/A",IF(E{rt}=D{rt},"Complete",IF(I{rt}>0,"Overdue items",IF(H{rt}=D{rt},"Not Started","In Progress"))))', None, True)
    calc_cell(ws, rt, 13, f'=IF(MIN(M{rt0 + 2}:M{rt - 1})=0,"",MIN(M{rt0 + 2}:M{rt - 1}))', DATE, True)
    calc_cell(ws, rt, 14, f"=SUM(N{rt0 + 2}:N{rt - 1})", "0", True)
    for c_ in range(2, 15):
        ws.cell(rt, c_).fill = fill(LGRAY)
    bar_cf(ws, f"J{rt0 + 2}:J{rt}", "2E7D32")
    bar_cf(ws, f"K{rt0 + 2}:K{rt}")
    text_cf(ws, f"L{rt0 + 2}:L{rt}", {"Complete": "Green", "In Progress": "Yellow", "Overdue items": "Red", "Not Started": "Gray"})
    ws.conditional_formatting.add(f"I{rt0 + 2}:I{rt}", FormulaRule(formula=[f"I{rt0 + 2}>0"], font=Font(color="C00000", bold=True)))
    # ---- needs attention
    r1 = rt + 2
    section(ws, r1, 2, "Needs attention (overdue first, then missing evidence, at risk, awaiting verification)", 13)
    hdr_row(ws, r1 + 1, 2, ["Ref", "Requirement", "", "", "Owner", "", "Due Date", "Health", "What's missing", "", "", "Next step", ""])
    spans = [(3, 5), (6, 7), (10, 12), (13, 14)]
    for a, b in spans:
        ws.merge_cells(start_row=r1 + 1, start_column=a, end_row=r1 + 1, end_column=b)
    for k in range(10):
        r = r1 + 2 + k
        m = f"MATCH(LARGE(OV_Key,{k + 1}),OV_Key,0)"
        g = f"LARGE(OV_Key,{k + 1})<=0"
        for a, b in spans:
            ws.merge_cells(start_row=r, start_column=a, end_row=r, end_column=b)
        calc_cell(ws, r, 2, f'=IF({g},"",INDEX(OV_Ref,{m})&"")', None, True, size=8)
        calc_cell(ws, r, 3, f'=IF({g},"",INDEX(OV_Req,{m})&"")', None, align="left", size=8)
        calc_cell(ws, r, 6, f'=IF({g},"",IF(INDEX(OV_Owner,{m})="","(no owner)",INDEX(OV_Owner,{m})&""))', None, align="left", size=8)
        calc_cell(ws, r, 8, f'=IF({g},"",IF(INDEX(OV_Due,{m})="","",INDEX(OV_Due,{m})))', DATE, size=8)
        calc_cell(ws, r, 9, f'=IF({g},"",INDEX(OV_Health,{m})&"")', None, size=8)
        calc_cell(ws, r, 10, f'=IF({g},"",INDEX(OV_Missing,{m})&IF(INDEX(OV_Gap,{m})="",""," "&INDEX(OV_Gap,{m})))', None, align="left", size=8)
        calc_cell(ws, r, 13, f'=IF({g},"",INDEX(OV_Next,{m})&"")', None, align="left", size=8)
        ws.row_dimensions[r].height = 24
    text_cf(ws, f"I{r1 + 2}:I{r1 + 11}", HEALTH)
    ws.conditional_formatting.add(f"F{r1 + 2}:F{r1 + 11}", FormulaRule(formula=[f'F{r1 + 2}="(no owner)"'], font=Font(color="C00000", bold=True)))
    # ---- owners + pillar links
    r2 = r1 + 13
    section(ws, r2, 2, "By owner", 6)
    hdr_row(ws, r2 + 1, 2, ["Owner", "", "Assigned", "Done", "Overdue", "Completion"])
    ws.merge_cells(start_row=r2 + 1, start_column=2, end_row=r2 + 1, end_column=3)
    for k in range(13):
        r = r2 + 2 + k
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        if k < 12:
            calc_cell(ws, r, 2, f'=INDEX(bl_OwnerAll,{k + 1})&""', None, align="left", size=8)
            O, guard = f"$B{r}", f'$B{r}=""'
        else:
            calc_cell(ws, r, 2, "(no owner)", None, align="left", size=8)
            O, guard = '""', "FALSE"
        calc_cell(ws, r, 4, f'=IF({guard},"",COUNTIFS(OV_Owner,{O},OV_Appl,1))', "0", size=8)
        calc_cell(ws, r, 5, f'=IF({guard},"",COUNTIFS(OV_Owner,{O},OV_Done,1))', "0", size=8)
        calc_cell(ws, r, 6, f'=IF({guard},"",COUNTIFS(OV_Owner,{O},OV_Overdue,1))', "0", size=8)
        calc_cell(ws, r, 7, f'=IF(N(D{r})=0,"",E{r}/D{r})', PCT, size=8)
    bar_cf(ws, f"G{r2 + 2}:G{r2 + 14}", "2E7D32")
    ws.conditional_formatting.add(f"F{r2 + 2}:F{r2 + 14}", FormulaRule(formula=[f"N(F{r2 + 2})>0"], font=Font(color="C00000", bold=True)))
    section(ws, r2, 9, "Go to pillar sheet", 6)
    for i, p in enumerate(ORDER):
        r = r2 + 2 + i
        ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=12)
        c = ws.cell(r, 9, f'=HYPERLINK("#\'{p}\'!A1","▸ {p} – {NAME[p]}  ({LAYER[p]})")')
        c.font = Font(name=FONT, size=9, color="1F4E79", underline="single")
    landscape(ws, f"A1:N{r2 + 15}", one_page=False)
    ws.page_setup.fitToHeight = 2
    protect(ws)


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "."
    build(False, f"{out}/CHG_Bronze_Level_Tracker.xlsx")
    build(True, f"{out}/CHG_Bronze_Level_Tracker_DEMO.xlsx")
    print("built")
