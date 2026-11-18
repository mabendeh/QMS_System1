"""Builds CHG_DunGPS_Deployment_Master.xlsx (production) and a DEMO version for testing.

Usage: python build_master.py <out_dir>
"""
import sys
import datetime as dt
import random
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.chart import BarChart, LineChart, DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from xlcommon import *  # noqa

S_EXEC = "Executive Dashboard"
S_TRK = "Pillar Deployment Tracker"
S_PDM = "Process Deployment Matrix"
S_GDB = "Gap Analysis Dashboard"
S_GAP = "Gap Register"
S_OWN = "Owner Directory"
S_BRZ = "Bronze Star Readiness"
S_VO = "VO & Tiger Team Projects"
S_DOC = "Document & Evidence Register"
S_MR = "Management Review"
S_CFG = "Configuration & Dropdown Lists"
S_INS = "Instructions & Data Definitions"
S_CALC = "Calc Engine"

N_TRK, N_GAP, N_OWN, N_BRZ, N_VO, N_DOC, N_PDM = 500, 300, 200, 60, 100, 300, 80
NPIL = 12  # pillar capacity on dashboards
SOURCE_DOC = "GPS_Pillar_Bronze_Level_Requirements.xlsx (CHG, uploaded 18-Nov-26)"

# ---- Official Bronze criteria: imported verbatim from the CHG source document ----
BRONZE = {
    "AM": ["Establish CILs, Centerlines, and defect handling.  Operator training",
           "Deep clean on selected asset(s)", "Creation of LOTO and Safety Maps"],
    "PM": ["Part Room 2S", "CMMS-Limble deployment or proficiency with current system",
           "Maintenance org structure (Maint Mgr, Planner/scheduler, Parts clerk), except for approved exceptions"],
    "FSQ": ["Management Commitment", "Key Performance Indicators", "Good Manufacturing Practices",
            "Control of Non-Conforming Product", "Cleaning & Sanitation", "Foreign Material Control",
            "Allergen Management", "5S for Lab & Sanitation Areas", "Process Control & Capability (SPC)"],
    "EHS": ["Implement Safestart", "Establish 21 Elements"],
    "LEAD": ["FMDS refresh"],
    "E&T": ["Perform one standard work Instruction per Supervisor", "Skill Matrix development for Operators",
            "Develop leader standard work for plant staff"],
    "VO": ["Basic Problem solving - 5Y & Fishbone including countermeasures and action",
           "Process Flow Map - 1 line", "OEE - Redzone APQ understanding and use",
           "SAC - Team leaders to pull reports to analyize and understand losses", "SPIOC - complete 1",
           "SMED on one asset", "Executing one FIT / Tiger team to eavaluate loss."],
    "IM": ["Establish monthly Plant  Steering Committee to review Ideas, Opportunities, Projects"],
    "SC": ["PONC related process", "SC process governance and gap analysis",
           "Procurement process governance and gap analysis"],
}
# CI-recommended evidence (NOT official CHG evidence requirements) – shown in a clearly labelled column
SUGGEST = {
    "AM": ["CIL sheets, centerline standards, defect tag log, operator training sign-off per asset",
           "Before/after photos, deep-clean checklist, abnormality list with tag closure",
           "Approved LOTO procedures per asset, posted safety maps, LOTO audit"],
    "PM": ["2S audit (Sort/Set) of parts room, red-tag log, labelled locations photos",
           "Limble usage report (WO compliance, PM schedule), user proficiency sign-off",
           "Approved org chart, role descriptions, exception approval (if any)"],
    "FSQ": ["Signed FS&Q policy / leadership review minutes", "FSQ KPI board & monthly trend report",
            "GMP audit results and corrective actions", "Hold/release procedure, NCP log, disposition records",
            "Master sanitation schedule, SSOPs, verification records (ATP / swab)",
            "FM control program, metal detection / x-ray verification records",
            "Allergen matrix, changeover validation, label verification records",
            "5S audit scores for lab & sanitation areas (baseline + current)",
            "Control charts / Cpk studies for critical parameters"],
    "EHS": ["SafeStart training roster & completion %, observation records",
            "21-element assessment with owner per element and gap plan"],
    "LEAD": ["Refreshed FMDS board standard, tier meeting attendance & audit"],
    "E&T": ["Approved standard work instruction per supervisor (list + documents)",
            "Operator skill matrix per area with training plan",
            "Leader standard work documents and adherence audits"],
    "VO": ["Completed 5Y / fishbone A3s with countermeasures and actions",
           "Approved process flow map for selected line", "Redzone APQ reports & training sign-off",
           "Team leader loss reports (SAC) with actions", "Completed SPIOC document",
           "SMED study: baseline vs. new changeover time, standard",
           "FIT / Tiger Team charter, loss analysis, results"],
    "IM": ["Steering committee charter, monthly minutes, idea/project pipeline"],
    "SC": ["PONC process map/procedure and tracking", "SC governance gap assessment and action plan",
           "Procurement governance gap assessment and action plan"],
}

PILLARS = [  # code, name (expansion = CI assumption to confirm), Teams folder, in Bronze source
    ("AM", "Autonomous Maintenance", "AM", "Yes"),
    ("PM", "Planned Maintenance", "PM", "Yes"),
    ("FSQ", "Food Safety & Quality", "FSQ", "Yes"),
    ("EHS", "Environment, Health & Safety", "EHS", "Yes"),
    ("LEAD", "Leadership", "LEAD", "Yes"),
    ("E&T", "Education & Training", "E&T", "Yes"),
    ("VO", "Value Optimization", "VO", "Yes"),
    ("IM", "IM – full name to be confirmed", "IM", "Yes"),
    ("SC", "Supply Chain", "SC", "Yes"),
    ("EEM", "Early Equipment Management", "EEM", "No – optional"),
]
LISTS = [  # header, key, values, system-controlled?
    ("Deployment Phase", "Phase", ["0 - Launch & Readiness", "1 - Bronze", "2 - Silver", "3 - Gold", "4 - Sustain"], False),
    ("Task Status", "Status", ["Not Started", "In Progress", "On Hold", "Complete"], True),
    ("Verification Status", "Verif", ["Not Submitted", "Submitted - Pending Review", "Verified", "Rejected - Rework"], True),
    ("Applicability", "Applic", ["Applicable", "N/A - Requested", "N/A - Approved"], True),
    ("Yes / No", "YesNo", ["Yes", "No"], True),
    ("Shift", "Shift", ["A", "B", "C", "All"], False),
    ("Process / Production Area", "Area", ["Pre-weigh", "Mixing", "Dough Processing / Forming", "Gemini Equipment",
                                           "Nugget Line", "Rolled Dough Lines", "Baking", "Cooling / Freezing",
                                           "Packaging", "Sanitation", "Maintenance / Parts Room", "Laboratory / QA",
                                           "Warehouse / Shipping", "Site-wide"], False),
    ("Severity", "Severity", ["Critical", "High", "Medium", "Low"], True),
    ("Severity Score", "SevScore", [4, 3, 2, 1], False),
    ("Gap Status", "GapStatus", ["Open", "In Progress", "Pending Verification", "Closed - Verified", "Cancelled"], True),
    ("Gap / Barrier Category", "GapCat", ["Standard / procedure missing", "Training / skills", "Resources / time",
                                          "Equipment condition", "Data / system (CMMS, Redzone)",
                                          "Leadership / accountability", "Documentation / evidence",
                                          "Supplier / material", "Safety / LOTO", "Food safety / GMP",
                                          "Communication", "Other"], False),
    ("Evidence Type", "EvType", ["SOP", "OPL", "Standard Work Instruction", "CIL Sheet", "Centerline Standard",
                                 "Defect Tag / Log", "LOTO Procedure", "Safety Map", "Training Record",
                                 "Skill Matrix", "Audit / Checklist", "KPI Report", "Meeting Minutes",
                                 "Project Charter", "Photo (Before / After)", "Process Map", "Other"], False),
    ("Document Status", "DocStatus", ["Draft", "In Review", "Approved", "Verified", "Obsolete"], True),
    ("Project Status", "ProjStatus", ["Planning — Baseline Pending", "Planning", "Active", "On Hold", "Implemented",
                                      "Sustaining", "Closed", "Cancelled"], True),
    ("Project Type", "ProjType", ["Tiger Team / FIT", "Kaizen Event", "SMED", "SPIOC", "DMAIC (Six Sigma)", "5S",
                                  "Process Flow / Layout", "Capex", "Just-Do-It"], False),
    ("Finance Approval", "FinApp", ["Not Submitted", "Submitted", "Approved", "Rejected"], True),
    ("Sustainability Check", "Sustain", ["Not Due", "30-Day Verified", "90-Day Verified", "Failed - Reopened"], True),
    ("Matrix Status", "MxStatus", ["Not Assessed", "N/A", "Not Started", "In Progress", "Implemented", "Verified"], True),
    ("Active / Inactive", "Active", ["Active", "Inactive"], True),
    ("Charter Status", "Charter", ["Not Started", "Draft", "Approved"], True),
    ("DMAIC Phase", "DMAIC", ["Define", "Measure", "Analyze", "Improve", "Control"], False),
    ("Department", "Dept", ["Plant Leadership", "Production", "Maintenance", "Quality / FSQ", "EHS", "Sanitation",
                            "Engineering", "Supply Chain", "Finance", "HR / Training"], False),
    ("Role", "Role", ["Plant Leadership", "Pillar Lead", "Pillar Member", "Area Owner", "Verifier / Auditor",
                      "Supervisor", "Team Leader", "Finance"], False),
    ("Bronze Verification", "BrzVerif", ["Not Assessed", "Self-Assessed", "Submitted - Pending Review", "Verified",
                                         "Rejected - Rework"], True),
    ("Bronze Applicability", "BrzApplic", ["Yes", "N/A - Approved by CHG"], True),
]
ELEMENTS = ["Safety Mapping", "CILT Deployment", "Centerline Standards", "Defect ID & Handling",
            "Operator Training", "Standard Work", "Loss Identification", "Corrective Actions",
            "Verification of Implementation"]

CFG_LIST_ROW0 = 6
CFG_LIST_ROWN = 55


def build(demo, path):
    random.seed(42)
    AS_OF = dt.date(2026, 11, 18)
    wb = Workbook()
    wb.remove(wb.active)
    sh = {}
    for n in [S_EXEC, S_TRK, S_PDM, S_GDB, S_GAP, S_OWN, S_BRZ, S_VO, S_DOC, S_MR, S_CFG, S_INS, S_CALC]:
        sh[n] = wb.create_sheet(n)
    tabc = {S_EXEC: NAVY, S_MR: NAVY, S_GDB: NAVY, S_TRK: BURG, S_GAP: BURG, S_BRZ: BURG, S_VO: BURG,
            S_DOC: BURG, S_PDM: BURG, S_OWN: "7F7F7F", S_CFG: "7F7F7F", S_INS: "7F7F7F", S_CALC: "BFBFBF"}
    for k, v in tabc.items():
        sh[k].sheet_properties.tabColor = v
    CALC = qs(sh[S_CALC])
    tagline = ("DEMONSTRATION DATA – fictitious records for testing only. Do not use for decisions."
               if demo else "C.H. Guenther & Son — Duncanville · DunGPS Pillar Deployment Management System")

    # ================= CONFIGURATION =================
    ws = sh[S_CFG]
    title_band(ws, "Configuration & Dropdown Lists",
               "Administrator sheet. Thresholds and lists drive every dropdown and dashboard. "
               "Gray-header lists are SYSTEM values used in formulas – add values only at the end, never rename.", 40)
    section(ws, 5, 2, "Parameters (edit the white cells)", 3)
    params = [
        ("cfg_SiteName", "Site name", "C.H. Guenther & Son — Duncanville", None, "Shown on dashboards."),
        ("cfg_AsOfOverride", "Status date override", AS_OF if demo else None, DATE,
         "Leave blank to use TODAY(). Enter a date to freeze reporting (e.g. for a meeting pack)."),
        ("AsOf", "ƒ Effective status date", "=IF(ISNUMBER(cfg_AsOfOverride),cfg_AsOfOverride,TODAY())", DATE,
         "All overdue, aging and planned-% logic uses this date."),
        ("cfg_UnverifiedCap", "Max credit before verification", 0.8, PCT,
         "A task marked Complete earns at most this credit until evidence is Verified (prevents 'self-declared' 100%)."),
        ("cfg_UseWeights", "Use task weights?", "No", None,
         "No = every required task counts equally (default formula). Yes = use the Weight column."),
        ("cfg_GreenTol", "Green tolerance (pts behind plan)", 0.05, PCT,
         "Completion credit within this many points of planned % = Green."),
        ("cfg_YellowTol", "Yellow tolerance (pts behind plan)", 0.15, PCT,
         "Behind plan by more than Green tolerance but within this = Yellow; beyond = Red."),
        ("cfg_RedOverdueShare", "Red if overdue share ≥", 0.10, PCT,
         "Pillar/phase turns Red when this share of in-scope tasks is overdue. Any overdue caps Green at Yellow."),
        ("cfg_DueSoonDays", "Due-soon window (days)", 14, "0", "Items due within this window are flagged at-risk."),
        ("cfg_EscalationDays", "Escalate when overdue > (days)", 14, "0",
         "Overdue tasks beyond this count as 'Critical issues requiring leadership intervention'."),
        ("cfg_Age1", "Gap aging bucket 1 (days)", 30, "0", "Aging buckets for open gaps."),
        ("cfg_Age2", "Gap aging bucket 2 (days)", 60, "0", ""),
        ("cfg_Age3", "Gap aging bucket 3 (days)", 90, "0", ""),
        ("cfg_TrendDays", "Trend period length (days)", 7, "0", "Planned vs actual trend: 26 periods ending at status date."),
        ("cfg_BronzeYellow", "Bronze pillar readiness Yellow ≥", 0.75, PCT, "100% verified = Green."),
        ("cfg_MatrixGreen", "Process matrix Green ≥", 0.80, PCT, "Area deployment score thresholds."),
        ("cfg_MatrixYellow", "Process matrix Yellow ≥", 0.50, PCT, ""),
        ("cfg_scIP", "Matrix score: In Progress", 0.25, PCT, "Not Started = 0%. Not Assessed / N/A excluded."),
        ("cfg_scImpl", "Matrix score: Implemented", 0.75, PCT, "Implemented but not yet verified."),
        ("cfg_scVer", "Matrix score: Verified", 1.0, PCT, ""),
        ("cfg_DocDueSoon", "Document review due-soon (days)", 30, "0", ""),
        ("cfg_DocFreq", "Default review frequency (months)", 12, "0", "Used when a document has no frequency."),
        ("cfg_TeamsRoot", "Teams evidence root", "Teams → DunGPS → General → Shared", None,
         "Standard folder path. Pillar sub-folders are set in the Pillar list."),
        ("cfg_SharePointBase", "SharePoint site URL (optional)", None, None,
         "Paste the DunGPS SharePoint/Teams 'Shared' folder URL when available."),
    ]
    r = 6
    for name, label, val, fmt, note in params:
        ws.cell(r, 2, label).font = f(9, True if label.startswith("ƒ") else False)
        c = ws.cell(r, 3, val)
        c.border = BORDER
        c.font = f(9, False, "0000FF" if not label.startswith("ƒ") else "000000")
        c.fill = fill(LGRAY if label.startswith("ƒ") else WHITE)
        c.alignment = Alignment(horizontal="center")
        if fmt:
            c.number_format = fmt
        ws.cell(r, 4, note).font = f(8, False, GRAY, True)
        define(wb, name, f"{qs(ws)}!$C${r}")
        r += 1
    dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("C10")
    ws.column_dimensions["B"].width = 32
    ws.column_dimensions["C"].width = 30
    ws.column_dimensions["D"].width = 60

    # Pillar table
    pc = 6  # col F
    section(ws, 4, pc, "DunGPS Pillars (editable – max 12)", 6)
    hdr_row(ws, 5, pc, ["Pillar Code", "Pillar Name", "Pillar Lead", "Teams Sub-folder", "Folder URL (optional)",
                        "In Bronze Source?"])
    for i in range(CFG_LIST_ROWN - CFG_LIST_ROW0 + 1):
        rr = CFG_LIST_ROW0 + i
        vals = PILLARS[i] if i < len(PILLARS) else (None, None, None, None)
        for j, v in enumerate([vals[0], vals[1], None, vals[2], None, vals[3]]):
            c = ws.cell(rr, pc + j, v)
            c.font = f(9)
            c.border = BORDER
    for j, w in enumerate([10, 30, 22, 14, 30, 14]):
        ws.column_dimensions[L(pc + j)].width = w
    ws.cell(CFG_LIST_ROW0 + 7, pc + 1).comment = Comment(
        "'IM' is listed in the CHG Bronze source without a full name. Confirm the official pillar name.", "DunGPS")
    ws.cell(5, pc + 1).comment = Comment("Pillar name expansions are CI assumptions based on standard TPM/GPS "
                                         "naming. Confirm against CHG documentation.", "DunGPS")
    rng = lambda col: f"{qs(ws)}!${L(col)}${CFG_LIST_ROW0}:${L(col)}${CFG_LIST_ROWN}"
    define(wb, "cfg_PillarCode", rng(pc))
    define(wb, "cfg_PillarName", rng(pc + 1))
    define(wb, "cfg_PillarFolder", rng(pc + 3))
    dyn_list(wb, "lst_Pillars", ws, L(pc), CFG_LIST_ROW0, CFG_LIST_ROWN)
    lead_dv = DataValidation(type="list", formula1="=lst_Owners", allow_blank=True)
    ws.add_data_validation(lead_dv)
    lead_dv.add(f"{L(pc + 2)}{CFG_LIST_ROW0}:{L(pc + 2)}{CFG_LIST_ROWN}")

    col = pc + 7
    listcol = {}
    for header, key, vals, system in LISTS:
        hdr_row(ws, 5, col, [header], calc=system)
        ws.column_dimensions[L(col)].width = max(14, min(30, max(len(str(v)) for v in vals) + 2))
        for i in range(CFG_LIST_ROWN - CFG_LIST_ROW0 + 1):
            c = ws.cell(CFG_LIST_ROW0 + i, col, vals[i] if i < len(vals) else None)
            c.font = f(9)
            c.border = BORDER
        listcol[key] = L(col)
        if key == "SevScore":
            define(wb, "lst_SevScore", f"{qs(ws)}!${L(col)}${CFG_LIST_ROW0}:${L(col)}${CFG_LIST_ROW0 + 3}")
        else:
            dyn_list(wb, f"lst_{key}", ws, L(col), CFG_LIST_ROW0, CFG_LIST_ROWN)
            define(wb, f"cfg_{key}", f"{qs(ws)}!${L(col)}${CFG_LIST_ROW0}:${L(col)}${CFG_LIST_ROWN}")
        col += 1
    define(wb, "lst_SevFixed", f"{qs(ws)}!${listcol['Severity']}${CFG_LIST_ROW0}:${listcol['Severity']}${CFG_LIST_ROW0 + 3}")
    ws.freeze_panes = "B6"

    # ================= OWNER DIRECTORY =================
    ws = sh[S_OWN]
    title_band(ws, "Owner Directory",
               "Single source for every owner dropdown. Add people in the next empty row – dropdowns update "
               "automatically. Only 'Active' people appear in dropdowns. No names are pre-loaded: CHG to confirm.", 16)
    ospec = [
        dict(key="Name", header="Employee Name", width=26),
        dict(key="EmpID", header="Employee ID", width=11),
        dict(key="Dept", header="Department", width=18, dv="lst_Dept"),
        dict(key="Position", header="Position", width=22),
        dict(key="Shift", header="Shift (A/B/C/All)", width=10, dv="lst_Shift"),
        dict(key="PillarResp", header="Pillar Responsibility", width=14, dv="lst_PillarResp"),
        dict(key="Role", header="DunGPS Role", width=16, dv="lst_Role"),
        dict(key="Active", header="Active / Inactive", width=11, dv="lst_Active", default="Active"),
        dict(key="Email", header="Email Address", width=28),
    ]
    ocols = Cols(ospec)
    nl, al = ocols.letter["Name"], ocols.letter["Active"]
    ospec += [
        dict(key="Seq", header="Dropdown Seq", width=9, calc=True, fmt="0",
             formula=f'=IF(AND({{Name}}<>"",{{Active}}="Active"),COUNTIFS(${nl}$6:{{Name}},"?*",${al}$6:{{Active}},"Active"),"")'),
        dict(key="Tasks", header="Open Tasks", width=8, calc=True, fmt="0",
             formula='=IF({Name}="","",SUMPRODUCT(--(T_Owner={Name}),T_InScope,1-T_Verified))'),
        dict(key="OverT", header="Overdue Tasks", width=8, calc=True, fmt="0",
             formula='=IF({Name}="","",SUMPRODUCT(--(T_Owner={Name}),T_InScope,T_Overdue))'),
        dict(key="Gaps", header="Open Gaps", width=8, calc=True, fmt="0",
             formula='=IF({Name}="","",SUMPRODUCT(--(G_Owner={Name}),G_Open))'),
        dict(key="Check", header="Check", width=18, calc=True,
             formula='=IF({Name}="","",IF(COUNTIF(OD_Name,{Name})>1,"Duplicate name; ","")&IF({Email}="","No email; ","")&IF(AND({Active}="Inactive",N({Tasks})+N({Gaps})>0),"Inactive w/ open items",""))'),
    ]
    ocols = Cols(ospec)
    odata = []
    if demo:
        roles = [("Plant Leadership", "Plant Manager", "Plant Leadership", "All", "Site / Multiple"),
                 ("Plant Leadership", "Operations Manager", "Production", "All", "Site / Multiple"),
                 ("Pillar Lead", "AM Pillar Lead", "Production", "A", "AM"),
                 ("Pillar Lead", "Maintenance Manager", "Maintenance", "All", "PM"),
                 ("Pillar Lead", "FSQ Manager", "Quality / FSQ", "All", "FSQ"),
                 ("Pillar Lead", "EHS Manager", "EHS", "All", "EHS"),
                 ("Pillar Lead", "Training Coordinator", "HR / Training", "All", "E&T"),
                 ("Pillar Lead", "CI Engineer", "Engineering", "All", "VO"),
                 ("Pillar Lead", "Supply Chain Manager", "Supply Chain", "All", "SC"),
                 ("Supervisor", "Production Supervisor", "Production", "A", "AM"),
                 ("Supervisor", "Production Supervisor", "Production", "B", "AM"),
                 ("Supervisor", "Production Supervisor", "Production", "C", "VO"),
                 ("Verifier / Auditor", "QA Auditor", "Quality / FSQ", "All", "FSQ"),
                 ("Finance", "Plant Controller", "Finance", "All", "VO"),
                 ("Supervisor", "Sanitation Supervisor", "Sanitation", "C", "FSQ")]
        for i, (role, pos, dept, shift, pr) in enumerate(roles, 1):
            odata.append(dict(Name=f"DEMO Owner {i:02d}", EmpID=f"D{i:03d}", Dept=dept, Position=pos,
                              Shift=shift, PillarResp=pr, Role=role,
                              Active="Inactive" if i == 15 else "Active",
                              Email=f"demo.owner{i:02d}@example.com" if i != 12 else None))
    build_table(ws, wb, ocols, 5, N_OWN, "tblOwners", odata, name_prefix="OD", freeze_col="C")
    # Owner list helpers live in Calc Engine (built below)

    # ================= TRACKER =================
    ws = sh[S_TRK]
    title_band(ws, "Pillar Deployment Tracker", tagline + "  ·  One row per task/step. Verified completion "
               "requires Status = Complete + Evidence link + Verification = Verified.", 48)
    tspec = [
        dict(key="ID", header="Task ID", width=10),
        dict(key="Pillar", header="Pillar", width=8, dv="lst_Pillars"),
        dict(key="Phase", header="Deployment Phase", width=13, dv="lst_Phase"),
        dict(key="Req", header="Step / Requirement", width=42, wrap=True),
        dict(key="BronzeRef", header="Bronze Req ID", width=10, dv="lst_BronzeIDs",
             note="Links the task to an official Bronze Star criterion (roll-up to Bronze Star Readiness)."),
        dict(key="Required", header="Required?", width=9, dv="lst_YesNo", default="Yes",
             note="Only required tasks count in completion %. Optional tasks are shown separately."),
        dict(key="Applic", header="Applicability", width=13, dv="lst_Applic", default="Applicable",
             note="'N/A - Approved' removes the task from denominators (requires an Approver). "
                  "'N/A - Requested' still counts until approved."),
        dict(key="Weight", header="Weight", width=7, dv="num", fmt="0.0", default=1),
        dict(key="Area", header="Process / Area", width=16, dv="lst_Area"),
        dict(key="Shift", header="Shift", width=6, dv="lst_Shift"),
        dict(key="Owner", header="Primary Owner", width=18, dv="lst_Owners"),
        dict(key="SupOwner", header="Supporting Owner", width=18, dv="lst_Owners"),
        dict(key="Team", header="Supporting Team", width=18, wrap=True),
        dict(key="Verifier", header="Approver / Verifier", width=18, dv="lst_Owners"),
        dict(key="Escal", header="Escalation Owner", width=18, dv="lst_Owners"),
        dict(key="Start", header="Start Date", width=10, dv="date", fmt=DATE),
        dict(key="Target", header="Target Date", width=10, dv="date", fmt=DATE),
        dict(key="Actual", header="Actual Completion", width=10, dv="date", fmt=DATE),
        dict(key="Status", header="Status", width=12, dv="lst_Status", default="Not Started"),
        dict(key="Progress", header="Work Progress %", width=9, dv="pct", fmt=PCT,
             note="Owner-reported progress (0–100%). Capped at the 'Max credit before verification' for completion credit."),
        dict(key="EvidReq", header="Evidence Required", width=30, wrap=True),
        dict(key="EvidLink", header="Evidence Link (Teams / SharePoint URL)", width=24),
        dict(key="EvidOpen", header="Open", width=8, calc=True,
             formula='=IF({EvidLink}="","",HYPERLINK({EvidLink},"Open ↗"))'),
        dict(key="VerStatus", header="Verification Status", width=15, dv="lst_Verif", default="Not Submitted"),
        dict(key="VerDate", header="Verification Date", width=10, dv="date", fmt=DATE),
        dict(key="BarrierCat", header="Barrier Category", width=16, dv="lst_GapCat"),
        dict(key="Gap", header="Gap / Barrier", width=26, wrap=True),
        dict(key="Next", header="Next Action", width=26, wrap=True),
        dict(key="Updated", header="Last Updated", width=10, dv="date", fmt=DATE),
        # ---- calculated ----
        dict(key="InScope", header="In Scope", width=6, calc=True, fmt="0",
             formula='=IF(AND({Pillar}<>"",{Required}<>"No",{Applic}<>"N/A - Approved"),1,0)'),
        dict(key="W", header="Eff. Weight", width=6, calc=True, fmt="0.0",
             formula='=IF({InScope}=1,IF(cfg_UseWeights="Yes",IF(ISNUMBER({Weight}),{Weight},1),1),0)'),
        dict(key="Prog", header="Reported Progress", width=8, calc=True, fmt=PCT,
             formula='=IF({Pillar}="","",IF({Applic}="N/A - Approved","N/A",IF({Status}="Complete",1,MIN(1,N({Progress})))))'),
        dict(key="Verified", header="Verified (1/0)", width=7, calc=True, fmt="0",
             formula='=IF(AND({Pillar}<>"",{Applic}<>"N/A - Approved",{Status}="Complete",{VerStatus}="Verified",{EvidLink}<>""),1,0)'),
        dict(key="Credit", header="Completion %", width=8, calc=True, fmt=PCT,
             formula='=IF({Pillar}="","",IF({Applic}="N/A - Approved","N/A",IF({Verified}=1,1,IF({Status}="Complete",cfg_UnverifiedCap,MIN(cfg_UnverifiedCap,N({Progress}))))))'),
        dict(key="DoneDate", header="Verified On", width=9, calc=True, fmt=DATE,
             formula='=IF({Verified}=1,IF(ISNUMBER({VerDate}),{VerDate},IF(ISNUMBER({Actual}),{Actual},AsOf)),"")'),
        dict(key="EvSub", header="Evidence Submitted", width=8, calc=True, fmt="0",
             formula='=IF(AND({Pillar}<>"",{EvidLink}<>""),1,0)'),
        dict(key="EvVer", header="Evidence Verified", width=8, calc=True, fmt="0",
             formula='=IF(AND({EvSub}=1,{VerStatus}="Verified"),1,0)'),
        dict(key="Planned", header="Planned % to Date", width=8, calc=True, fmt=PCT,
             formula='=IF({Pillar}="","",IF({Applic}="N/A - Approved","N/A",IF(NOT(ISNUMBER({Target})),0,IF(NOT(ISNUMBER({Start})),IF(AsOf>={Target},1,0),MAX(0,MIN(1,(AsOf-{Start})/MAX(1,{Target}-{Start})))))))'),
        dict(key="Overdue", header="Overdue (1/0)", width=7, calc=True, fmt="0",
             formula='=IF(AND({Pillar}<>"",{Applic}<>"N/A - Approved",{Verified}=0,ISNUMBER({Target})),IF({Target}<AsOf,1,0),0)'),
        dict(key="Days", header="Days Late (+) / To Due (−)", width=9, calc=True, fmt="0;-0;0",
             formula='=IF(OR({Pillar}="",{Verified}=1,{Applic}="N/A - Approved",NOT(ISNUMBER({Target}))),"",AsOf-{Target})'),
        dict(key="DashStatus", header="Dashboard Status", width=16, calc=True,
             formula='=IF({Pillar}="","",IF({Applic}="N/A - Approved","N/A",IF({Verified}=1,"Verified Complete",IF({Overdue}=1,"Overdue",IF({Status}="Complete","Complete - Pending Verification",IF(OR({Status}="In Progress",{Status}="On Hold",N({Progress})>0),"In Progress","Not Started"))))))'),
        dict(key="RAG", header="RAG", width=8, calc=True,
             formula='=IF({Pillar}="","",IF({Applic}="N/A - Approved","N/A",IF({Verified}=1,"Green",IF({Overdue}=1,"Red",IF(OR({VerStatus}="Rejected - Rework",{Status}="On Hold"),"Yellow",IF(AND({DashStatus}="Not Started",N({Planned})=0),"Gray",IF(N({Credit})>=N({Planned})-cfg_GreenTol,"Green",IF(N({Credit})>=N({Planned})-cfg_YellowTol,"Yellow","Red"))))))))'),
        dict(key="Issues", header="Data Quality Flags", width=30, calc=True, wrap=True,
             formula=('=IF({Pillar}="","",IF({ID}="","Missing Task ID; ","")&IF(AND({ID}<>"",COUNTIF(T_ID,{ID})>1),"Duplicate Task ID; ","")'
                      '&IF({Owner}="","Missing owner; ",IF(COUNTIFS(OD_Name,{Owner},OD_Active,"Active")=0,"Owner not active in directory; ",""))'
                      '&IF(AND({Applic}<>"N/A - Approved",NOT(ISNUMBER({Target}))),"Missing target date; ","")'
                      '&IF(AND(ISNUMBER({Start}),ISNUMBER({Target})),IF({Target}<{Start},"Target before start; ",""),"")'
                      '&IF(AND({Status}="Complete",NOT(ISNUMBER({Actual}))),"Complete w/o actual date; ","")'
                      '&IF(AND({Status}="Complete",{EvidLink}=""),"Complete w/o evidence; ","")'
                      '&IF(AND({VerStatus}="Verified",{EvidLink}=""),"Verified w/o evidence link; ","")'
                      '&IF(AND({VerStatus}="Verified",{Status}<>"Complete"),"Verified but not Complete; ","")'
                      '&IF(AND({VerStatus}="Verified",{Verifier}=""),"Verified w/o verifier; ","")'
                      '&IF(AND({Applic}="N/A - Approved",{Verifier}=""),"N/A w/o approver; ","")'
                      '&IF({Applic}="N/A - Requested","N/A pending approval; ","")'
                      '&IF(ISNUMBER({Actual}),IF({Actual}>AsOf,"Actual date in future; ",""),"")'
                      '&IF(AND({Status}="Not Started",N({Progress})>0),"Progress >0 but Not Started; ","")'
                      '&IF({BronzeRef}<>"",IF(COUNTIF(B_ReqID,{BronzeRef})=0,"Bronze ID not found; ",""),""))')),
        dict(key="Filter", header="MR Filter", width=6, calc=True, fmt="0",
             formula='=IF({Pillar}="",0,IF(AND(OR(f_Pillar="(All)",f_Pillar="",{Pillar}=f_Pillar),OR(f_Owner="(All)",f_Owner="",{Owner}=f_Owner),OR(f_Area="(All)",f_Area="",{Area}=f_Area),OR(f_Shift="(All)",f_Shift="",{Shift}=f_Shift),OR(f_Phase="(All)",f_Phase="",{Phase}=f_Phase)),1,0))'),
        dict(key="KeyOver", header="Overdue Rank Key", width=7, calc=True, fmt="0.0",
             formula='=IF(AND({Overdue}=1,{InScope}=1),N({Days})+ROW()/100000,0)'),
        dict(key="KeyOverF", header="Overdue Rank Key (Filtered)", width=7, calc=True, fmt="0.0",
             formula='={KeyOver}*{Filter}'),
    ]
    tcols = Cols(tspec)
    tdata = []
    bronze_ids = []
    for p, reqs in BRONZE.items():
        for i, req in enumerate(reqs, 1):
            bronze_ids.append((f"B-{p}-{i:02d}", p, req, SUGGEST[p][i - 1]))
    if not demo:
        for bid, p, req, sug in bronze_ids:
            tdata.append(dict(ID=bid.replace("B-", "T-"), Pillar=p, Phase="1 - Bronze", Req=req, BronzeRef=bid,
                              Required="Yes", Applic="Applicable", Weight=1, Status="Not Started",
                              VerStatus="Not Submitted"))
    else:
        tdata = demo_tasks(bronze_ids, AS_OF)
    first, last = build_table(ws, wb, tcols, 5, N_TRK, "tblTracker", tdata, name_prefix="T", freeze_col="F")
    TL = tcols.letter
    rag_cf(ws, f"{TL['RAG']}{first}:{TL['RAG']}{last}")
    text_cf(ws, f"{TL['DashStatus']}{first}:{TL['DashStatus']}{last}",
            {"Verified Complete": "Green", "Complete - Pending Verification": "Yellow", "In Progress": "Yellow",
             "Overdue": "Red", "Not Started": "Gray", "N/A": "Gray"})
    ws.conditional_formatting.add(f"{TL['Target']}{first}:{TL['Target']}{last}",
                                  FormulaRule(formula=[f"${TL['Overdue']}{first}=1"], font=Font(color="C00000", bold=True)))
    ws.conditional_formatting.add(f"{TL['Issues']}{first}:{TL['Issues']}{last}",
                                  FormulaRule(formula=[f'LEN(${TL["Issues"]}{first})>0'], font=Font(color="C55A11")))
    bar_cf(ws, f"{TL['Credit']}{first}:{TL['Credit']}{last}")
    ws.cell(4, 2, "Tip: filter/sort with the header arrows. Add tasks in the next empty row (500 rows pre-formatted). "
                  "To add more, drag the table's bottom-right handle – formulas copy automatically.").font = f(8, False, GRAY, True)

    # ================= GAP REGISTER =================
    ws = sh[S_GAP]
    title_band(ws, "Gap Register", tagline + "  ·  A gap is closed only when Status = Closed - Verified, "
               "Verification = Verified and a Closure Date is entered.", 40)
    gspec = [
        dict(key="ID", header="Gap ID", width=9),
        dict(key="DateId", header="Date Identified", width=10, dv="date", fmt=DATE),
        dict(key="Pillar", header="Pillar", width=8, dv="lst_Pillars"),
        dict(key="BronzeRef", header="Bronze Req ID", width=10, dv="lst_BronzeIDs"),
        dict(key="Requirement", header="Deployment Requirement", width=26, wrap=True),
        dict(key="TaskID", header="Linked Task ID", width=10),
        dict(key="Area", header="Process / Area", width=15, dv="lst_Area"),
        dict(key="Category", header="Gap Category", width=18, dv="lst_GapCat"),
        dict(key="Current", header="Current Condition", width=24, wrap=True),
        dict(key="RequiredC", header="Required Condition", width=24, wrap=True),
        dict(key="Desc", header="Gap Description", width=28, wrap=True),
        dict(key="Sev", header="Severity", width=9, dv="lst_Severity"),
        dict(key="Root", header="Root Cause", width=24, wrap=True),
        dict(key="CA", header="Corrective Action", width=26, wrap=True),
        dict(key="Owner", header="Assigned Owner", width=18, dv="lst_Owners"),
        dict(key="Due", header="Due Date", width=10, dv="date", fmt=DATE),
        dict(key="Status", header="Status", width=14, dv="lst_GapStatus", default=None),
        dict(key="EvidLink", header="Evidence (link)", width=22),
        dict(key="EvidOpen", header="Open", width=7, calc=True,
             formula='=IF({EvidLink}="","",HYPERLINK({EvidLink},"Open ↗"))'),
        dict(key="Verif", header="Verification", width=14, dv="lst_Verif"),
        dict(key="VerifiedBy", header="Verified By", width=16, dv="lst_Owners"),
        dict(key="Closure", header="Closure Date", width=10, dv="date", fmt=DATE),
        dict(key="Leader", header="Leadership Support Needed?", width=10, dv="lst_YesNo"),
        dict(key="Decision", header="Support / Decision Required", width=26, wrap=True),
        dict(key="Counted", header="Counted", width=6, calc=True, fmt="0",
             formula='=IF(AND({Pillar}<>"",{Status}<>"Cancelled"),1,0)'),
        dict(key="ClosedV", header="Closed & Verified", width=7, calc=True, fmt="0",
             formula='=IF(AND({Counted}=1,{Status}="Closed - Verified",{Verif}="Verified",ISNUMBER({Closure})),1,0)'),
        dict(key="Open", header="Open (1/0)", width=6, calc=True, fmt="0",
             formula='=IF(AND({Counted}=1,{ClosedV}=0),1,0)'),
        dict(key="Overdue", header="Overdue (1/0)", width=7, calc=True, fmt="0",
             formula='=IF(AND({Open}=1,ISNUMBER({Due})),IF({Due}<AsOf,1,0),0)'),
        dict(key="Age", header="Age (days)", width=7, calc=True, fmt="0",
             formula='=IF(OR({Counted}=0,NOT(ISNUMBER({DateId}))),"",IF({ClosedV}=1,{Closure},AsOf)-{DateId})'),
        dict(key="Bucket", header="Aging Bucket", width=9, calc=True,
             formula='=IF(OR({Open}=0,{Age}=""),"",IF({Age}<=cfg_Age1,"0-"&cfg_Age1&" d",IF({Age}<=cfg_Age2,cfg_Age1+1&"-"&cfg_Age2&" d",IF({Age}<=cfg_Age3,cfg_Age2+1&"-"&cfg_Age3&" d",">"&cfg_Age3&" d"))))'),
        dict(key="RAG", header="RAG", width=8, calc=True,
             formula='=IF({Pillar}="","",IF({Status}="Cancelled","N/A",IF({ClosedV}=1,"Green",IF(OR({Overdue}=1,{Sev}="Critical"),"Red",IF(OR({Sev}="High",{Owner}="",NOT(ISNUMBER({Due}))),"Yellow",IF({Due}-AsOf<=cfg_DueSoonDays,"Yellow","Green"))))))'),
        dict(key="Issues", header="Data Quality Flags", width=28, calc=True, wrap=True,
             formula=('=IF({Pillar}="","",IF({ID}="","Missing Gap ID; ","")&IF(AND({ID}<>"",COUNTIF(G_ID,{ID})>1),"Duplicate Gap ID; ","")'
                      '&IF({Owner}="","Missing owner; ","")&IF(NOT(ISNUMBER({Due})),"Missing due date; ","")'
                      '&IF({Sev}="","Missing severity; ","")&IF({Status}="","Missing status; ","")'
                      '&IF(NOT(ISNUMBER({DateId})),"Missing date identified; ","")'
                      '&IF(AND({Status}="Closed - Verified",{Verif}<>"Verified"),"Closed w/o verification; ","")'
                      '&IF(AND({Status}="Closed - Verified",NOT(ISNUMBER({Closure}))),"Closed w/o closure date; ","")'
                      '&IF(AND({Status}="Closed - Verified",{EvidLink}=""),"Closed w/o evidence; ","")'
                      '&IF(AND(OR({Status}="Pending Verification",{Status}="Closed - Verified"),{Root}=""),"Root cause missing; ","")'
                      '&IF({TaskID}<>"",IF(COUNTIF(T_ID,{TaskID})=0,"Linked task not found; ",""),""))')),
        dict(key="Filter", header="MR Filter", width=6, calc=True, fmt="0",
             formula='=IF({Pillar}="",0,IF(AND(OR(f_Pillar="(All)",f_Pillar="",{Pillar}=f_Pillar),OR(f_Owner="(All)",f_Owner="",{Owner}=f_Owner),OR(f_Area="(All)",f_Area="",{Area}=f_Area)),1,0))'),
        dict(key="Key", header="Priority Key", width=7, calc=True, fmt="0.0",
             formula='=IF({Open}=1,IFERROR(INDEX(lst_SevScore,MATCH({Sev},lst_SevFixed,0)),0)*10000+N({Age})+ROW()/100000,0)'),
        dict(key="KeyF", header="Priority Key (Filtered)", width=7, calc=True, fmt="0.0", formula='={Key}*{Filter}'),
        dict(key="LeadKey", header="Leadership Key", width=7, calc=True, fmt="0.0",
             formula='=IF(AND({Open}=1,OR({Leader}="Yes",{Sev}="Critical")),{Key},0)'),
        dict(key="LeadKeyF", header="Leadership Key (Filtered)", width=7, calc=True, fmt="0.0",
             formula='={LeadKey}*{Filter}'),
        dict(key="DecisionTxt", header="Leadership Item Text", width=20, calc=True, wrap=True,
             formula='=IF({Pillar}="","",IF({Decision}<>"",{Decision},"(No decision text) "&{Desc}))'),
    ]
    gcols = Cols(gspec)
    gdata = demo_gaps(AS_OF, bronze_ids) if demo else []
    first, last = build_table(ws, wb, gcols, 5, N_GAP, "tblGaps", gdata, name_prefix="G", freeze_col="E")
    GL = gcols.letter
    rag_cf(ws, f"{GL['RAG']}{first}:{GL['RAG']}{last}")
    text_cf(ws, f"{GL['Sev']}{first}:{GL['Sev']}{last}", {"Critical": "Red", "High": "Yellow"})
    ws.conditional_formatting.add(f"{GL['Issues']}{first}:{GL['Issues']}{last}",
                                  FormulaRule(formula=[f'LEN(${GL["Issues"]}{first})>0'], font=Font(color="C55A11")))

    # ================= BRONZE STAR READINESS =================
    ws = sh[S_BRZ]
    title_band(ws, "Bronze Star Readiness", "Criteria imported verbatim from " + SOURCE_DOC +
               ". The site is never shown as ready until every mandatory, applicable criterion is verified.", 30)
    bspec = [
        dict(key="ReqID", header="Req ID", width=10),
        dict(key="Pillar", header="Pillar", width=8, dv="lst_Pillars"),
        dict(key="Requirement", header="Bronze Requirement (official text)", width=40, wrap=True),
        dict(key="Source", header="Source", width=16, wrap=True),
        dict(key="Mandatory", header="Mandatory?", width=10, dv="lst_YesNo", default="Yes",
             note="Defaulted to Yes because the source lists these as Bronze Level requirements. Confirm with CHG."),
        dict(key="Applic", header="Applicable?", width=12, dv="lst_BrzApplic", default="Yes"),
        dict(key="Owner", header="Responsible Owner", width=18, dv="lst_Owners"),
        dict(key="Sugg", header="Suggested Evidence (CI recommendation – NOT CHG official)", width=34, wrap=True),
        dict(key="EvidReq", header="Required Evidence (CHG official – to be provided)", width=26, wrap=True),
        dict(key="EvidLink", header="Evidence Link", width=20),
        dict(key="EvidOpen", header="Open", width=7, calc=True,
             formula='=IF({EvidLink}="","",HYPERLINK({EvidLink},"Open ↗"))'),
        dict(key="Verif", header="Verification Status", width=14, dv="lst_BrzVerif", default="Not Assessed"),
        dict(key="Verifier", header="Verified By", width=16, dv="lst_Owners"),
        dict(key="VerDate", header="Verification Date", width=10, dv="date", fmt=DATE),
        dict(key="Due", header="Due Date", width=10, dv="date", fmt=DATE),
        dict(key="Gap", header="Gap", width=26, wrap=True),
        dict(key="Linked", header="Linked Tasks", width=7, calc=True, fmt="0",
             formula='=IF({ReqID}="","",COUNTIFS(T_BronzeRef,{ReqID},T_InScope,1))'),
        dict(key="TaskVer", header="Task Verified %", width=8, calc=True, fmt=PCT,
             formula='=IF(OR({ReqID}="",N({Linked})=0),"",SUMPRODUCT(--(T_BronzeRef={ReqID}),T_W,T_Verified)/MAX(0.0001,SUMPRODUCT(--(T_BronzeRef={ReqID}),T_W)))'),
        dict(key="TaskProg", header="Task Completion %", width=8, calc=True, fmt=PCT,
             formula='=IF(OR({ReqID}="",N({Linked})=0),0,SUMPRODUCT(--(T_BronzeRef={ReqID}),T_W,T_Credit)/MAX(0.0001,SUMPRODUCT(--(T_BronzeRef={ReqID}),T_W)))'),
        dict(key="OpenGaps", header="Open Gaps", width=6, calc=True, fmt="0",
             formula='=IF({ReqID}="","",COUNTIFS(G_BronzeRef,{ReqID},G_Open,1))'),
        dict(key="CritGaps", header="Critical Gaps", width=6, calc=True, fmt="0",
             formula='=IF({ReqID}="","",COUNTIFS(G_BronzeRef,{ReqID},G_Open,1,G_Sev,"Critical"))'),
        dict(key="CompStatus", header="Completion Status", width=15, calc=True,
             formula='=IF({ReqID}="","",IF({Applic}="N/A - Approved by CHG","N/A",IF(N({Linked})=0,"No tasks linked",IF(N({TaskVer})>=1,"Tasks Verified",IF(N({TaskProg})>0,"In Progress","Not Started")))))'),
        dict(key="Ready", header="Readiness Result", width=18, calc=True,
             formula='=IF({ReqID}="","",IF({Applic}="N/A - Approved by CHG","N/A (Approved)",IF(N({CritGaps})>0,"Blocked - Critical Gap",IF(AND({Verif}="Verified",{EvidLink}<>"",OR(N({Linked})=0,N({TaskVer})>=1),N({OpenGaps})=0),"Verified - Ready",IF(AND({EvidLink}="",N({TaskProg})=0,{Verif}<>"Verified"),"Not Started",IF({EvidLink}="","Missing Evidence",IF({Verif}<>"Verified","Pending Verification","Open Items Remaining")))))))'),
        dict(key="Counted", header="Mandatory & Applicable", width=7, calc=True, fmt="0",
             formula='=IF(AND({ReqID}<>"",{Mandatory}<>"No",{Applic}<>"N/A - Approved by CHG"),1,0)'),
        dict(key="MissingEv", header="Missing Evidence", width=7, calc=True, fmt="0",
             formula='=IF(AND({Counted}=1,{EvidLink}=""),1,0)'),
        dict(key="Blocker", header="Critical Blocker", width=7, calc=True, fmt="0",
             formula='=IF({Counted}=1,IF(OR(N({CritGaps})>0,AND(ISNUMBER({Due}),{Due}<AsOf,{Ready}<>"Verified - Ready")),1,0),0)'),
        dict(key="RAG", header="RAG", width=8, calc=True,
             formula='=IF({ReqID}="","",IF({Counted}=0,"N/A",IF({Ready}="Verified - Ready","Green",IF({Blocker}=1,"Red",IF({Ready}="Not Started","Gray","Yellow")))))'),
    ]
    bcols = Cols(bspec)
    bdata = [dict(ReqID=bid, Pillar=p, Requirement=req, Source="CHG Bronze Level Requirements", Sugg=sug,
                  Mandatory="Yes", Applic="Yes", Verif="Not Assessed") for bid, p, req, sug in bronze_ids]
    if demo:
        demo_bronze(bdata, tdata, AS_OF)
    HR = 17
    first, last = build_table(ws, wb, bcols, HR, N_BRZ, "tblBronze", bdata, name_prefix="B")
    BL = bcols.letter
    dyn_list(wb, "lst_BronzeIDs", ws, BL["ReqID"], first, last)
    rag_cf(ws, f"{BL['RAG']}{first}:{BL['RAG']}{last}")
    text_cf(ws, f"{BL['Ready']}{first}:{BL['Ready']}{last}",
            {"Verified - Ready": "Green", "Blocked - Critical Gap": "Red", "Missing Evidence": "Yellow",
             "Pending Verification": "Yellow", "Open Items Remaining": "Yellow", "Not Started": "Gray"})
    ws.freeze_panes = f"C{first}"
    ws.freeze_panes = None
    # summary (tiles + banner + horizontal pillar readiness)
    tile(ws, 5, 2, "BRONZE READINESS (verified mandatory criteria)", "=k_BronzePct", PCT,
         '=k_BronzeReady&" of "&k_BronzeMand&" verified"', 2)
    tile(ws, 5, 5, "MISSING EVIDENCE", "=k_BronzeMissingEv", "0", '="mandatory criteria w/o evidence link"', 2, BURG)
    tile(ws, 5, 8, "CRITICAL BLOCKERS", "=k_BronzeBlockers", "0", '="critical gap or past due"', 2, BURG)
    tile(ws, 5, 11, "NOT STARTED", '=COUNTIFS(B_Counted,1,B_Ready,"Not Started")', "0", '="mandatory criteria"', 2, GRAY)
    tile(ws, 5, 14, "PENDING VERIFICATION", '=COUNTIFS(B_Counted,1,B_Ready,"Pending Verification")+COUNTIFS(B_Counted,1,B_Ready,"Open Items Remaining")', "0", '="evidence submitted / open items"', 2, GRAY)
    ws.merge_cells("B9:P9")
    ws["B9"] = "=k_BronzeBanner"
    ws["B9"].font = f(12, True, WHITE)
    ws["B9"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[9].height = 24
    ws.conditional_formatting.add("B9:P9", FormulaRule(formula=['LEFT($B$9,5)="BRONZ"'], fill=fill("2E7D32")))
    ws.conditional_formatting.add("B9:P9", FormulaRule(formula=['LEFT($B$9,5)<>"BRONZ"'], fill=fill(BURG)))
    ws.cell(11, 2, "Pillar").font = f(9, True, WHITE)
    for j, lab in enumerate(["Mandatory", "Verified - Ready", "Readiness %", "RAG"]):
        ws.cell(12 + j, 2, lab).font = f(9, True, NAVY)
        ws.cell(12 + j, 2).border = BORDER
    ws.cell(11, 2).fill = fill(NAVY)
    for i in range(NPIL):
        c = 3 + i
        pr = CFG_LIST_ROW0 + i
        hc = ws.cell(11, c, f'=IF({qs(sh[S_CFG])}!$F${pr}="","",{qs(sh[S_CFG])}!$F${pr})')
        hc.font = f(9, True, WHITE)
        hc.fill = fill(NAVY)
        hc.alignment = Alignment(horizontal="center")
        P = f"{L(c)}$11"
        calc_cell(ws, 12, c, f'=IF({P}="","",COUNTIFS(B_Pillar,{P},B_Counted,1))', "0")
        calc_cell(ws, 13, c, f'=IF({P}="","",COUNTIFS(B_Pillar,{P},B_Counted,1,B_Ready,"Verified - Ready"))', "0")
        calc_cell(ws, 14, c, f'=IF(N({L(c)}12)=0,"",{L(c)}13/{L(c)}12)', PCT)
        calc_cell(ws, 15, c, f'=IF({P}="","",IF(N({L(c)}12)=0,"N/A",IF({L(c)}14>=1,"Green",IF(COUNTIFS(B_Pillar,{P},B_Blocker,1)>0,"Red",IF({L(c)}14>=cfg_BronzeYellow,"Yellow",IF(COUNTIFS(B_Pillar,{P},B_Counted,1,B_Ready,"Not Started")={L(c)}12,"Gray","Red"))))))')
        ws.column_dimensions[L(c)].width = max(ws.column_dimensions[L(c)].width or 0, 9)
    bar_cf(ws, f"C14:{L(2 + NPIL)}14")
    rag_cf(ws, f"C15:{L(2 + NPIL)}15")

    # ================= PROCESS DEPLOYMENT MATRIX =================
    ws = sh[S_PDM]
    title_band(ws, "Process Deployment Matrix", tagline + "  ·  Status per area/equipment/shift and deployment "
               "element. 'Not Assessed' and 'N/A' are excluded from scores; nothing is assumed to apply.", 36)
    mspec = [dict(key="RowID", header="Row ID", width=8),
             dict(key="Area", header="Production Area", width=20, dv="lst_Area"),
             dict(key="Equip", header="Equipment / Line", width=18),
             dict(key="Shift", header="Shift", width=6, dv="lst_Shift"),
             dict(key="Owner", header="Area Owner", width=18, dv="lst_Owners")]
    for j, e in enumerate(ELEMENTS, 1):
        mspec.append(dict(key=f"E{j}", header=e, width=12, dv="lst_MxStatus", default="Not Assessed", align="center"))
    mspec += [dict(key="LastVer", header="Last Verified", width=10, dv="date", fmt=DATE),
              dict(key="Notes", header="Notes", width=24, wrap=True)]
    tmp = Cols(mspec)
    er = f"{{E1}}:{{E9}}"
    E1, E9 = tmp.letter["E1"], tmp.letter["E9"]
    rngt = "${0}{{r}}:${1}{{r}}".format(E1, E9)
    def cnt(s):
        return f'COUNTIF(${E1}{{ROW}}:${E9}{{ROW}},"{s}")'
    appl = "(" + "+".join(cnt(s) for s in ["Not Started", "In Progress", "Implemented", "Verified"]) + ")"
    ssum = f'({cnt("In Progress")}*cfg_scIP+{cnt("Implemented")}*cfg_scImpl+{cnt("Verified")}*cfg_scVer)'
    mspec += [
        dict(key="Applic", header="Applicable Elements", width=8, calc=True, fmt="0",
             formula='=IF({Area}="","",' + appl + ')'),
        dict(key="Assessed", header="Assessed Elements", width=8, calc=True, fmt="0",
             formula='=IF({Area}="","",' + appl + '+' + cnt("N/A") + ')'),
        dict(key="ScoreSum", header="Score Points", width=7, calc=True, fmt="0.00",
             formula='=IF({Area}="",0,' + ssum + ')'),
        dict(key="VerCnt", header="Verified Elements", width=7, calc=True, fmt="0",
             formula='=IF({Area}="",0,' + cnt("Verified") + ')'),
        dict(key="Score", header="Deployment Score", width=9, calc=True, fmt=PCT,
             formula='=IF(OR({Area}="",N({Applic})=0),"",{ScoreSum}/{Applic})'),
        dict(key="VerPct", header="Verified %", width=8, calc=True, fmt=PCT,
             formula='=IF(OR({Area}="",N({Applic})=0),"",{VerCnt}/{Applic})'),
        dict(key="Coverage", header="Assessment Coverage", width=9, calc=True, fmt=PCT,
             formula=f'=IF({{Area}}="","",N({{Assessed}})/{len(ELEMENTS)})'),
        dict(key="RAG", header="RAG", width=8, calc=True,
             formula='=IF({Area}="","",IF(N({Applic})=0,"Gray",IF({Score}>=cfg_MatrixGreen,"Green",IF({Score}>=cfg_MatrixYellow,"Yellow","Red"))))'),
    ]
    for c in mspec:
        if "formula" in c:
            c["formula"] = c["formula"].replace("{ROW}", "{r}")
    mcols = Cols(mspec)
    # Cols.tpl does not know {r}: patch formulas per row after build
    mdata = demo_matrix(AS_OF) if demo else [dict(RowID=f"PA-{i + 1:02d}", Area=a, Shift="All")
                                             for i, a in enumerate(LISTS[6][2][:10])]
    HRm = 7
    first, last = build_table(ws, wb, mcols, HRm, N_PDM, "tblProcess", mdata, name_prefix="M", freeze_col="E")
    for c in mcols.spec:
        if c.get("calc"):
            col = mcols.first_col + mcols.spec.index(c)
            for r in range(first, last + 1):
                ws.cell(r, col).value = ws.cell(r, col).value.replace("{r}", str(r))
    ML = mcols.letter
    rng = f"{ML['E1']}{first}:{ML['E9']}{last}"
    for txt, (bg, fg) in {"Verified": RAG["Green"], "Implemented": ("DDEBF7", "1F4E79"),
                          "In Progress": RAG["Yellow"], "Not Started": RAG["Gray"],
                          "N/A": ("FFFFFF", "A6A6A6"), "Not Assessed": ("FFFFFF", "A6A6A6")}.items():
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{txt}"'], fill=fill(bg),
                                                      font=Font(name=FONT, color=fg)))
    rag_cf(ws, f"{ML['RAG']}{first}:{ML['RAG']}{last}")
    bar_cf(ws, f"{ML['Score']}{first}:{ML['Score']}{last}")
    # element summary aligned above element columns
    ws.cell(5, mcols.first_col + 4, "Element deployment score →").font = f(9, True, NAVY)
    ws.cell(6, mcols.first_col + 4, "Element verified % →").font = f(9, True, NAVY)
    for j in range(1, 10):
        Lc = ML[f"E{j}"]
        rg = f"${Lc}${first}:${Lc}${last}"
        ap = "+".join(f'COUNTIF({rg},"{s}")' for s in ["Not Started", "In Progress", "Implemented", "Verified"])
        sc = f'(COUNTIF({rg},"In Progress")*cfg_scIP+COUNTIF({rg},"Implemented")*cfg_scImpl+COUNTIF({rg},"Verified")*cfg_scVer)'
        calc_cell(ws, 5, mcols.first_col + 4 + j, f'=IF(({ap})=0,"",{sc}/({ap}))', PCT)
        calc_cell(ws, 6, mcols.first_col + 4 + j, f'=IF(({ap})=0,"",COUNTIF({rg},"Verified")/({ap}))', PCT)
    bar_cf(ws, f"{ML['E1']}5:{ML['E9']}5")
    ws.cell(4, 2, "Statuses: Not Assessed · N/A · Not Started · In Progress · Implemented · Verified. "
                  "Add equipment rows per area/shift as needed (80 rows).").font = f(8, False, GRAY, True)
    # Area / shift summary to the right of the table
    sc0 = mcols.first_col + len(mcols.spec) + 1
    section(ws, HRm - 2, sc0, "Deployment by Production Area (matrix + tracker)", 8)
    hdr_row(ws, HRm, sc0, ["Area", "Matrix Rows", "Matrix Score", "Matrix Verified %", "Coverage",
                           "Tracker Tasks", "Tracker Verified %", "Open Gaps"])
    for j, w in enumerate([22, 8, 9, 9, 9, 8, 9, 8]):
        ws.column_dimensions[L(sc0 + j)].width = w
    for i in range(20):
        r = HRm + 1 + i
        A = f"${L(sc0)}{r}"
        lr = CFG_LIST_ROW0 + i
        calc_cell(ws, r, sc0, f'=IF({qs(sh[S_CFG])}!${listcol["Area"]}${lr}="","",{qs(sh[S_CFG])}!${listcol["Area"]}${lr})', align="left")
        calc_cell(ws, r, sc0 + 1, f'=IF({A}="","",COUNTIF(M_Area,{A}))', "0")
        calc_cell(ws, r, sc0 + 2, f'=IF({A}="","",IF(SUMIF(M_Area,{A},M_Applic)=0,"",SUMIF(M_Area,{A},M_ScoreSum)/SUMIF(M_Area,{A},M_Applic)))', PCT)
        calc_cell(ws, r, sc0 + 3, f'=IF({A}="","",IF(SUMIF(M_Area,{A},M_Applic)=0,"",SUMIF(M_Area,{A},M_VerCnt)/SUMIF(M_Area,{A},M_Applic)))', PCT)
        calc_cell(ws, r, sc0 + 4, f'=IF(OR({A}="",N({L(sc0 + 1)}{r})=0),"",SUMIF(M_Area,{A},M_Assessed)/({L(sc0 + 1)}{r}*{len(ELEMENTS)}))', PCT)
        calc_cell(ws, r, sc0 + 5, f'=IF({A}="","",SUMPRODUCT(--(T_Area={A}),T_InScope))', "0")
        calc_cell(ws, r, sc0 + 6, f'=IF({A}="","",IF(SUMPRODUCT(--(T_Area={A}),T_W)=0,"",SUMPRODUCT(--(T_Area={A}),T_W,T_Verified)/SUMPRODUCT(--(T_Area={A}),T_W)))', PCT)
        calc_cell(ws, r, sc0 + 7, f'=IF({A}="","",COUNTIFS(G_Area,{A},G_Open,1))', "0")
    bar_cf(ws, f"{L(sc0 + 2)}{HRm + 1}:{L(sc0 + 2)}{HRm + 20}")
    bar_cf(ws, f"{L(sc0 + 6)}{HRm + 1}:{L(sc0 + 6)}{HRm + 20}")
    r0 = HRm + 23
    section(ws, r0, sc0, "Deployment by Shift", 8)
    hdr_row(ws, r0 + 1, sc0, ["Shift", "Matrix Rows", "Matrix Score", "Matrix Verified %", "", "Tracker Tasks",
                              "Tracker Verified %", ""])
    for i in range(6):
        r = r0 + 2 + i
        A = f"${L(sc0)}{r}"
        lr = CFG_LIST_ROW0 + i
        calc_cell(ws, r, sc0, f'=IF({qs(sh[S_CFG])}!${listcol["Shift"]}${lr}="","",{qs(sh[S_CFG])}!${listcol["Shift"]}${lr})', align="left")
        calc_cell(ws, r, sc0 + 1, f'=IF({A}="","",COUNTIF(M_Shift,{A}))', "0")
        calc_cell(ws, r, sc0 + 2, f'=IF({A}="","",IF(SUMIF(M_Shift,{A},M_Applic)=0,"",SUMIF(M_Shift,{A},M_ScoreSum)/SUMIF(M_Shift,{A},M_Applic)))', PCT)
        calc_cell(ws, r, sc0 + 3, f'=IF({A}="","",IF(SUMIF(M_Shift,{A},M_Applic)=0,"",SUMIF(M_Shift,{A},M_VerCnt)/SUMIF(M_Shift,{A},M_Applic)))', PCT)
        calc_cell(ws, r, sc0 + 5, f'=IF({A}="","",SUMPRODUCT(--(T_Shift={A}),T_InScope))', "0")
        calc_cell(ws, r, sc0 + 6, f'=IF({A}="","",IF(SUMPRODUCT(--(T_Shift={A}),T_W)=0,"",SUMPRODUCT(--(T_Shift={A}),T_W,T_Verified)/SUMPRODUCT(--(T_Shift={A}),T_W)))', PCT)
    pdm_area_ref = (sc0, HRm + 1)

    # ================= VO & TIGER TEAM PROJECTS =================
    ws = sh[S_VO]
    title_band(ws, "VO & Tiger Team Projects", tagline + "  ·  Estimated opportunity and Finance-verified savings "
               "are reported separately. Savings count only when Finance Approval = Approved.", 44)
    vspec = [
        dict(key="ID", header="Project ID", width=9),
        dict(key="Name", header="Project Name", width=30, wrap=True),
        dict(key="Pillar", header="Pillar", width=7, dv="lst_Pillars", default="VO"),
        dict(key="BronzeRef", header="Bronze Req ID", width=10, dv="lst_BronzeIDs"),
        dict(key="Area", header="Process / Area", width=15, dv="lst_Area"),
        dict(key="Type", header="Project Type", width=15, dv="lst_ProjType"),
        dict(key="Charter", header="Charter Status", width=10, dv="lst_Charter"),
        dict(key="CharterLink", header="Project Charter (link)", width=18),
        dict(key="CharterOpen", header="Open", width=7, calc=True,
             formula='=IF({CharterLink}="","",HYPERLINK({CharterLink},"Open ↗"))'),
        dict(key="Sponsor", header="Project Sponsor", width=16, dv="lst_Owners"),
        dict(key="Lead", header="Project Lead", width=16, dv="lst_Owners"),
        dict(key="Members", header="Tiger Team Members", width=22, wrap=True),
        dict(key="Problem", header="Problem Statement", width=30, wrap=True),
        dict(key="KPI", header="Primary KPI", width=16, wrap=True),
        dict(key="Unit", header="Unit", width=8),
        dict(key="Baseline", header="Baseline", width=9, fmt="#,##0.00"),
        dict(key="Target", header="Target", width=9, fmt="#,##0.00"),
        dict(key="Actual", header="Actual Performance", width=9, fmt="#,##0.00"),
        dict(key="Start", header="Start Date", width=10, dv="date", fmt=DATE),
        dict(key="TargetDate", header="Target Completion", width=10, dv="date", fmt=DATE),
        dict(key="ActualDate", header="Actual Completion", width=10, dv="date", fmt=DATE),
        dict(key="DMAIC", header="DMAIC Phase", width=9, dv="lst_DMAIC"),
        dict(key="Status", header="Implementation Status", width=17, dv="lst_ProjStatus"),
        dict(key="EstOpp", header="Est. Financial Opportunity ($/yr)", width=12, fmt="$#,##0;($#,##0);-", dv="num"),
        dict(key="VerSav", header="Verified Savings ($/yr)", width=12, fmt="$#,##0;($#,##0);-", dv="num"),
        dict(key="FinApp", header="Finance Approval", width=11, dv="lst_FinApp", default="Not Submitted"),
        dict(key="FinBy", header="Finance Approver", width=16, dv="lst_Owners"),
        dict(key="FinDate", header="Finance Approval Date", width=10, dv="date", fmt=DATE),
        dict(key="Sustain", header="Sustainability Verification", width=13, dv="lst_Sustain", default="Not Due"),
        dict(key="SustainDate", header="Sustain Check Date", width=10, dv="date", fmt=DATE),
        dict(key="Next", header="Next Action", width=24, wrap=True),
        dict(key="Notes", header="Notes", width=26, wrap=True),
        dict(key="Progress", header="Progress to Target", width=9, calc=True, fmt=PCT,
             formula='=IF(AND(ISNUMBER({Baseline}),ISNUMBER({Target}),ISNUMBER({Actual})),IF({Target}={Baseline},"",MAX(0,({Actual}-{Baseline})/({Target}-{Baseline})))," ")'),
        dict(key="IsActive", header="Active", width=6, calc=True, fmt="0",
             formula='=IF(OR({Status}="Planning — Baseline Pending",{Status}="Planning",{Status}="Active",{Status}="On Hold"),1,0)'),
        dict(key="IsDone", header="Completed", width=7, calc=True, fmt="0",
             formula='=IF(OR({Status}="Implemented",{Status}="Sustaining",{Status}="Closed"),1,0)'),
        dict(key="Opp", header="Opportunity Counted", width=10, calc=True, fmt="$#,##0;($#,##0);-",
             formula='=IF(OR({Name}="",{Status}="Cancelled"),0,N({EstOpp}))'),
        dict(key="FinSav", header="Finance-Verified Savings", width=10, calc=True, fmt="$#,##0;($#,##0);-",
             formula='=IF(AND({FinApp}="Approved",{Status}<>"Cancelled"),N({VerSav}),0)'),
        dict(key="Overdue", header="Overdue", width=6, calc=True, fmt="0",
             formula='=IF(AND({IsActive}=1,ISNUMBER({TargetDate})),IF({TargetDate}<AsOf,1,0),0)'),
        dict(key="RAG", header="RAG", width=7, calc=True,
             formula='=IF({Name}="","",IF({Status}="Cancelled","N/A",IF({Overdue}=1,"Red",IF(OR({Status}="On Hold",{Status}="Planning — Baseline Pending",AND({IsDone}=1,{Sustain}="Failed - Reopened")),"Yellow",IF(OR({Status}="Closed",{Status}="Sustaining"),"Green",IF({Status}="Planning","Gray","Green"))))))'),
        dict(key="Issues", header="Data Quality Flags", width=26, calc=True, wrap=True,
             formula=('=IF({Name}="","",IF({Lead}="","No project lead; ","")&IF({Sponsor}="","No sponsor; ","")'
                      '&IF(AND(N({VerSav})>0,{FinApp}<>"Approved"),"Savings not Finance-approved; ","")'
                      '&IF(AND({Status}="Active",{Charter}<>"Approved"),"Active w/o approved charter; ","")'
                      '&IF(AND(OR({Status}="Active",{IsDone}=1),NOT(ISNUMBER({Baseline}))),"Baseline missing; ","")'
                      '&IF(AND({Status}="Closed",{Sustain}="Not Due"),"Closed w/o sustainability check; ",""))')),
        dict(key="Filter", header="MR Filter", width=6, calc=True, fmt="0",
             formula='=IF({Name}="",0,IF(AND(OR(f_Pillar="(All)",f_Pillar="",{Pillar}=f_Pillar),OR(f_Owner="(All)",f_Owner="",{Lead}=f_Owner,{Sponsor}=f_Owner),OR(f_Area="(All)",f_Area="",{Area}=f_Area)),1,0))'),
    ]
    vcols = Cols(vspec)
    vdata = [dict(ID="VO-001", Name="Pre-Weigh Room Optimization & Relocation", Pillar="VO", Area="Pre-weigh",
                  Type="Process Flow / Layout", Charter="Draft", Status="Planning — Baseline Pending",
                  FinApp="Not Submitted", Sustain="Not Due",
                  Problem="To be defined in the project charter once baseline data is collected.",
                  Next="Collect baseline (labor hours, travel distance, weigh accuracy, throughput); finalize charter.",
                  Notes="Initial example project. Baseline, target and savings intentionally blank – do not "
                        "populate until measured and validated.")]
    if demo:
        vdata += demo_projects(AS_OF)
    HRv = 13
    first, last = build_table(ws, wb, vcols, HRv, N_VO, "tblProjects", vdata, name_prefix="V")
    VL = vcols.letter
    rag_cf(ws, f"{VL['RAG']}{first}:{VL['RAG']}{last}")
    ws.conditional_formatting.add(f"{VL['Issues']}{first}:{VL['Issues']}{last}",
                                  FormulaRule(formula=[f'LEN(${VL["Issues"]}{first})>0'], font=Font(color="C55A11")))
    tile(ws, 5, 2, "ACTIVE PROJECTS", "=SUM(V_IsActive)", "0", '=COUNTIF(V_Status,"Planning — Baseline Pending")&" baseline pending"', 2)
    tile(ws, 5, 5, "COMPLETED PROJECTS", "=SUM(V_IsDone)", "0", '=COUNTIF(V_Sustain,"90-Day Verified")&" sustained 90 days"', 2)
    tile(ws, 5, 8, "EST. OPPORTUNITY ($/yr)", "=SUM(V_Opp)", "$#,##0", '="estimate – not validated"', 2, GRAY)
    tile(ws, 5, 11, "FINANCE-VERIFIED SAVINGS ($/yr)", "=SUM(V_FinSav)", "$#,##0", '=COUNTIF(V_FinApp,"Approved")&" project(s) approved by Finance"', 2, "2E7D32")
    tile(ws, 5, 14, "OVERDUE PROJECTS", "=SUM(V_Overdue)", "0", '="active past target date"', 2, BURG)
    ws.cell(9, 2, "Projects by status").font = f(9, True, NAVY)
    stat_vals = LISTS[13][2]
    for j, s in enumerate(stat_vals):
        c = 3 + j
        hc = ws.cell(9, c + 1, s)
        hc.font = f(8, True, WHITE)
        hc.fill = fill(NAVY)
        hc.alignment = Alignment(wrap_text=True, horizontal="center")
        calc_cell(ws, 10, c + 1, f'=COUNTIF(V_Status,{L(c + 1)}9)', "0")
        calc_cell(ws, 11, c + 1, f'=SUMIF(V_Status,{L(c + 1)}9,V_FinSav)', "$#,##0;($#,##0);-")
    ws.cell(10, 3, "Count").font = f(8, True)
    ws.cell(11, 3, "Fin.-verified $").font = f(8, True)
    ws.row_dimensions[9].height = 30

    # ================= DOCUMENT & EVIDENCE REGISTER =================
    ws = sh[S_DOC]
    title_band(ws, "Document & Evidence Register", tagline + "  ·  Link to the approved file in Teams → DunGPS → "
               "General → Shared → <Pillar>. Do not duplicate files here.", 30)
    dspec = [
        dict(key="ID", header="Doc ID", width=9),
        dict(key="Name", header="Document Name", width=32, wrap=True),
        dict(key="Pillar", header="Pillar", width=8, dv="lst_Pillars"),
        dict(key="Linked", header="Linked Task / Req ID", width=11),
        dict(key="EvType", header="Evidence Type", width=16, dv="lst_EvType"),
        dict(key="Owner", header="Document Owner", width=18, dv="lst_Owners"),
        dict(key="Rev", header="Revision", width=7),
        dict(key="LastReview", header="Last Review Date", width=10, dv="date", fmt=DATE),
        dict(key="Freq", header="Review Freq. (months)", width=8, dv="num", fmt="0"),
        dict(key="Status", header="Approval / Verification Status", width=13, dv="lst_DocStatus"),
        dict(key="Link", header="Document Link", width=26),
        dict(key="Open", header="Open", width=7, calc=True,
             formula='=IF({Link}="","",HYPERLINK({Link},"Open ↗"))'),
        dict(key="Folder", header="Standard Teams Folder", width=30, calc=True,
             formula='=IF({Pillar}="","",cfg_TeamsRoot&" → "&IFERROR(INDEX(cfg_PillarFolder,MATCH({Pillar},cfg_PillarCode,0))&"",{Pillar}))'),
        dict(key="NextReview", header="Next Review Date", width=10, calc=True, fmt=DATE,
             formula='=IF(OR({Name}="",NOT(ISNUMBER({LastReview}))),"",EDATE({LastReview},IF(N({Freq})>0,{Freq},cfg_DocFreq)))'),
        dict(key="HasLink", header="Link Present", width=6, calc=True, fmt="0",
             formula='=IF(AND({Name}<>"",{Link}<>""),1,0)'),
        dict(key="RevOverdue", header="Review Overdue", width=7, calc=True, fmt="0",
             formula='=IF(AND({Name}<>"",ISNUMBER({NextReview}),{Status}<>"Obsolete"),IF({NextReview}<AsOf,1,0),0)'),
        dict(key="DueSoon", header="Review Due Soon", width=7, calc=True, fmt="0",
             formula='=IF(AND({RevOverdue}=0,ISNUMBER({NextReview})),IF({NextReview}-AsOf<=cfg_DocDueSoon,1,0),0)'),
        dict(key="Counted", header="Counted", width=6, calc=True, fmt="0",
             formula='=IF(AND({Name}<>"",{Status}<>"Obsolete"),1,0)'),
        dict(key="Complete", header="Complete", width=7, calc=True, fmt="0",
             formula='=IF(AND({Counted}=1,OR({Status}="Approved",{Status}="Verified"),{HasLink}=1,{RevOverdue}=0,ISNUMBER({LastReview})),1,0)'),
        dict(key="RAG", header="RAG", width=7, calc=True,
             formula='=IF({Name}="","",IF({Status}="Obsolete","N/A",IF({Complete}=1,IF({DueSoon}=1,"Yellow","Green"),IF(OR({RevOverdue}=1,{HasLink}=0),"Red","Yellow"))))'),
        dict(key="Issues", header="Data Quality Flags", width=26, calc=True, wrap=True,
             formula='=IF({Name}="","",IF({Link}="","No link; ","")&IF({Owner}="","No owner; ","")&IF({Rev}="","No revision; ","")&IF(NOT(ISNUMBER({LastReview})),"No review date; ","")&IF({RevOverdue}=1,"Review overdue; ",""))'),
    ]
    dcols = Cols(dspec)
    ddata = demo_docs(AS_OF, bronze_ids) if demo else []
    HRd = 15
    first, last = build_table(ws, wb, dcols, HRd, N_DOC, "tblDocs", ddata, name_prefix="D")
    DL = dcols.letter
    rag_cf(ws, f"{DL['RAG']}{first}:{DL['RAG']}{last}")
    tile(ws, 5, 2, "DOCUMENT COMPLETENESS", '=IF(SUM(D_Counted)=0,"–",SUM(D_Complete)/SUM(D_Counted))', PCT,
         '=SUM(D_Complete)&" of "&SUM(D_Counted)&" complete"', 2)
    tile(ws, 5, 5, "LINKED TO TEAMS", '=IF(SUM(D_Counted)=0,"–",SUMPRODUCT(D_Counted,D_HasLink)/SUM(D_Counted))', PCT,
         '="documents with a valid link"', 2, GRAY)
    tile(ws, 5, 8, "REVIEW OVERDUE", "=SUM(D_RevOverdue)", "0", '=SUM(D_DueSoon)&" due within "&cfg_DocDueSoon&" days"', 2, BURG)
    tile(ws, 5, 11, "APPROVED / VERIFIED", '=COUNTIF(D_Status,"Approved")+COUNTIF(D_Status,"Verified")', "0",
         '=COUNTIF(D_Status,"Draft")+COUNTIF(D_Status,"In Review")&" draft / in review"', 2, GRAY)
    ws.cell(9, 2, "Pillar").font = f(9, True, WHITE)
    ws.cell(9, 2).fill = fill(NAVY)
    for j, lab in enumerate(["Documents", "Complete", "Completeness %"]):
        ws.cell(10 + j, 2, lab).font = f(9, True, NAVY)
    for i in range(NPIL):
        c = 3 + i
        pr = CFG_LIST_ROW0 + i
        hc = ws.cell(9, c, f'=IF({qs(sh[S_CFG])}!$F${pr}="","",{qs(sh[S_CFG])}!$F${pr})')
        hc.font = f(9, True, WHITE)
        hc.fill = fill(NAVY)
        hc.alignment = Alignment(horizontal="center")
        P = f"{L(c)}$9"
        calc_cell(ws, 10, c, f'=IF({P}="","",SUMIF(D_Pillar,{P},D_Counted))', "0")
        calc_cell(ws, 11, c, f'=IF({P}="","",SUMIF(D_Pillar,{P},D_Complete))', "0")
        calc_cell(ws, 12, c, f'=IF(N({L(c)}10)=0,"",{L(c)}11/{L(c)}10)', PCT)
    bar_cf(ws, f"C12:{L(2 + NPIL)}12")

    # ================= CALC ENGINE =================
    build_calc(wb, sh, listcol)

    # ================= DASHBOARDS =================
    build_exec(wb, sh, demo)
    build_gapdash(wb, sh, demo)
    build_mr(wb, sh, demo)
    build_instructions(wb, sh, demo)

    # Owner-dependent DV for pillar responsibility list needs Calc names (defined in build_calc)
    for name in [S_TRK, S_GAP, S_OWN, S_BRZ, S_VO, S_DOC, S_PDM]:
        sh[name].sheet_view.zoomScale = 90
    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


# ------------------------------------------------------------------------------------------
def build_calc(wb, sh, listcol):
    ws = sh[S_CALC]
    CF = qs(sh[S_CFG])
    title_band(ws, "Calc Engine (system – do not edit)",
               "All dashboard numbers are calculated here from the source tables. Protected; no inputs.", 45, legend=False)
    # ---- helper dropdown lists (cols B..H, rows 6..) ----
    hdr_row(ws, 5, 2, ["Filter: Pillar", "Filter: Owner", "Filter: Area", "Filter: Shift", "Filter: Phase",
                       "Active Owners", "Pillar Resp."], calc=True)
    N = 205
    for k in range(N):
        r = 6 + k
        lr = CFG_LIST_ROW0 + k - 1
        if k == 0:
            for c in range(2, 7):
                ws.cell(r, c, "(All)")
            ws.cell(r, 8, "Site / Multiple")
        else:
            if lr <= CFG_LIST_ROWN:
                ws.cell(r, 2, f'=IF({CF}!$F${lr}="","",{CF}!$F${lr}&"")')
                ws.cell(r, 4, f'=IF({CF}!${listcol["Area"]}${lr}="","",{CF}!${listcol["Area"]}${lr}&"")')
                ws.cell(r, 5, f'=IF({CF}!${listcol["Shift"]}${lr}="","",{CF}!${listcol["Shift"]}${lr}&"")')
                ws.cell(r, 6, f'=IF({CF}!${listcol["Phase"]}${lr}="","",{CF}!${listcol["Phase"]}${lr}&"")')
                ws.cell(r, 8, f'=IF({CF}!$F${lr}="","",{CF}!$F${lr}&"")')
            ws.cell(r, 3, f"=G{r - 1}")
        if k < N_OWN:
            ws.cell(r, 7, f'=IFERROR(INDEX(OD_Name,MATCH({k + 1},OD_Seq,0))&"","")')
        for c in range(2, 9):
            ws.cell(r, c).font = f(8, False, GRAY)
    for c, nm in zip("BCDEFGH", ["lst_fPillar", "lst_fOwner", "lst_fArea", "lst_fShift", "lst_fPhase", "lst_Owners",
                                 "lst_PillarResp"]):
        dyn_list(wb, nm, ws, c, 6, 6 + N - 1, text_only=True)
        ws.column_dimensions[c].width = 14

    # ---- pillar blocks ----
    def tflt(filtered):
        return ",T_Filter" if filtered else ""

    def gflt(filtered):
        return ",G_Filter" if filtered else ""

    PCOLS = ["Pillar", "Name", "Tasks", "In Scope", "N/A Approved", "Optional", "Weight", "Verified Wt",
             "Verified %", "Completion Credit %", "Work Progress %", "Status Complete %", "Evidence Submitted %",
             "Evidence Verified %", "Planned %", "Variance (Credit−Plan)", "Overdue", "Not Started", "In Progress",
             "Pending Verification", "Verified Complete", "Open Gaps", "Critical Open Gaps", "Overdue Gaps",
             "Bronze Mandatory", "Bronze Verified", "Bronze Readiness %", "Planned Finish", "RAG"]
    C0 = 10  # column J
    col = {n: L(C0 + i) for i, n in enumerate(PCOLS)}

    def pillar_rows(r0, filtered, key_rng="T_Pillar", gkey="G_Pillar", bkey="B_Pillar", nrows=NPIL,
                    label_src="pillar", title=""):
        tf, gf = tflt(filtered), gflt(filtered)
        bf = ',--((f_Pillar="(All)")+(f_Pillar="")+(B_Pillar=f_Pillar)>0)' if filtered else ""
        section(ws, r0 - 1, C0, title, 12)
        hdr_row(ws, r0, C0, PCOLS, calc=True)
        rows = list(range(r0 + 1, r0 + 1 + nrows)) + [r0 + 1 + nrows]
        for r in rows:
            total = (r == rows[-1])
            P = f"$J{r}"
            if total:
                cond = '--(T_Pillar<>"")'
                gcond = '--(G_Pillar<>"")'
                bcond = '--(B_ReqID<>"")'
                ws.cell(r, C0, "TOTAL")
                ws.cell(r, C0 + 1, "Overall DunGPS")
                guard = "FALSE"
            else:
                i = r - r0 - 1
                lr = CFG_LIST_ROW0 + i
                if label_src == "pillar":
                    ws.cell(r, C0, f'=IF({CF}!$F${lr}="","",{CF}!$F${lr}&"")')
                    ws.cell(r, C0 + 1, f'=IF({CF}!$G${lr}="","",{CF}!$G${lr}&"")')
                else:
                    ws.cell(r, C0, f'=IF({CF}!${listcol["Phase"]}${lr}="","",{CF}!${listcol["Phase"]}${lr}&"")')
                cond = f"--({key_rng}={P})"
                gcond = f"--({gkey}={P})" if gkey else None
                bcond = f"--({bkey}={P})" if bkey else None
                guard = f'{P}=""'
            def w(cname, formula, fmt=None):
                c = ws.cell(r, C0 + PCOLS.index(cname), f"=IF({guard},\"\",{formula})")
                c.font = f(8, total)
                if fmt:
                    c.number_format = fmt
            R = lambda n: f"{col[n]}{r}"
            w("Tasks", f"SUMPRODUCT({cond}{tf})", "0")
            w("In Scope", f"SUMPRODUCT({cond},T_InScope{tf})", "0")
            w("N/A Approved", f'SUMPRODUCT({cond},--(T_Applic="N/A - Approved"){tf})', "0")
            w("Optional", f'SUMPRODUCT({cond},--(T_Required="No"){tf})', "0")
            w("Weight", f"SUMPRODUCT({cond},T_W{tf})", "0.0")
            w("Verified Wt", f"SUMPRODUCT({cond},T_W,T_Verified{tf})", "0.0")
            Wt = R("Weight")
            Ms = R("In Scope")
            w("Verified %", f'IF(N({Wt})=0,"",{R("Verified Wt")}/{Wt})', PCT)
            w("Completion Credit %", f'IF(N({Wt})=0,"",SUMPRODUCT({cond},T_W,T_Credit{tf})/{Wt})', PCT)
            w("Work Progress %", f'IF(N({Wt})=0,"",SUMPRODUCT({cond},T_W,T_Prog{tf})/{Wt})', PCT)
            w("Status Complete %", f'IF(N({Wt})=0,"",SUMPRODUCT({cond},T_W,--(T_Status="Complete"){tf})/{Wt})', PCT)
            w("Evidence Submitted %", f'IF(N({Ms})=0,"",SUMPRODUCT({cond},T_InScope,T_EvSub{tf})/{Ms})', PCT)
            w("Evidence Verified %", f'IF(N({Ms})=0,"",SUMPRODUCT({cond},T_InScope,T_EvVer{tf})/{Ms})', PCT)
            w("Planned %", f'IF(N({Wt})=0,"",SUMPRODUCT({cond},T_W,T_Planned{tf})/{Wt})', PCT)
            w("Variance (Credit−Plan)", f'IF(N({Wt})=0,"",{R("Completion Credit %")}-{R("Planned %")})', '+0%;-0%;0%')
            w("Overdue", f"SUMPRODUCT({cond},T_InScope,T_Overdue{tf})", "0")
            for nm, st in [("Not Started", "Not Started"), ("In Progress", "In Progress"),
                           ("Pending Verification", "Complete - Pending Verification"),
                           ("Verified Complete", "Verified Complete")]:
                w(nm, f'SUMPRODUCT({cond},T_InScope,--(T_DashStatus="{st}"){tf})', "0")
            if gcond:
                w("Open Gaps", f"SUMPRODUCT({gcond},G_Open{gf})", "0")
                w("Critical Open Gaps", f'SUMPRODUCT({gcond},G_Open,--(G_Sev="Critical"){gf})', "0")
                w("Overdue Gaps", f"SUMPRODUCT({gcond},G_Overdue{gf})", "0")
            else:
                for nm in ["Open Gaps", "Critical Open Gaps", "Overdue Gaps"]:
                    w(nm, "0", "0")
            if bcond:
                w("Bronze Mandatory", f"SUMPRODUCT({bcond},B_Counted{bf})", "0")
                w("Bronze Verified", f'SUMPRODUCT({bcond},B_Counted,--(B_Ready="Verified - Ready"){bf})', "0")
                w("Bronze Readiness %", f'IF(N({R("Bronze Mandatory")})=0,"",{R("Bronze Verified")}/{R("Bronze Mandatory")})', PCT)
            else:
                for nm in ["Bronze Mandatory", "Bronze Verified"]:
                    w(nm, "0", "0")
                w("Bronze Readiness %", '""', PCT)
            fl = "*T_Filter" if filtered else ""
            cnd = "(T_Pillar<>\"\")" if total else f"({key_rng}={P})"
            w("Planned Finish", f'IFERROR(1/(1/SUMPRODUCT(MAX({cnd}*T_InScope*(1-T_Verified)*T_Target{fl}))),"")', DATE)
            crit = R("Critical Open Gaps")
            w("RAG", (f'IF(N({Wt})=0,"Gray",IF(OR(N({crit})>0,N({R("Overdue")})/MAX(1,N({Ms}))>=cfg_RedOverdueShare),"Red",'
                      f'IF(AND(N({R("Completion Credit %")})=0,N({R("Planned %")})=0),"Gray",'
                      f'IF({R("Variance (Credit−Plan)")}>=-cfg_GreenTol,IF(N({R("Overdue")})>0,"Yellow","Green"),'
                      f'IF({R("Variance (Credit−Plan)")}>=-cfg_YellowTol,"Yellow","Red")))))'))
            if not total:
                pass
        return rows

    rows_p = pillar_rows(6, False, title="Pillar metrics – site (unfiltered)")
    rows_pf = pillar_rows(6 + NPIL + 5, True, title="Pillar metrics – Management Review (filtered)")
    r_phase = 6 + 2 * (NPIL + 5)
    rows_ph = pillar_rows(r_phase, False, key_rng="T_Phase", gkey=None, bkey=None, nrows=8, label_src="phase",
                          title="Deployment phase metrics (unfiltered)")
    define(wb, "c_PillarTop", f"{qs(ws)}!$J${rows_p[0]}")
    meta = dict(col=col, rows_p=rows_p, rows_pf=rows_pf, rows_ph=rows_ph, PCOLS=PCOLS)

    # ---- trend (26 periods ending at status date) ----
    r0 = rows_ph[-1] + 4
    section(ws, r0 - 1, C0, "Planned vs actual – cumulative verified completion (26 periods)", 4)
    hdr_row(ws, r0, C0, ["Period End", "Planned Cum. %", "Actual Verified Cum. %", "Label"], calc=True)
    tot = rows_p[-1]
    Wtot = f"$P${tot}"
    for k in range(26):
        r = r0 + 1 + k
        ws.cell(r, C0, f"=AsOf-({25 - k})*cfg_TrendDays").number_format = DATE
        ws.cell(r, C0 + 1, f'=IF(N({Wtot})=0,0,SUMPRODUCT(T_W,--(T_Target<>""),--(T_Target<=$J{r}))/{Wtot})').number_format = PCT
        ws.cell(r, C0 + 2, f'=IF(N({Wtot})=0,0,SUMPRODUCT(T_W,T_Verified,--(T_DoneDate<=$J{r}))/{Wtot})').number_format = PCT
        ws.cell(r, C0 + 3, f'=TEXT($J{r},"dd-mmm")')
    meta["trend"] = (r0 + 1, r0 + 26)

    # ---- forward load ----
    r0 = r0 + 30
    section(ws, r0 - 1, C0, "Closure forecast – open items by due month", 4)
    hdr_row(ws, r0, C0, ["Due Bucket", "Open Tasks", "Open Gaps", "Start", "End"], calc=True)
    labels = ["Overdue"] + [None] * 6 + ["Later", "No due date"]
    for k, lab in enumerate(labels):
        r = r0 + 1 + k
        if k == 0:
            ws.cell(r, C0, "Overdue")
            ws.cell(r, C0 + 1, "=SUMPRODUCT(T_InScope,T_Overdue)")
            ws.cell(r, C0 + 2, "=SUM(G_Overdue)")
        elif k <= 6:
            m = k - 1
            ws.cell(r, C0 + 3, f"=IF({m}=0,AsOf,EOMONTH(AsOf,{m - 1})+1)").number_format = DATE
            ws.cell(r, C0 + 4, f"=EOMONTH(AsOf,{m})").number_format = DATE
            ws.cell(r, C0, f'=TEXT(EOMONTH(AsOf,{m}),"mmm-yy")')
            ws.cell(r, C0 + 1, f'=COUNTIFS(T_InScope,1,T_Verified,0,T_Target,">="&$M{r},T_Target,"<="&$N{r})')
            ws.cell(r, C0 + 2, f'=COUNTIFS(G_Open,1,G_Due,">="&$M{r},G_Due,"<="&$N{r})')
        elif lab == "Later":
            ws.cell(r, C0, "Later")
            ws.cell(r, C0 + 1, '=COUNTIFS(T_InScope,1,T_Verified,0,T_Target,">"&EOMONTH(AsOf,5))')
            ws.cell(r, C0 + 2, '=COUNTIFS(G_Open,1,G_Due,">"&EOMONTH(AsOf,5))')
        else:
            ws.cell(r, C0, "No due date")
            ws.cell(r, C0 + 1, '=COUNTIFS(T_InScope,1,T_Verified,0,T_Target,"")')
            ws.cell(r, C0 + 2, '=COUNTIFS(G_Open,1,G_Due,"")')
    meta["fwd"] = (r0 + 1, r0 + 9)

    # ---- gap monthly ----
    r0 = r0 + 13
    section(ws, r0 - 1, C0, "Gaps – new vs. verified closed per month (last 12 months)", 5)
    hdr_row(ws, r0, C0, ["Month", "New Gaps", "Closed (Verified)", "Open Backlog (month end)", "Start", "End"], calc=True)
    for k in range(12):
        r = r0 + 1 + k
        m = k - 11
        ws.cell(r, C0 + 4, f"=EOMONTH(AsOf,{m - 1})+1").number_format = DATE
        ws.cell(r, C0 + 5, f"=EOMONTH(AsOf,{m})").number_format = DATE
        ws.cell(r, C0, f'=TEXT($N{r},"mmm-yy")')
        ws.cell(r, C0 + 1, f'=COUNTIFS(G_Counted,1,G_DateId,">="&$N{r},G_DateId,"<="&$O{r})')
        ws.cell(r, C0 + 2, f'=COUNTIFS(G_ClosedV,1,G_Closure,">="&$N{r},G_Closure,"<="&$O{r})')
        ws.cell(r, C0 + 3, f'=COUNTIFS(G_Counted,1,G_DateId,"<="&$O{r})-COUNTIFS(G_ClosedV,1,G_Closure,"<="&$O{r})')
    meta["monthly"] = (r0 + 1, r0 + 12)

    # ---- Pareto ----
    r0 = r0 + 16
    section(ws, r0 - 1, C0, "Gap / barrier category Pareto (gaps + tracker barriers)", 8)
    hdr_row(ws, r0, C0, ["Category", "Gaps", "Task Barriers", "Total", "Sort Key", "Rank", "Sorted Category",
                         "Sorted Total", "Cumulative %"], calc=True)
    NC = 20
    for k in range(NC):
        r = r0 + 1 + k
        lr = CFG_LIST_ROW0 + k
        gl = listcol["GapCat"]
        ws.cell(r, C0, f'=IF({CF}!${gl}${lr}="","",{CF}!${gl}${lr}&"")')
        ws.cell(r, C0 + 1, f'=IF($J{r}="",0,COUNTIFS(G_Category,$J{r},G_Counted,1))')
        ws.cell(r, C0 + 2, f'=IF($J{r}="",0,COUNTIFS(T_BarrierCat,$J{r},T_InScope,1,T_Verified,0))')
        ws.cell(r, C0 + 3, f"=K{r}+L{r}")
        ws.cell(r, C0 + 4, f'=IF($J{r}="",-1,M{r}+({NC}-{k})/1000)')
        ws.cell(r, C0 + 5, k + 1)
        ws.cell(r, C0 + 6, f'=IF(LARGE($N${r0 + 1}:$N${r0 + NC},{k + 1})<0,"",INDEX($J${r0 + 1}:$J${r0 + NC},MATCH(LARGE($N${r0 + 1}:$N${r0 + NC},{k + 1}),$N${r0 + 1}:$N${r0 + NC},0)))')
        ws.cell(r, C0 + 7, f'=IF(P{r}="",0,INDEX($M${r0 + 1}:$M${r0 + NC},MATCH(P{r},$J${r0 + 1}:$J${r0 + NC},0)))')
        ws.cell(r, C0 + 8, f'=IF(SUM($Q${r0 + 1}:$Q${r0 + NC})=0,0,SUM($Q${r0 + 1}:Q{r})/SUM($Q${r0 + 1}:$Q${r0 + NC}))').number_format = PCT
    meta["pareto"] = (r0 + 1, r0 + 10)  # show top 10

    # ---- severity & aging ----
    r0 = r0 + NC + 4
    section(ws, r0 - 1, C0, "Gaps by severity", 5)
    hdr_row(ws, r0, C0, ["Severity", "Open", "Overdue", "Closed (Verified)", "Total Identified"], calc=True)
    for k, s in enumerate(["Critical", "High", "Medium", "Low"]):
        r = r0 + 1 + k
        ws.cell(r, C0, s)
        ws.cell(r, C0 + 1, f'=COUNTIFS(G_Sev,$J{r},G_Open,1)')
        ws.cell(r, C0 + 2, f'=COUNTIFS(G_Sev,$J{r},G_Overdue,1)')
        ws.cell(r, C0 + 3, f'=COUNTIFS(G_Sev,$J{r},G_ClosedV,1)')
        ws.cell(r, C0 + 4, f'=COUNTIFS(G_Sev,$J{r},G_Counted,1)')
    r = r0 + 5
    ws.cell(r, C0, "Not set")
    ws.cell(r, C0 + 1, '=COUNTIFS(G_Sev,"",G_Open,1)')
    ws.cell(r, C0 + 2, '=COUNTIFS(G_Sev,"",G_Overdue,1)')
    ws.cell(r, C0 + 3, '=COUNTIFS(G_Sev,"",G_ClosedV,1)')
    ws.cell(r, C0 + 4, '=COUNTIFS(G_Sev,"",G_Counted,1)')
    meta["sev"] = (r0 + 1, r0 + 5)
    r0 = r0 + 9
    section(ws, r0 - 1, C0, "Open gap aging", 3)
    hdr_row(ws, r0, C0, ["Age Bucket", "Open Gaps"], calc=True)
    ag = [('="0-"&cfg_Age1&" days"', '=COUNTIFS(G_Open,1,G_Age,"<="&cfg_Age1)'),
          ('=cfg_Age1+1&"-"&cfg_Age2&" days"', '=COUNTIFS(G_Open,1,G_Age,">"&cfg_Age1,G_Age,"<="&cfg_Age2)'),
          ('=cfg_Age2+1&"-"&cfg_Age3&" days"', '=COUNTIFS(G_Open,1,G_Age,">"&cfg_Age2,G_Age,"<="&cfg_Age3)'),
          ('=">"&cfg_Age3&" days"', '=COUNTIFS(G_Open,1,G_Age,">"&cfg_Age3)'),
          ("No date identified", '=COUNTIFS(G_Open,1,G_Age,"")')]
    for k, (a, b) in enumerate(ag):
        ws.cell(r0 + 1 + k, C0, a)
        ws.cell(r0 + 1 + k, C0 + 1, b)
    meta["aging"] = (r0 + 1, r0 + 5)

    # ---- owner table ----
    r0 = r0 + 9
    section(ws, r0 - 1, C0, "Owner accountability (one row per Owner Directory row)", 12)
    OC = ["Owner", "Active", "Tasks (in scope)", "Verified", "Overdue Tasks", "Open Gaps", "Overdue Gaps",
          "Verified %", "Rank Key", "Gap Rank Key", "F: Tasks", "F: Verified", "F: Overdue", "F: Open Gaps",
          "F: Overdue Gaps", "F: Verified %", "F: Rank Key"]
    hdr_row(ws, r0, C0, OC, calc=True)
    oc = {n: L(C0 + i) for i, n in enumerate(OC)}
    for i in range(N_OWN + 1):
        r = r0 + 1 + i
        un = (i == N_OWN)
        if un:
            ws.cell(r, C0, "(Unassigned)")
            nm = '""'
            g = "FALSE"
        else:
            ws.cell(r, C0, f'=IF(INDEX(OD_Name,{i + 1})="","",INDEX(OD_Name,{i + 1})&"")')
            ws.cell(r, C0 + 1, f'=IF($J{r}="","",INDEX(OD_Active,{i + 1})&"")')
            nm = f"$J{r}"
            g = f'$J{r}=""'
        def o(cn, fm, fmt="0"):
            c = ws.cell(r, C0 + OC.index(cn), f'=IF({g},0,{fm})')
            c.number_format = fmt
        o("Tasks (in scope)", f"SUMPRODUCT(--(T_Owner={nm}),T_InScope)")
        o("Verified", f"SUMPRODUCT(--(T_Owner={nm}),T_InScope,T_Verified)")
        o("Overdue Tasks", f"SUMPRODUCT(--(T_Owner={nm}),T_InScope,T_Overdue)")
        o("Open Gaps", f"SUMPRODUCT(--(G_Owner={nm}),G_Open)")
        o("Overdue Gaps", f"SUMPRODUCT(--(G_Owner={nm}),G_Overdue)")
        o("Verified %", f'IF({oc["Tasks (in scope)"]}{r}=0,0,{oc["Verified"]}{r}/{oc["Tasks (in scope)"]}{r})', PCT)
        o("Rank Key", f'IF({oc["Tasks (in scope)"]}{r}+{oc["Open Gaps"]}{r}=0,0,{oc["Overdue Tasks"]}{r}*1000000+{oc["Overdue Gaps"]}{r}*10000+({oc["Tasks (in scope)"]}{r}-{oc["Verified"]}{r})*10+{oc["Open Gaps"]}{r}+{(N_OWN + 2 - i)}/100000)', "0.00")
        o("Gap Rank Key", f'IF({oc["Overdue Gaps"]}{r}=0,0,{oc["Overdue Gaps"]}{r}*1000+{oc["Open Gaps"]}{r}+{(N_OWN + 2 - i)}/100000)', "0.00")
        o("F: Tasks", f"SUMPRODUCT(--(T_Owner={nm}),T_InScope,T_Filter)")
        o("F: Verified", f"SUMPRODUCT(--(T_Owner={nm}),T_InScope,T_Verified,T_Filter)")
        o("F: Overdue", f"SUMPRODUCT(--(T_Owner={nm}),T_InScope,T_Overdue,T_Filter)")
        o("F: Open Gaps", f"SUMPRODUCT(--(G_Owner={nm}),G_Open,G_Filter)")
        o("F: Overdue Gaps", f"SUMPRODUCT(--(G_Owner={nm}),G_Overdue,G_Filter)")
        o("F: Verified %", f'IF({oc["F: Tasks"]}{r}=0,0,{oc["F: Verified"]}{r}/{oc["F: Tasks"]}{r})', PCT)
        o("F: Rank Key", f'IF({oc["F: Tasks"]}{r}+{oc["F: Open Gaps"]}{r}=0,0,{oc["F: Overdue"]}{r}*1000000+{oc["F: Overdue Gaps"]}{r}*10000+({oc["F: Tasks"]}{r}-{oc["F: Verified"]}{r})*10+{oc["F: Open Gaps"]}{r}+{(N_OWN + 2 - i)}/100000)', "0.00")
    o_first, o_last = r0 + 1, r0 + 1 + N_OWN
    for cn, nmn in [("Owner", "c_OwnName"), ("Rank Key", "c_OwnKey"), ("Gap Rank Key", "c_OwnGapKey"),
                    ("F: Rank Key", "c_OwnKeyF"), ("Overdue Tasks", "c_OwnOverdue")]:
        define(wb, nmn, f"{qs(ws)}!${oc[cn]}${o_first}:${oc[cn]}${o_last}")
    meta["owners"] = (o_first, o_last, oc)

    # ---- VO by status (filtered) ----
    r0 = o_last + 4
    section(ws, r0 - 1, C0, "VO projects by status", 6)
    hdr_row(ws, r0, C0, ["Status", "Projects", "Est. Opportunity", "Finance-Verified", "F: Projects",
                         "F: Est. Opportunity", "F: Finance-Verified"], calc=True)
    for k, s in enumerate(LISTS[13][2]):
        r = r0 + 1 + k
        ws.cell(r, C0, s)
        ws.cell(r, C0 + 1, f'=COUNTIF(V_Status,$J{r})')
        ws.cell(r, C0 + 2, f'=SUMIF(V_Status,$J{r},V_Opp)').number_format = "$#,##0"
        ws.cell(r, C0 + 3, f'=SUMIF(V_Status,$J{r},V_FinSav)').number_format = "$#,##0"
        ws.cell(r, C0 + 4, f'=COUNTIFS(V_Status,$J{r},V_Filter,1)')
        ws.cell(r, C0 + 5, f'=SUMIFS(V_Opp,V_Status,$J{r},V_Filter,1)').number_format = "$#,##0"
        ws.cell(r, C0 + 6, f'=SUMIFS(V_FinSav,V_Status,$J{r},V_Filter,1)').number_format = "$#,##0"
    meta["vo"] = (r0 + 1, r0 + len(LISTS[13][2]))

    # ---- scalar KPIs ----
    K0 = 45  # col AS
    hdr_row(ws, 5, K0, ["KPI (named)", "Value"], calc=True)
    ws.column_dimensions[L(K0)].width = 26
    ws.column_dimensions[L(K0 + 1)].width = 16
    T = rows_p[-1]
    TF = rows_pf[-1]
    kp = [
        ("k_GapTotal", "=SUM(G_Counted)", "0"),
        ("k_GapOpen", "=SUM(G_Open)", "0"),
        ("k_GapOverdue", "=SUM(G_Overdue)", "0"),
        ("k_GapCritOpen", '=COUNTIFS(G_Open,1,G_Sev,"Critical")', "0"),
        ("k_GapPendVer", '=COUNTIFS(G_Open,1,G_Status,"Pending Verification")', "0"),
        ("k_GapClosedV", "=SUM(G_ClosedV)", "0"),
        ("k_GapClosure", '=IF(k_GapTotal=0,"",k_GapClosedV/k_GapTotal)', PCT),
        ("k_GapAvgAge", '=IFERROR(AVERAGEIFS(G_Age,G_Open,1),"")', "0"),
        ("k_GapLeader", '=COUNTIFS(G_Open,1,G_Leader,"Yes")', "0"),
        ("k_GapLeaderNonCrit", '=COUNTIFS(G_Open,1,G_Leader,"Yes",G_Sev,"<>Critical")', "0"),
        ("k_GapNoDue", '=COUNTIFS(G_Open,1,G_Due,"")', "0"),
        ("k_GapMaxDue", "=IFERROR(SUMPRODUCT(MAX(G_Open*G_Due)),0)", DATE),
        ("k_TaskEscalated", "=SUMPRODUCT(T_InScope,T_Overdue,--(T_KeyOver>cfg_EscalationDays))", "0"),
        ("k_LeaderIssues", "=k_GapCritOpen+k_GapLeaderNonCrit+k_TaskEscalated", "0"),
        ("k_DataFlags", '=COUNTIF(T_Issues,"?*")', "0"),
        ("k_GapDataFlags", '=COUNTIF(G_Issues,"?*")', "0"),
        ("k_NA", '=COUNTIFS(T_Applic,"N/A - Approved",T_Pillar,"<>")', "0"),
        ("k_NAReq", '=COUNTIFS(T_Applic,"N/A - Requested",T_Pillar,"<>")', "0"),
        ("k_Optional", '=COUNTIFS(T_Required,"No",T_Pillar,"<>")', "0"),
        ("k_Unassigned", '=SUMPRODUCT(T_InScope,--(T_Owner=""))', "0"),
        ("k_EvSubCount", "=SUMPRODUCT(T_InScope,T_EvSub)", "0"),
        ("k_MaxTarget", "=IFERROR(SUMPRODUCT(MAX(T_InScope*(1-T_Verified)*T_Target)),0)", DATE),
        ("k_BronzeMand", "=SUM(B_Counted)", "0"),
        ("k_BronzeReady", '=COUNTIFS(B_Counted,1,B_Ready,"Verified - Ready")', "0"),
        ("k_BronzePct", '=IF(k_BronzeMand=0,0,k_BronzeReady/k_BronzeMand)', PCT),
        ("k_BronzeMissingEv", "=SUM(B_MissingEv)", "0"),
        ("k_BronzeBlockers", "=SUM(B_Blocker)", "0"),
        ("k_BronzeBanner", '=IF(AND(k_BronzeMand>0,k_BronzeReady=k_BronzeMand),"BRONZE STAR READY – all "&k_BronzeMand&" mandatory criteria verified","NOT BRONZE STAR READY – "&k_BronzeReady&" of "&k_BronzeMand&" mandatory criteria verified · "&k_BronzeMissingEv&" missing evidence · "&k_BronzeBlockers&" critical blockers")', None),
        ("k_TopOwner", '=IF(MAX(c_OwnOverdue)=0,"",INDEX(c_OwnName,MATCH(MAX(c_OwnOverdue),c_OwnOverdue,0)))', None),
        ("k_TopOwnerN", "=MAX(c_OwnOverdue)", "0"),
        ("k_OwnersWithOverdue", '=COUNTIF(c_OwnOverdue,">0")', "0"),
        ("k_VOActive", "=SUM(V_IsActive)", "0"),
        ("k_VODone", "=SUM(V_IsDone)", "0"),
        ("k_VOOpp", "=SUM(V_Opp)", "$#,##0"),
        ("k_VOFinSav", "=SUM(V_FinSav)", "$#,##0"),
        ("k_VOActiveF", "=SUMPRODUCT(V_IsActive,V_Filter)", "0"),
        ("k_VOOppF", "=SUMPRODUCT(V_Opp,V_Filter)", "$#,##0"),
        ("k_VOFinSavF", "=SUMPRODUCT(V_FinSav,V_Filter)", "$#,##0"),
        ("k_GapOpenF", "=SUMPRODUCT(G_Open,G_Filter)", "0"),
        ("k_GapOverdueF", "=SUMPRODUCT(G_Overdue,G_Filter)", "0"),
        ("k_DocPct", '=IF(SUM(D_Counted)=0,"",SUM(D_Complete)/SUM(D_Counted))', PCT),
    ]
    for i, (nm, fm, fmt) in enumerate(kp):
        r = 6 + i
        ws.cell(r, K0, nm).font = f(8)
        c = ws.cell(r, K0 + 1, fm)
        c.font = f(8)
        if fmt:
            c.number_format = fmt
        define(wb, nm, f"{qs(ws)}!${L(K0 + 1)}${r}")
    for cn in PCOLS:
        define(wb, "tot_" + "".join(ch for ch in cn if ch.isalnum()), f"{qs(ws)}!${col[cn]}${T}")
        define(wb, "totf_" + "".join(ch for ch in cn if ch.isalnum()), f"{qs(ws)}!${col[cn]}${TF}")
    for c in range(C0, C0 + 30):
        ws.column_dimensions[L(c)].width = 11
    ws.column_dimensions["K"].width = 22
    protect(ws)
    wb._dunmeta = meta


# ------------------------------------------------------------------------------------------
def top_list(ws, r0, c0, n, key_rng, out_cols, widths=None):
    """Writes top-N lookup rows using LARGE on a numeric key range. out_cols: list of (range name, fmt)."""
    for k in range(1, n + 1):
        r = r0 + k - 1
        m = f"MATCH(LARGE({key_rng},{k}),{key_rng},0)"
        for j, (src, fmt) in enumerate(out_cols):
            fm = f'=IF(LARGE({key_rng},{k})<=0,"",INDEX({src},{m})&"")' if fmt is None else \
                 f'=IF(LARGE({key_rng},{k})<=0,"",IF(INDEX({src},{m})="","",INDEX({src},{m})))'
            c = calc_cell(ws, r, c0 + j, fm, fmt, align="left" if fmt is None else "center", size=8)


def build_exec(wb, sh, demo):
    ws = sh[S_EXEC]
    meta = wb._dunmeta
    col, rows_p, rows_ph = meta["col"], meta["rows_p"], meta["rows_ph"]
    CA = qs(sh[S_CALC])
    title_band(ws, "DunGPS Pillar Deployment — Executive Dashboard",
               "", 17, legend=False)
    ws["B3"] = ('=cfg_SiteName&"  ·  Status as of "&TEXT(AsOf,"dd-mmm-yyyy")&"  ·  Verified completion = verified required tasks ÷ applicable required tasks"'
                + ('&"  ·  DEMONSTRATION DATA"' if demo else ""))
    widths = {"A": 2, "B": 9, "C": 24}
    for c in range(4, 18):
        ws.column_dimensions[L(c)].width = 10
    for k, v in widths.items():
        ws.column_dimensions[k].width = v
    # KPI tiles: B..Q (16 cols) -> 8 tiles x 2
    T = lambda n: "tot_" + "".join(ch for ch in n if ch.isalnum())
    tiles = [
        ("VERIFIED COMPLETION", f'=IF(N({T("Weight")})=0,0,{T("Verified %")})', PCT,
         f'=TEXT(N({T("Verified Complete")}),"0")&" of "&TEXT(N({T("In Scope")}),"0")&" required tasks verified"', NAVY),
        ("COMPLETION CREDIT", f'=IF(N({T("Weight")})=0,0,{T("Completion Credit %")})', PCT,
         f'="Reported work progress: "&TEXT(N({T("Work Progress %")}),"0%")', NAVY),
        ("PLANNED TO DATE", f'=IF(N({T("Weight")})=0,0,{T("Planned %")})', PCT,
         f'="Variance: "&TEXT(N({T("Variance (Credit−Plan)")})*100,"+0;-0;0")&" pts vs plan"', NAVY),
        ("OVERDUE ACTIVITIES", f'=N({T("Overdue")})', "0", '=k_TaskEscalated&" overdue > "&cfg_EscalationDays&" days (escalate)"', BURG),
        ("OPEN GAPS (BACKLOG)", "=k_GapOpen", "0", '=k_GapOverdue&" overdue · "&k_GapCritOpen&" critical"', BURG),
        ("GAP CLOSURE RATE", '=IF(k_GapTotal=0,0,k_GapClosure)', PCT, '=k_GapClosedV&" verified closed of "&k_GapTotal&" identified"', NAVY),
        ("EVIDENCE VERIFIED", f'=IF(N({T("In Scope")})=0,0,{T("Evidence Verified %")})', PCT,
         f'="Evidence submitted: "&TEXT(N({T("Evidence Submitted %")}),"0%")', NAVY),
        ("BRONZE STAR READINESS", "=k_BronzePct", PCT, '=k_BronzeReady&" of "&k_BronzeMand&" mandatory criteria verified"', "2E7D32"),
    ]
    for i, (lab, fm, fmt, sub, acc) in enumerate(tiles):
        tile(ws, 5, 2 + 2 * i, lab, fm, fmt, sub, 2, acc)
    ws.row_dimensions[6].height = 30
    ws.row_dimensions[7].height = 24
    # Bronze banner
    ws.merge_cells("B9:Q9")
    ws["B9"] = "=k_BronzeBanner"
    ws["B9"].font = f(11, True, WHITE)
    ws["B9"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[9].height = 22
    ws.conditional_formatting.add("B9:Q9", FormulaRule(formula=['LEFT($B$9,5)="BRONZ"'], fill=fill("2E7D32")))
    ws.conditional_formatting.add("B9:Q9", FormulaRule(formula=['LEFT($B$9,5)<>"BRONZ"'], fill=fill(BURG)))
    # Five leadership questions
    section(ws, 11, 2, "Leadership questions – answered from live data", 16)
    qa = [
        ("Where are we?", f'=IF(N({T("Weight")})=0,"No in-scope tasks loaded yet.","Verified completion "&TEXT({T("Verified %")},"0%")&" · completion credit "&TEXT({T("Completion Credit %")},"0%")&" vs "&TEXT({T("Planned %")},"0%")&" planned to date ("&TEXT({T("Variance (Credit−Plan)")}*100,"+0;-0;0")&" pts). Overall status: "&{T("RAG")}&".")'),
        ("What is complete?", f'=N({T("Verified Complete")})&" of "&N({T("In Scope")})&" required tasks verified complete · "&N({T("Pending Verification")})&" complete but awaiting verification · "&k_BronzeReady&" of "&k_BronzeMand&" Bronze criteria verified."'),
        ("What is missing?", f'=N({T("Not Started")})&" tasks not started · "&N({T("Overdue")})&" overdue · "&(N({T("In Scope")})-k_EvSubCount)&" tasks without evidence · "&(k_BronzeMand-k_BronzeReady)&" Bronze criteria not verified ("&k_BronzeMissingEv&" missing evidence) · "&k_DataFlags&" task rows with data-quality flags."'),
        ("Who owns it?", '=IF(k_TopOwner="","No overdue tasks are assigned to a named owner.","Most overdue: "&k_TopOwner&" ("&k_TopOwnerN&" tasks) · "&k_OwnersWithOverdue&" owner(s) with overdue tasks")&" · "&k_Unassigned&" in-scope tasks have no owner."'),
        ("When will we close the gaps?", '=IF(k_GapOpen=0,"No open gaps.",k_GapOpen&" open gaps · latest committed due date "&IF(k_GapMaxDue>0,TEXT(k_GapMaxDue,"dd-mmm-yy"),"not set")&" · "&k_GapNoDue&" without a due date · "&k_GapOverdue&" overdue.")&" Remaining task plan ends "&IF(k_MaxTarget>0,TEXT(k_MaxTarget,"dd-mmm-yy"),"– (no target dates)")&"."'),
    ]
    for i, (q, a) in enumerate(qa):
        r = 12 + i
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=17)
        c = ws.cell(r, 2, q)
        c.font = f(9, True, WHITE)
        c.fill = fill(NAVY)
        c.alignment = Alignment(vertical="center", indent=1)
        c = ws.cell(r, 4, a)
        c.font = f(9, False, "1F1F1F")
        c.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
        for cc in range(2, 18):
            ws.cell(r, cc).border = BORDER
        ws.row_dimensions[r].height = 22
    # Status strip
    section(ws, 18, 2, "Activity status (required, applicable tasks)", 16)
    strip = [("VERIFIED COMPLETE", f"=N({T('Verified Complete')})", "Green"),
             ("COMPLETE – PENDING VERIFICATION", f"=N({T('Pending Verification')})", "Yellow"),
             ("IN PROGRESS", f"=N({T('In Progress')})", "Yellow"),
             ("NOT STARTED", f"=N({T('Not Started')})", "Gray"),
             ("OVERDUE", f"=N({T('Overdue')})", "Red"),
             ("N/A – APPROVED (EXCLUDED)", "=k_NA", "Gray"),
             ("OPTIONAL (NOT REQUIRED)", "=k_Optional", "Gray"),
             ("LEADERSHIP ISSUES", "=k_LeaderIssues", "Red")]
    for i, (lab, fm, rg) in enumerate(strip):
        c0 = 2 + 2 * i
        ws.merge_cells(start_row=19, start_column=c0, end_row=19, end_column=c0 + 1)
        ws.merge_cells(start_row=20, start_column=c0, end_row=20, end_column=c0 + 1)
        a = ws.cell(19, c0, lab)
        a.font = f(7, True, RAG[rg][1])
        a.fill = fill(RAG[rg][0])
        a.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        b = ws.cell(20, c0, fm)
        b.font = f(14, True, RAG[rg][1])
        b.alignment = Alignment(horizontal="center")
        b.number_format = "0"
        for rr in (19, 20):
            for cc in (c0, c0 + 1):
                ws.cell(rr, cc).border = BORDER
    ws.row_dimensions[19].height = 24
    # Pillar table
    r0 = 22
    section(ws, r0, 2, "Completion by DunGPS pillar", 16)
    heads = ["Pillar", "Pillar Name", "In-Scope Tasks", "Verified Tasks", "Verified %", "Completion Credit %",
             "Planned %", "Variance (pts)", "Evidence Submitted %", "Evidence Verified %", "Overdue", "Open Gaps",
             "Critical Gaps", "Bronze Readiness %", "Planned Finish", "RAG"]
    srcs = ["Pillar", "Name", "In Scope", "Verified Complete", "Verified %", "Completion Credit %", "Planned %",
            "Variance (Credit−Plan)", "Evidence Submitted %", "Evidence Verified %", "Overdue", "Open Gaps",
            "Critical Open Gaps", "Bronze Readiness %", "Planned Finish", "RAG"]
    fmts = [None, None, "0", "0", PCT, PCT, PCT, '+0%;-0%;0%', PCT, PCT, "0", "0", "0", PCT, DATE, None]
    hdr_row(ws, r0 + 1, 2, heads)
    ws.row_dimensions[r0 + 1].height = 30
    for i, rr in enumerate(rows_p):
        r = r0 + 2 + i
        total = rr == rows_p[-1]
        for j, s in enumerate(srcs):
            ref = f"{CA}!${col[s]}${rr}"
            c = calc_cell(ws, r, 2 + j, f'=IF({ref}="","",{ref})' if not total else f"={ref}", fmts[j],
                          bold=total, align="left" if j == 1 else "center", size=9)
            c.fill = fill(LGRAY if total else WHITE)
    pr = (r0 + 2, r0 + 2 + len(rows_p) - 1)
    bar_cf(ws, f"F{pr[0]}:F{pr[1]}")
    bar_cf(ws, f"O{pr[0]}:O{pr[1]}", "2E7D32")
    rag_cf(ws, f"Q{pr[0]}:Q{pr[1]}")
    ws.conditional_formatting.add(f"I{pr[0]}:I{pr[1]}", CellIsRule(operator="lessThan", formula=["-cfg_YellowTol"], font=Font(color="C00000", bold=True)))
    ws.conditional_formatting.add(f"L{pr[0]}:L{pr[1]}", CellIsRule(operator="greaterThan", formula=["0"], font=Font(color="C00000", bold=True)))
    # phase table
    r1 = pr[1] + 2
    section(ws, r1, 2, "Completion by deployment phase", 16)
    hdr_row(ws, r1 + 1, 2, [""] + ["Deployment Phase"] + heads[2:])
    ws.row_dimensions[r1 + 1].height = 30
    for i, rr in enumerate(rows_ph):
        r = r1 + 2 + i
        total = rr == rows_ph[-1]
        for j, s in enumerate(srcs):
            if j == 0:
                calc_cell(ws, r, 2, "" if not total else "TOTAL", None, total)
                continue
            if s in ("Open Gaps", "Critical Open Gaps", "Bronze Readiness %"):
                calc_cell(ws, r, 2 + j, "–" if s != "Bronze Readiness %" else "–", None)
                continue
            src = "Pillar" if s == "Name" else s
            ref = f"{CA}!${col[src]}${rr}"
            c = calc_cell(ws, r, 2 + j, f'=IF({ref}="","",{ref})', fmts[j] if s != "Name" else None, total,
                          align="left" if j == 1 else "center")
            if total and s == "Name":
                c.value = "All phases"
    ph = (r1 + 2, r1 + 2 + len(rows_ph) - 1)
    bar_cf(ws, f"F{ph[0]}:F{ph[1]}")
    rag_cf(ws, f"Q{ph[0]}:Q{ph[1]}")
    # Critical issues lists
    r2 = ph[1] + 2
    section(ws, r2, 2, "Critical issues requiring leadership intervention", 16)
    hdr_row(ws, r2 + 1, 2, ["Task ID", "Overdue activity (oldest first)", "Pillar", "Owner", "", "Target", "Days Late"])
    ws.merge_cells(start_row=r2 + 1, start_column=5, end_row=r2 + 1, end_column=6)
    top_list(ws, r2 + 2, 2, 5, "T_KeyOver", [("T_ID", None), ("T_Req", None), ("T_Pillar", None), ("T_Owner", None)])
    for k in range(5):
        r = r2 + 2 + k
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)
        m = f"MATCH(LARGE(T_KeyOver,{k + 1}),T_KeyOver,0)"
        calc_cell(ws, r, 7, f'=IF(LARGE(T_KeyOver,{k + 1})<=0,"",IF(INDEX(T_Target,{m})="","",INDEX(T_Target,{m})))', DATE, size=8)
        calc_cell(ws, r, 8, f'=IF(LARGE(T_KeyOver,{k + 1})<=0,"",IF(INDEX(T_Days,{m})="","",INDEX(T_Days,{m})))', "0", size=8)
        ws.cell(r, 3).alignment = Alignment(wrap_text=True, vertical="center")
    hdr_row(ws, r2 + 1, 10, ["Gap ID", "Critical / leadership gap", "", "", "Sev.", "Owner", "Due", "Age"])
    ws.merge_cells(start_row=r2 + 1, start_column=11, end_row=r2 + 1, end_column=13)
    for k in range(5):
        r = r2 + 2 + k
        m = f"MATCH(LARGE(G_LeadKey,{k + 1}),G_LeadKey,0)"
        g = f"LARGE(G_LeadKey,{k + 1})<=0"
        ws.merge_cells(start_row=r, start_column=11, end_row=r, end_column=13)
        calc_cell(ws, r, 10, f'=IF({g},"",INDEX(G_ID,{m})&"")', None, size=8)
        calc_cell(ws, r, 11, f'=IF({g},"",INDEX(G_Pillar,{m})&": "&INDEX(G_Desc,{m}))', None, size=8, align="left")
        calc_cell(ws, r, 14, f'=IF({g},"",INDEX(G_Sev,{m})&"")', None, size=8)
        calc_cell(ws, r, 15, f'=IF({g},"",INDEX(G_Owner,{m})&"")', None, size=8, align="left")
        calc_cell(ws, r, 16, f'=IF({g},"",IF(INDEX(G_Due,{m})="","",INDEX(G_Due,{m})))', DATE, size=8)
        calc_cell(ws, r, 17, f'=IF({g},"",IF(INDEX(G_Age,{m})="","",INDEX(G_Age,{m})))', "0", size=8)
        ws.row_dimensions[r].height = 24
    text_cf(ws, f"N{r2 + 2}:N{r2 + 6}", {"Critical": "Red", "High": "Yellow"})
    # Charts
    r3 = r2 + 9
    section(ws, r3, 2, "Trends & readiness", 16)
    calc = sh[S_CALC]
    p0, p1 = rows_p[0], rows_p[-2]
    ci = lambda n: list(col).index(n) + 10
    ch = BarChart()
    ch.type = "col"
    ch.grouping = "clustered"
    data = Reference(calc, min_col=ci("Verified %"), min_row=p0 - 1, max_row=p1)
    ch.add_data(data, titles_from_data=True)
    ch.add_data(Reference(calc, min_col=ci("Planned %"), min_row=p0 - 1, max_row=p1), titles_from_data=True)
    ch.add_data(Reference(calc, min_col=ci("Bronze Readiness %"), min_row=p0 - 1, max_row=p1), titles_from_data=True)
    ch.set_categories(Reference(calc, min_col=10, min_row=p0, max_row=p1))
    color_series(ch.series[0], NAVY)
    color_series(ch.series[1], "BFBFBF")
    color_series(ch.series[2], "2E7D32")
    ch.y_axis.scaling.min = 0
    ch.y_axis.scaling.max = 1
    ch.y_axis.number_format = "0%"
    style_chart(ch, "Verified vs planned completion & Bronze readiness by pillar", w=17, h=7.5)
    ws.add_chart(ch, f"B{r3 + 2}")
    t0, t1 = meta["trend"]
    lc = LineChart()
    lc.add_data(Reference(calc, min_col=11, min_row=t0 - 1, max_row=t1), titles_from_data=True)
    lc.add_data(Reference(calc, min_col=12, min_row=t0 - 1, max_row=t1), titles_from_data=True)
    lc.set_categories(Reference(calc, min_col=13, min_row=t0, max_row=t1))
    color_series(lc.series[0], "A6A6A6", True)
    color_series(lc.series[1], BURG, True)
    lc.series[0].graphicalProperties.line.dashStyle = "dash"
    lc.y_axis.scaling.min = 0
    lc.y_axis.number_format = "0%"
    style_chart(lc, "Planned vs actual verified completion (cumulative)", w=15, h=7.5)
    lc.x_axis.tickLblSkip = 4
    ws.add_chart(lc, f"J{r3 + 2}")
    # status doughnut
    dc = DoughnutChart()
    lab_r = r3 + 18
    ws.cell(lab_r, 2, "Status mix (chart data)").font = f(7, False, GRAY)
    names = ["Verified", "Pending Verif.", "In Progress", "Not Started", "Overdue"]
    srcn = ["Verified Complete", "Pending Verification", "In Progress", "Not Started", "Overdue"]
    for i, (n, s) in enumerate(zip(names, srcn)):
        ws.cell(lab_r + 1 + i, 2, n).font = f(7, False, GRAY)
        ws.cell(lab_r + 1 + i, 3, f"=N({T(s)})").font = f(7, False, GRAY)
    dc.add_data(Reference(ws, min_col=3, min_row=lab_r + 1, max_row=lab_r + 5))
    dc.set_categories(Reference(ws, min_col=2, min_row=lab_r + 1, max_row=lab_r + 5))
    from openpyxl.chart.series import DataPoint
    for i, cl in enumerate(["2E7D32", "C9A227", "F2C94C", "BFBFBF", "C62828"]):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = cl
        dc.series[0].dPt.append(pt)
    dc.holeSize = 55
    style_chart(dc, "Activity status mix", w=9, h=7, legend_pos="r")
    ws.add_chart(dc, f"B{r3 + 17}")
    f0, f1 = meta["fwd"]
    fc = BarChart()
    fc.type = "col"
    fc.add_data(Reference(calc, min_col=11, min_row=f0 - 1, max_row=f1), titles_from_data=True)
    fc.add_data(Reference(calc, min_col=12, min_row=f0 - 1, max_row=f1), titles_from_data=True)
    fc.set_categories(Reference(calc, min_col=10, min_row=f0, max_row=f1))
    color_series(fc.series[0], NAVY)
    color_series(fc.series[1], BURG)
    style_chart(fc, "Closure forecast – open tasks & gaps by due month", w=23, h=7)
    ws.add_chart(fc, f"G{r3 + 17}")
    ws.freeze_panes = "A4"
    landscape(ws, f"A1:Q{r3 + 32}", one_page=False)
    ws.page_setup.fitToHeight = 2
    protect(ws)


def build_gapdash(wb, sh, demo):
    ws = sh[S_GDB]
    meta = wb._dunmeta
    calc = sh[S_CALC]
    CF = qs(sh[S_CFG])
    title_band(ws, "Gap Analysis Dashboard", "", 17, legend=False)
    ws["B3"] = ('="Status as of "&TEXT(AsOf,"dd-mmm-yyyy")&"  ·  Closure rate = verified closed gaps ÷ total gaps identified (cumulative). Open backlog shown separately."'
                + ('&"  ·  DEMONSTRATION DATA"' if demo else ""))
    ws.column_dimensions["B"].width = 22
    for c in range(3, 18):
        ws.column_dimensions[L(c)].width = 10
    tiles = [("TOTAL GAPS IDENTIFIED", "=k_GapTotal", "0", '="cumulative, excl. cancelled"', NAVY),
             ("OPEN BACKLOG", "=k_GapOpen", "0", '=k_GapPendVer&" pending verification"', BURG),
             ("OVERDUE GAPS", "=k_GapOverdue", "0", '=k_GapNoDue&" open w/o due date"', BURG),
             ("CRITICAL OPEN", "=k_GapCritOpen", "0", '=k_GapLeader&" need leadership support"', BURG),
             ("CLOSURE RATE (CUM.)", "=IF(k_GapTotal=0,0,k_GapClosure)", PCT, '=k_GapClosedV&" verified closed"', "2E7D32"),
             ("AVG AGE OF OPEN GAPS", '=IF(k_GapAvgAge="",0,k_GapAvgAge)', '0" d"', '="days since identified"', GRAY),
             ("LATEST COMMITTED DUE", '=IF(k_GapMaxDue>0,k_GapMaxDue,"–")', DATE, '="projected backlog closure"', GRAY),
             ("DATA-QUALITY FLAGS", "=k_GapDataFlags", "0", '="gap rows needing correction"', GRAY)]
    ws.column_dimensions["B"].width = 11
    ws.column_dimensions["C"].width = 11
    for i, (lab, fm, fmt, sub, acc) in enumerate(tiles):
        tile(ws, 5, 2 + 2 * i, lab, fm, fmt, sub, 2, acc)
    ws.row_dimensions[6].height = 30
    # pillar x severity table
    r0 = 10
    section(ws, r0, 2, "Open gaps by pillar and severity", 9)
    hdr_row(ws, r0 + 1, 2, ["Pillar", "Critical", "High", "Medium", "Low", "Open Total", "Overdue", "Closed (Verified)",
                            "Closure Rate"])
    for i in range(NPIL):
        r = r0 + 2 + i
        lr = CFG_LIST_ROW0 + i
        P = f"$B{r}"
        calc_cell(ws, r, 2, f'=IF({CF}!$F${lr}="","",{CF}!$F${lr}&"")', None, True)
        for j, s in enumerate(["Critical", "High", "Medium", "Low"]):
            calc_cell(ws, r, 3 + j, f'=IF({P}="","",COUNTIFS(G_Pillar,{P},G_Open,1,G_Sev,"{s}"))', "0")
        calc_cell(ws, r, 7, f'=IF({P}="","",COUNTIFS(G_Pillar,{P},G_Open,1))', "0", True)
        calc_cell(ws, r, 8, f'=IF({P}="","",COUNTIFS(G_Pillar,{P},G_Overdue,1))', "0")
        calc_cell(ws, r, 9, f'=IF({P}="","",COUNTIFS(G_Pillar,{P},G_ClosedV,1))', "0")
        calc_cell(ws, r, 10, f'=IF({P}="","",IF(COUNTIFS(G_Pillar,{P},G_Counted,1)=0,"",I{r}/COUNTIFS(G_Pillar,{P},G_Counted,1)))', PCT)
    rt = r0 + 2 + NPIL
    calc_cell(ws, rt, 2, "TOTAL", None, True)
    for j in range(3, 10):
        calc_cell(ws, rt, j, f"=SUM({L(j)}{r0 + 2}:{L(j)}{rt - 1})", "0", True)
    calc_cell(ws, rt, 10, '=IF(k_GapTotal=0,"",k_GapClosure)', PCT, True)
    ws.conditional_formatting.add(f"C{r0 + 2}:C{rt - 1}", CellIsRule(operator="greaterThan", formula=["0"], fill=fill(RAG["Red"][0]), font=Font(color=RAG["Red"][1], bold=True)))
    ws.conditional_formatting.add(f"H{r0 + 2}:H{rt - 1}", CellIsRule(operator="greaterThan", formula=["0"], font=Font(color="C00000", bold=True)))
    bar_cf(ws, f"J{r0 + 2}:J{rt - 1}", "2E7D32")
    # overdue by owner (top 10)
    section(ws, r0, 12, "Overdue gaps by owner (top 10)", 6)
    hdr_row(ws, r0 + 1, 12, ["Owner", "", "", "Overdue", "Open", "Verified Tasks %"])
    ws.merge_cells(start_row=r0 + 1, start_column=12, end_row=r0 + 1, end_column=14)
    o_first, o_last, oc = meta["owners"]
    CA = qs(calc)
    for k in range(10):
        r = r0 + 2 + k
        m = f"MATCH(LARGE(c_OwnGapKey,{k + 1}),c_OwnGapKey,0)"
        g = f"LARGE(c_OwnGapKey,{k + 1})<=0"
        ws.merge_cells(start_row=r, start_column=12, end_row=r, end_column=14)
        calc_cell(ws, r, 12, f'=IF({g},"",INDEX(c_OwnName,{m}))', None, align="left")
        calc_cell(ws, r, 15, f'=IF({g},"",INDEX({CA}!${oc["Overdue Gaps"]}${o_first}:${oc["Overdue Gaps"]}${o_last},{m}))', "0", True)
        calc_cell(ws, r, 16, f'=IF({g},"",INDEX({CA}!${oc["Open Gaps"]}${o_first}:${oc["Open Gaps"]}${o_last},{m}))', "0")
        calc_cell(ws, r, 17, f'=IF({g},"",INDEX({CA}!${oc["Verified %"]}${o_first}:${oc["Verified %"]}${o_last},{m}))', PCT)
    r = r0 + 12
    calc_cell(ws, r, 12, "(Unassigned open gaps)", None, align="left")
    ws.merge_cells(start_row=r, start_column=12, end_row=r, end_column=14)
    calc_cell(ws, r, 15, '=COUNTIFS(G_Overdue,1,G_Owner,"")', "0", True)
    calc_cell(ws, r, 16, '=COUNTIFS(G_Open,1,G_Owner,"")', "0")
    # severity & aging small tables
    r1 = rt + 2
    section(ws, r1, 2, "Gaps by severity", 5)
    hdr_row(ws, r1 + 1, 2, ["Severity", "Open", "Overdue", "Closed (Verified)", "Total"])
    s0, s1 = meta["sev"]
    for k in range(5):
        for j in range(5):
            calc_cell(ws, r1 + 2 + k, 2 + j, f"={CA}!{L(10 + j)}{s0 + k}", None if j == 0 else "0", j == 0)
    text_cf(ws, f"B{r1 + 2}:B{r1 + 3}", {"Critical": "Red", "High": "Yellow"})
    section(ws, r1, 8, "Open gap aging", 3)
    hdr_row(ws, r1 + 1, 8, ["Age Bucket", "", "Open Gaps"])
    a0, a1 = meta["aging"]
    for k in range(5):
        ws.merge_cells(start_row=r1 + 2 + k, start_column=8, end_row=r1 + 2 + k, end_column=9)
        calc_cell(ws, r1 + 2 + k, 8, f"={CA}!J{a0 + k}", None, True, align="left")
        calc_cell(ws, r1 + 2 + k, 10, f"={CA}!K{a0 + k}", "0")
    # Pareto table / recurring barriers
    section(ws, r1, 12, "Top recurring barriers (gaps + task barriers)", 6)
    hdr_row(ws, r1 + 1, 12, ["Category", "", "", "Count", "Cum. %", "Share"])
    ws.merge_cells(start_row=r1 + 1, start_column=12, end_row=r1 + 1, end_column=14)
    p0, p1 = meta["pareto"]
    for k in range(5):
        r = r1 + 2 + k
        ws.merge_cells(start_row=r, start_column=12, end_row=r, end_column=14)
        calc_cell(ws, r, 12, f'=IF({CA}!Q{p0 + k}=0,"",{CA}!P{p0 + k})', None, align="left")
        calc_cell(ws, r, 15, f'=IF({CA}!Q{p0 + k}=0,"",{CA}!Q{p0 + k})', "0", True)
        calc_cell(ws, r, 16, f'=IF({CA}!Q{p0 + k}=0,"",{CA}!R{p0 + k})', PCT)
        calc_cell(ws, r, 17, f'=IF(OR({CA}!Q{p0 + k}=0,SUM({CA}!$Q${p0}:$Q${p0 + 19})=0),"",{CA}!Q{p0 + k}/SUM({CA}!$Q${p0}:$Q${p0 + 19}))', PCT)
    bar_cf(ws, f"Q{r1 + 2}:Q{r1 + 6}", BURG)
    # top 5 gaps
    r2 = r1 + 9
    section(ws, r2, 2, "Top 5 open gaps (severity, then age)", 16)
    hdr_row(ws, r2 + 1, 2, ["Gap ID", "Pillar", "Severity", "Gap description", "", "", "", "", "Root cause", "", "",
                            "Owner", "", "Due", "Age", "Status"])
    ws.merge_cells(start_row=r2 + 1, start_column=5, end_row=r2 + 1, end_column=9)
    ws.merge_cells(start_row=r2 + 1, start_column=10, end_row=r2 + 1, end_column=12)
    ws.merge_cells(start_row=r2 + 1, start_column=13, end_row=r2 + 1, end_column=14)
    for k in range(5):
        r = r2 + 2 + k
        m = f"MATCH(LARGE(G_Key,{k + 1}),G_Key,0)"
        g = f"LARGE(G_Key,{k + 1})<=0"
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=9)
        ws.merge_cells(start_row=r, start_column=10, end_row=r, end_column=12)
        ws.merge_cells(start_row=r, start_column=13, end_row=r, end_column=14)
        for c, src, fmt, al in [(2, "G_ID", None, "center"), (3, "G_Pillar", None, "center"), (4, "G_Sev", None, "center"),
                                (5, "G_Desc", None, "left"), (10, "G_Root", None, "left"), (13, "G_Owner", None, "left"),
                                (15, "G_Due", DATE, "center"), (16, "G_Age", "0", "center"), (17, "G_Status", None, "left")]:
            fm = f'=IF({g},"",INDEX({src},{m})&"")' if fmt is None else f'=IF({g},"",IF(INDEX({src},{m})="","",INDEX({src},{m})))'
            calc_cell(ws, r, c, fm, fmt, align=al, size=8)
        ws.row_dimensions[r].height = 26
    text_cf(ws, f"D{r2 + 2}:D{r2 + 6}", {"Critical": "Red", "High": "Yellow", "Medium": "Gray"})
    # charts
    r3 = r2 + 8
    section(ws, r3, 2, "Gap analytics", 16)
    st = BarChart()
    st.type = "bar"
    st.grouping = "stacked"
    st.overlap = 100
    for j, cl in zip(range(3, 7), ["C62828", "E69F00", "BFBFBF", "D9E1F2"]):
        st.add_data(Reference(ws, min_col=j, min_row=r0 + 1, max_row=r0 + 1 + NPIL), titles_from_data=True)
        color_series(st.series[-1], cl)
    st.set_categories(Reference(ws, min_col=2, min_row=r0 + 2, max_row=r0 + 1 + NPIL))
    style_chart(st, "Open gaps by pillar & severity", w=12, h=8)
    st.x_axis.scaling.orientation = "maxMin"
    ws.add_chart(st, f"B{r3 + 2}")
    # Pareto combo
    pc = BarChart()
    pc.add_data(Reference(calc, min_col=17, min_row=p0 - 1, max_row=p1), titles_from_data=True)
    pc.set_categories(Reference(calc, min_col=16, min_row=p0, max_row=p1))
    color_series(pc.series[0], NAVY)
    ln = LineChart()
    ln.add_data(Reference(calc, min_col=18, min_row=p0 - 1, max_row=p1), titles_from_data=True)
    color_series(ln.series[0], BURG, True)
    ln.y_axis.axId = 200
    ln.y_axis.scaling.min = 0
    ln.y_axis.scaling.max = 1
    ln.y_axis.number_format = "0%"
    ln.y_axis.crosses = "max"
    ln.y_axis.delete = False
    pc += ln
    style_chart(pc, "Pareto – gap / barrier categories", w=15, h=8)
    ws.add_chart(pc, f"H{r3 + 2}")
    m0, m1 = meta["monthly"]
    nc = BarChart()
    nc.type = "col"
    nc.add_data(Reference(calc, min_col=11, min_row=m0 - 1, max_row=m1), titles_from_data=True)
    nc.add_data(Reference(calc, min_col=12, min_row=m0 - 1, max_row=m1), titles_from_data=True)
    nc.set_categories(Reference(calc, min_col=10, min_row=m0, max_row=m1))
    color_series(nc.series[0], BURG)
    color_series(nc.series[1], "2E7D32")
    bl = LineChart()
    bl.add_data(Reference(calc, min_col=13, min_row=m0 - 1, max_row=m1), titles_from_data=True)
    color_series(bl.series[0], NAVY, True)
    bl.y_axis.axId = 300
    bl.y_axis.crosses = "max"
    bl.y_axis.delete = False
    nc += bl
    style_chart(nc, "New vs verified-closed gaps per month & open backlog", w=14, h=8)
    ws.add_chart(nc, f"B{r3 + 18}")
    ac = BarChart()
    ac.type = "col"
    ac.add_data(Reference(calc, min_col=11, min_row=a0 - 1, max_row=a1), titles_from_data=True)
    ac.set_categories(Reference(calc, min_col=10, min_row=a0, max_row=a1))
    color_series(ac.series[0], GRAY)
    style_chart(ac, "Open gap aging", w=13, h=8)
    ac.legend = None
    ws.add_chart(ac, f"J{r3 + 18}")
    ws.freeze_panes = "A4"
    landscape(ws, f"A1:Q{r3 + 34}", one_page=False)
    ws.page_setup.fitToHeight = 2
    protect(ws)


def build_mr(wb, sh, demo):
    ws = sh[S_MR]
    meta = wb._dunmeta
    calc = sh[S_CALC]
    CA = qs(calc)
    col, rows_pf = meta["col"], meta["rows_pf"]
    title_band(ws, "Management Review — Weekly DunGPS Deployment", "", 18, legend=False)
    ws["B3"] = ('=cfg_SiteName&"  ·  Status as of "&TEXT(AsOf,"dd-mmm-yyyy")&"  ·  Filters apply to tasks (all 5), gaps (pillar/owner/area) and VO projects (pillar/owner/area)."'
                + ('&"  ·  DEMONSTRATION DATA"' if demo else ""))
    widths = [2, 9, 9, 9, 9, 9, 9, 9, 9, 8, 2, 9, 9, 9, 26, 15, 9, 8]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[L(i)].width = w
    # filters
    flt = [("Pillar", "f_Pillar", "lst_fPillar", 2), ("Owner", "f_Owner", "lst_fOwner", 4),
           ("Process / Area", "f_Area", "lst_fArea", 7), ("Shift", "f_Shift", "lst_fShift", 12),
           ("Phase", "f_Phase", "lst_fPhase", 15)]
    for lab, nm, lst, c in flt:
        lc = ws.cell(5, c, lab + " ▼")
        lc.font = f(8, True, NAVY)
        span = {2: 1, 4: 2, 7: 2, 12: 2, 15: 1}[c]
        vc = ws.cell(6, c, "(All)")
        ws.merge_cells(start_row=6, start_column=c, end_row=6, end_column=c + span)
        vc.font = f(10, True, "0000FF")
        vc.fill = fill("FFF9E5")
        vc.alignment = Alignment(horizontal="center")
        for cc in range(c, c + span + 1):
            ws.cell(6, cc).border = Border(left=thin, right=thin, top=Side(style="medium", color=NAVY),
                                           bottom=Side(style="medium", color=NAVY))
        dv = DataValidation(type="list", formula1=f"={lst}", allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(ws.cell(6, c).coordinate)
        define(wb, nm, f"{qs(ws)}!${L(c)}$6")
    ws.cell(5, 17, "Select (All) to clear").font = f(7, False, GRAY, True)
    TF = lambda n: "totf_" + "".join(ch for ch in n if ch.isalnum())
    tiles = [("VERIFIED COMPLETION", f'=IF(N({TF("Weight")})=0,0,{TF("Verified %")})', PCT,
              f'=N({TF("Verified Complete")})&" / "&N({TF("In Scope")})&" tasks"', NAVY, 2),
             ("PLANNED vs CREDIT", f'=IF(N({TF("Weight")})=0,0,{TF("Planned %")})', PCT,
              f'="Credit "&TEXT(N({TF("Completion Credit %")}),"0%")&" ("&TEXT(N({TF("Variance (Credit−Plan)")})*100,"+0;-0;0")&" pts)"', NAVY, 2),
             ("OVERDUE ACTIONS", f'=N({TF("Overdue")})', "0", '="tasks past target"', BURG, 2),
             ("OPEN GAPS", "=k_GapOpenF", "0", '=k_GapOverdueF&" overdue"', BURG, 2),
             ("BRONZE READINESS", f'=IF(N({TF("Bronze Mandatory")})=0,0,{TF("Bronze Readiness %")})', PCT,
              f'=N({TF("Bronze Verified")})&" / "&N({TF("Bronze Mandatory")})&" criteria"', "2E7D32", 2),
             ("ACTIVE VO PROJECTS", "=k_VOActiveF", "0", '="Est. opp. "&TEXT(k_VOOppF,"$#,##0")', GRAY, 3),
             ("FINANCE-VERIFIED SAVINGS", "=k_VOFinSavF", "$#,##0", '="validated by Finance"', "2E7D32", 3)]
    c = 2
    for (lab, fm, fmt, sub, acc, _w), w in zip(tiles, [2, 2, 2, 2, 2, 3, 4]):
        tile(ws, 8, c, lab, fm, fmt, sub, w, acc)
        c += w
    ws.row_dimensions[9].height = 28
    # pillar table filtered (B..J)
    r0 = 12
    section(ws, r0, 2, "Pillar completion (filtered)", 9)
    hdr_row(ws, r0 + 1, 2, ["Pillar", "In-Scope", "Verified %", "Planned %", "Variance", "Overdue", "Open Gaps",
                            "Bronze %", "RAG"])
    srcs = ["Pillar", "In Scope", "Verified %", "Planned %", "Variance (Credit−Plan)", "Overdue", "Open Gaps",
            "Bronze Readiness %", "RAG"]
    fmts = [None, "0", PCT, PCT, '+0%;-0%;0%', "0", "0", PCT, None]
    for i, rr in enumerate(rows_pf):
        r = r0 + 2 + i
        total = rr == rows_pf[-1]
        for j, s in enumerate(srcs):
            ref = f"{CA}!${col[s]}${rr}"
            cc = calc_cell(ws, r, 2 + j, f'=IF({ref}="","",{ref})', fmts[j], total, size=8)
            if total:
                cc.fill = fill(LGRAY)
    pe = r0 + 2 + len(rows_pf) - 1
    bar_cf(ws, f"D{r0 + 2}:D{pe}")
    rag_cf(ws, f"J{r0 + 2}:J{pe}")
    # top 5 gaps (L..R)
    section(ws, r0, 12, "Top 5 gaps (filtered)", 7)
    hdr_row(ws, r0 + 1, 12, ["Gap ID", "Pillar", "Sev.", "Description", "Owner", "Due", "Age"])
    for k in range(5):
        r = r0 + 2 + k
        m = f"MATCH(LARGE(G_KeyF,{k + 1}),G_KeyF,0)"
        g = f"LARGE(G_KeyF,{k + 1})<=0"
        for cc, src, fmt, al in [(12, "G_ID", None, "center"), (13, "G_Pillar", None, "center"), (14, "G_Sev", None, "center"),
                                 (15, "G_Desc", None, "left"), (16, "G_Owner", None, "left"), (17, "G_Due", DATE, "center"),
                                 (18, "G_Age", "0", "center")]:
            fm = f'=IF({g},"",INDEX({src},{m})&"")' if fmt is None else f'=IF({g},"",IF(INDEX({src},{m})="","",INDEX({src},{m})))'
            calc_cell(ws, r, cc, fm, fmt, align=al, size=8)
        ws.row_dimensions[r].height = 24
    text_cf(ws, f"N{r0 + 2}:N{r0 + 6}", {"Critical": "Red", "High": "Yellow"})
    # leadership decisions (L..R) below top gaps
    rL = r0 + 8
    section(ws, rL, 12, "Decisions requiring leadership support (filtered)", 7)
    hdr_row(ws, rL + 1, 12, ["Gap ID", "Pillar", "Sev.", "Support / decision required", "Owner", "Due", "Age"])
    for k in range(5):
        r = rL + 2 + k
        m = f"MATCH(LARGE(G_LeadKeyF,{k + 1}),G_LeadKeyF,0)"
        g = f"LARGE(G_LeadKeyF,{k + 1})<=0"
        for cc, src, fmt, al in [(12, "G_ID", None, "center"), (13, "G_Pillar", None, "center"), (14, "G_Sev", None, "center"),
                                 (15, "G_DecisionTxt", None, "left"), (16, "G_Owner", None, "left"), (17, "G_Due", DATE, "center"),
                                 (18, "G_Age", "0", "center")]:
            fm = f'=IF({g},"",INDEX({src},{m})&"")' if fmt is None else f'=IF({g},"",IF(INDEX({src},{m})="","",INDEX({src},{m})))'
            calc_cell(ws, r, cc, fm, fmt, align=al, size=8)
        ws.row_dimensions[r].height = 24
    text_cf(ws, f"N{rL + 2}:N{rL + 6}", {"Critical": "Red", "High": "Yellow"})
    # overdue actions (B..J)
    r1 = pe + 2
    section(ws, r1, 2, "Overdue actions (filtered, most late first)", 9)
    hdr_row(ws, r1 + 1, 2, ["Task ID", "Pillar", "Requirement / step", "", "", "", "Owner", "", "Days Late"])
    ws.merge_cells(start_row=r1 + 1, start_column=4, end_row=r1 + 1, end_column=7)
    ws.merge_cells(start_row=r1 + 1, start_column=8, end_row=r1 + 1, end_column=9)
    for k in range(8):
        r = r1 + 2 + k
        m = f"MATCH(LARGE(T_KeyOverF,{k + 1}),T_KeyOverF,0)"
        g = f"LARGE(T_KeyOverF,{k + 1})<=0"
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=7)
        ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=9)
        for cc, src, fmt, al in [(2, "T_ID", None, "center"), (3, "T_Pillar", None, "center"), (4, "T_Req", None, "left"),
                                 (8, "T_Owner", None, "left"), (10, "T_Days", "0", "center")]:
            fm = f'=IF({g},"",INDEX({src},{m})&"")' if fmt is None else f'=IF({g},"",IF(INDEX({src},{m})="","",INDEX({src},{m})))'
            calc_cell(ws, r, cc, fm, fmt, align=al, size=8)
    ws.conditional_formatting.add(f"J{r1 + 2}:J{r1 + 9}", CellIsRule(operator="greaterThan", formula=["cfg_EscalationDays"], fill=fill(RAG["Red"][0]), font=Font(color=RAG["Red"][1], bold=True)))
    # owner accountability (L..R)
    rO = rL + 8
    section(ws, rO, 12, "Owner accountability (filtered, most overdue first)", 7)
    hdr_row(ws, rO + 1, 12, ["Owner", "", "", "", "Tasks", "Verified %", "Overdue"])
    ws.merge_cells(start_row=rO + 1, start_column=12, end_row=rO + 1, end_column=15)
    o_first, o_last, oc = meta["owners"]
    rngc = lambda cn: f"{CA}!${oc[cn]}${o_first}:${oc[cn]}${o_last}"
    for k in range(8):
        r = rO + 2 + k
        m = f"MATCH(LARGE(c_OwnKeyF,{k + 1}),c_OwnKeyF,0)"
        g = f"LARGE(c_OwnKeyF,{k + 1})<=0"
        ws.merge_cells(start_row=r, start_column=12, end_row=r, end_column=15)
        calc_cell(ws, r, 12, f'=IF({g},"",INDEX(c_OwnName,{m})&"  ("&INDEX({rngc("F: Open Gaps")},{m})&" open gaps)")', None, align="left", size=8)
        calc_cell(ws, r, 16, f'=IF({g},"",INDEX({rngc("F: Tasks")},{m}))', "0", size=8)
        calc_cell(ws, r, 17, f'=IF({g},"",INDEX({rngc("F: Verified %")},{m}))', PCT, size=8)
        calc_cell(ws, r, 18, f'=IF({g},"",INDEX({rngc("F: Overdue")},{m}))', "0", size=8)
    ws.conditional_formatting.add(f"R{rO + 2}:R{rO + 9}", CellIsRule(operator="greaterThan", formula=["0"], font=Font(color="C00000", bold=True)))
    # VO status (B..J)
    r2 = r1 + 11
    section(ws, r2, 2, "VO & Tiger Team project status (filtered)", 9)
    hdr_row(ws, r2 + 1, 2, ["Status", "", "", "Projects", "Est. Opportunity", "", "Finance-Verified", "", ""])
    ws.merge_cells(start_row=r2 + 1, start_column=2, end_row=r2 + 1, end_column=4)
    ws.merge_cells(start_row=r2 + 1, start_column=6, end_row=r2 + 1, end_column=7)
    ws.merge_cells(start_row=r2 + 1, start_column=8, end_row=r2 + 1, end_column=10)
    v0, v1 = meta["vo"]
    for k in range(v1 - v0 + 1):
        r = r2 + 2 + k
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7)
        ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=10)
        calc_cell(ws, r, 2, f"={CA}!J{v0 + k}", None, align="left", size=8)
        calc_cell(ws, r, 5, f"={CA}!N{v0 + k}", "0", size=8)
        calc_cell(ws, r, 6, f"={CA}!O{v0 + k}", "$#,##0;($#,##0);-", size=8)
        calc_cell(ws, r, 8, f"={CA}!P{v0 + k}", "$#,##0;($#,##0);-", size=8)
    # meeting decisions log (L..R)
    rD = rO + 11
    section(ws, rD, 12, "Meeting decisions / support agreed (type here)", 7)
    hdr_row(ws, rD + 1, 12, ["Date", "Decision / action", "", "", "Owner", "Due", "Status"])
    ws.merge_cells(start_row=rD + 1, start_column=13, end_row=rD + 1, end_column=15)
    unlocked = []
    for k in range(6):
        r = rD + 2 + k
        ws.merge_cells(start_row=r, start_column=13, end_row=r, end_column=15)
        for cc in range(12, 19):
            ws.cell(r, cc).border = BORDER
            ws.cell(r, cc).font = f(8)
        ws.cell(r, 12).number_format = DATE
        ws.cell(r, 17).number_format = DATE
        unlocked.append(f"L{r}:R{r}")
    dvo = DataValidation(type="list", formula1="=lst_Owners", allow_blank=True)
    ws.add_data_validation(dvo)
    dvo.add(f"P{rD + 2}:P{rD + 7}")
    last_row = max(r2 + 2 + (v1 - v0), rD + 7) + 1
    ws.cell(last_row + 1, 2, "Definitions: Verified % = verified required tasks ÷ applicable required tasks. Credit caps "
                             "unverified 'Complete' at the configured maximum. Bronze % = verified mandatory criteria ÷ "
                             "mandatory applicable criteria.").font = f(7, False, GRAY, True)
    landscape(ws, f"A1:R{last_row + 1}", one_page=True)
    protect(ws, ["B6:R6"] + unlocked)


def build_instructions(wb, sh, demo):
    ws = sh[S_INS]
    title_band(ws, "Instructions & Data Definitions", "How the DunGPS Deployment Management System works, "
               "who updates what, and how every number is calculated.", 6, legend=False)
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 110
    rows = [
        ("H", "PURPOSE"),
        ("The system answers five leadership questions", "Where are we? What is complete? What is missing? Who owns it? When will we close the gaps? – see the Executive Dashboard 'Leadership questions' block."),
        ("Status of this file", "DEMONSTRATION DATA – fictitious owners, dates and links for testing. Use the production file for real work." if demo else
         "PRODUCTION – pre-loaded only with the 32 official Bronze Level requirements from the CHG source document, the 9 pillars named in it, and the Pre-Weigh project (baseline pending). No owners, dates or savings are invented."),
        ("H", "SHEET MAP"),
        ("Executive Dashboard", "Site-level KPIs, pillar and phase completion, Bronze status banner, critical issues, trends and closure forecast. Read-only."),
        ("Pillar Deployment Tracker", "MASTER TABLE. One row per deployment task/step. Owners update status, progress, evidence link; verifiers set Verification Status."),
        ("Process Deployment Matrix", "Area × equipment × shift deployment of 9 elements (safety mapping, CILT, centerlines…). Nothing is assumed applicable: default 'Not Assessed'."),
        ("Gap Analysis Dashboard", "Open backlog vs cumulative closure rate, severity, aging, owners, Pareto, new vs closed trend. Read-only."),
        ("Gap Register", "One row per gap: current vs required condition, root cause, corrective action, owner, due date, evidence, verification, closure."),
        ("Owner Directory", "Every owner dropdown in the workbook reads from here. Only 'Active' people appear in dropdowns."),
        ("Bronze Star Readiness", "Official Bronze criteria (verbatim). Roll-up of linked tasks and gaps, evidence, formal verification and readiness result."),
        ("VO & Tiger Team Projects", "Charter, sponsor, team, baseline/target/actual, estimated opportunity vs Finance-verified savings, sustainability checks."),
        ("Document & Evidence Register", "Links to controlled documents in Teams → DunGPS → General → Shared → <Pillar>; revision and review cycle."),
        ("Management Review", "One-page weekly review. Filter by pillar, owner, process/area, shift and phase (yellow cells). Print landscape on one page."),
        ("Configuration & Dropdown Lists", "Thresholds (RAG, escalation, aging, credit cap, weights) and all dropdown lists. Administrator only."),
        ("Calc Engine", "System calculations feeding all dashboards. Protected (no password) – do not edit."),
        ("H", "WEEKLY OPERATING CADENCE"),
        ("1. Owners (by Wednesday)", "Update Status, Work Progress %, Actual Completion, Evidence Link, Gap/Barrier and Next Action; set Last Updated."),
        ("2. Verifiers (by Thursday)", "Open evidence, set Verification Status = Verified (or Rejected - Rework) and Verification Date. Only verified tasks count as complete."),
        ("3. Pillar leads", "Log new gaps in the Gap Register; review Data Quality Flags; correct missing owners/dates."),
        ("4. Management review (Friday)", "Open Management Review, filter as needed, review top gaps, overdue actions and decisions; record decisions in the Meeting decisions log. Print one page."),
        ("H", "INPUT RULES"),
        ("Cell colours", "White cells (navy header) = inputs. Gray cells (gray header, 'ƒ' prefix) = formulas – do not type over them. Dashboards and Calc Engine are protected (no password; Review → Unprotect Sheet)."),
        ("Adding rows", "Tables are pre-formatted (tracker 500, gaps 300, documents 300, owners 200, projects 100, process rows 80). Type in the next empty row. To extend, drag the table's bottom-right handle; calculated columns copy automatically and dashboards use full-table ranges."),
        ("Adding owners", "Type the person on Owner Directory with Active/Inactive = Active. They appear in every owner dropdown immediately. Set Inactive to remove from dropdowns (history kept)."),
        ("Adding pillars / areas / lists", "Add values at the END of the list on Configuration & Dropdown Lists (no blank gaps inside a list). Dashboards show up to 12 pillars. Do not rename gray-header SYSTEM values – formulas test for them."),
        ("Links", "Paste the full Teams/SharePoint URL (… → Copy link) into Evidence/Document Link. The 'Open ↗' column makes it clickable."),
        ("H", "CALCULATION DEFINITIONS"),
        ("In scope (denominator)", "Pillar entered AND Required? ≠ No AND Applicability ≠ 'N/A - Approved'. N/A-approved and optional tasks are counted and shown separately, never in the denominator. 'N/A - Requested' stays in the denominator until approved."),
        ("Verified (task)", "Status = Complete AND Evidence Link present AND Verification Status = Verified. Selecting 'Complete' alone never counts as verified."),
        ("Verified Completion %", "Σ(weight × verified) ÷ Σ(weight) over in-scope tasks. With weights OFF (default) = Verified required tasks ÷ Total applicable required tasks × 100. Calculated at task, Bronze requirement (step), pillar, phase and overall level."),
        ("Completion Credit %", "Verified = 100%; Complete-unverified = 'Max credit before verification' (default 80%); otherwise reported progress capped at the same maximum. Used for schedule RAG."),
        ("Work Progress %", "Owner-reported progress (Complete = 100%), uncapped. Shown for transparency only."),
        ("Evidence Submitted / Verified %", "In-scope tasks with an evidence link ÷ in-scope tasks; and those whose evidence is Verified ÷ in-scope tasks."),
        ("Planned % to date", "Linear plan from Start to Target date at the status date (no start date: 0% until target, then 100%). Planned vs actual trend uses target dates vs verified dates over 26 periods."),
        ("Overdue", "Target date < status date and not verified (N/A-approved excluded). Escalation when days late > configured threshold."),
        ("Task RAG", "Gray = N/A or not started and not yet planned · Green = verified, or credit within Green tolerance of plan · Yellow = within Yellow tolerance, On Hold, or verification rejected · Red = overdue or further behind plan."),
        ("Pillar / phase RAG", "Gray = no scope or nothing started/planned · Red = any critical open gap or overdue share ≥ threshold, or behind plan beyond Yellow tolerance · Yellow = behind plan within tolerance or any overdue task · Green = on plan, nothing overdue."),
        ("Gap closed", "Status = Closed - Verified AND Verification = Verified AND Closure Date entered. Otherwise the gap stays in the open backlog."),
        ("Gap Closure Rate", "Verified closed gaps ÷ total gaps identified (excl. cancelled) × 100 – cumulative. The open backlog is reported separately."),
        ("Gap aging", "Open gaps: status date − date identified. Closed gaps: closure − identified. Buckets configurable."),
        ("Bronze readiness", "Per criterion 'Verified - Ready' only if Verification = Verified, evidence linked, all linked tasks verified and no open gaps. Site banner says READY only when every mandatory, applicable criterion is Verified - Ready."),
        ("VO savings", "Estimated opportunity and Finance-verified savings are separate. Savings count only when Finance Approval = Approved."),
        ("Process matrix score", "Not Started 0 · In Progress 25% · Implemented 75% · Verified 100% (configurable), averaged over applicable elements. Coverage = assessed ÷ 9 elements."),
        ("Document complete", "Status Approved/Verified AND link present AND review date entered AND not past next review date."),
        ("H", "TECHNICAL NOTES"),
        ("Excel version", "Built with Excel 2010+ functions (SUMPRODUCT, COUNTIFS, INDEX/MATCH, EDATE, EOMONTH, HYPERLINK). No macros, Power Query, Power Pivot or Microsoft 365-only functions are required. Works in Excel desktop and Excel for the web (SharePoint/Teams)."),
        ("Status date", "Uses TODAY() unless 'Status date override' is set on the Configuration sheet."),
        ("Recalculation", "Automatic. If numbers look stale press F9."),
        ("H", "INFORMATION REQUIRED FROM CHG TO CONFIGURE THE OFFICIAL SYSTEM"),
        ("1. Official pillar names", "Full name for 'IM' and confirmation of AM, PM, FSQ, EHS, LEAD, E&T, VO, SC expansions; whether EEM is in scope."),
        ("2. Bronze criteria detail", "For each of the 32 criteria: mandatory vs optional, acceptance criteria, required evidence (official), verifier and any approved exceptions (e.g. PM org structure)."),
        ("3. Silver / Gold criteria", "Requirements for later levels so phases 2–4 can be loaded."),
        ("4. Owners", "Names, departments, positions, shifts, emails and pillar responsibilities; pillar leads, verifiers and escalation owners."),
        ("5. Dates", "Program launch date, Bronze assessment target date and step-level target dates per pillar."),
        ("6. Weights", "Whether CHG weights steps/requirements (and the weights)."),
        ("7. Process & equipment list", "Confirmed areas, line/equipment names (e.g. Gemini, Nugget, Rolled Dough lines) and shifts; which elements apply to each."),
        ("8. Teams/SharePoint", "Confirmed folder structure under DunGPS → General → Shared and the site URL for links."),
        ("9. VO / Finance rules", "Savings validation rules, Finance approver, and baseline for the Pre-Weigh Room project."),
        ("10. RAG thresholds", "Confirm or replace default tolerances (5 / 15 pts, 10% overdue, 14-day escalation)."),
    ]
    r = 5
    for a, b in rows:
        if a == "H":
            section(ws, r, 2, b, 2)
            r += 1
            continue
        ws.cell(r, 2, a).font = f(9, True, NAVY)
        c = ws.cell(r, 3, b)
        c.font = f(9)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(r, 2).alignment = Alignment(vertical="top", wrap_text=True)
        ws.row_dimensions[r].height = max(15, 13 * (len(b) // 120 + 1))
        r += 1


# ------------------------------------------------------------------------------------------
# DEMO DATA (clearly fictitious)
def demo_tasks(bronze_ids, as_of):
    owners = [f"DEMO Owner {i:02d}" for i in range(1, 15)]
    pl = {"AM": 3, "PM": 4, "FSQ": 5, "EHS": 6, "LEAD": 1, "E&T": 7, "VO": 8, "IM": 2, "SC": 9}
    areas = ["Pre-weigh", "Mixing", "Dough Processing / Forming", "Gemini Equipment", "Nugget Line",
             "Rolled Dough Lines", "Baking", "Cooling / Freezing", "Packaging", "Sanitation"]
    out = []
    n = 0
    for bid, p, req, sug in bronze_ids:
        for s in range(random.choice([2, 2, 3])):
            n += 1
            start = as_of - dt.timedelta(days=random.randint(20, 160))
            target = start + dt.timedelta(days=random.randint(30, 140))
            u = random.random()
            row = dict(ID=f"T-{p}-{n:03d}", Pillar=p, Phase="1 - Bronze",
                       Req=f"{['Plan', 'Pilot', 'Roll out'][s]}: {req[:60]}", BronzeRef=bid, Required="Yes",
                       Applic="Applicable", Weight=1, Area=random.choice(areas) if p in ("AM", "VO", "FSQ") else "Site-wide",
                       Shift=random.choice(["A", "B", "C", "All"]), Owner=f"DEMO Owner {pl[p]:02d}",
                       SupOwner=random.choice(owners), Team="DEMO cross-functional team",
                       Verifier="DEMO Owner 13", Escal="DEMO Owner 02", Start=start, Target=target,
                       EvidReq=sug[:80], Updated=as_of - dt.timedelta(days=random.randint(0, 14)))
            if u < 0.35:
                done = min(as_of - dt.timedelta(days=3), target + dt.timedelta(days=random.randint(-20, 10)))
                row.update(Status="Complete", Progress=1, Actual=done, VerStatus="Verified",
                           VerDate=done + dt.timedelta(days=2),
                           EvidLink=f"https://example.sharepoint.com/sites/DunGPS/Shared/{p}/DEMO_{n:03d}.pdf")
            elif u < 0.47:
                row.update(Status="Complete", Progress=1, Actual=as_of - dt.timedelta(days=5),
                           VerStatus="Submitted - Pending Review",
                           EvidLink=f"https://example.sharepoint.com/sites/DunGPS/Shared/{p}/DEMO_{n:03d}.pdf")
            elif u < 0.80:
                row.update(Status="In Progress", Progress=round(random.choice([0.2, 0.4, 0.5, 0.6, 0.75]), 2),
                           VerStatus="Not Submitted", BarrierCat=random.choice(
                               ["Resources / time", "Training / skills", "Data / system (CMMS, Redzone)",
                                "Standard / procedure missing"]), Gap="DEMO barrier text",
                           Next="DEMO next action")
            else:
                row.update(Status="Not Started", Progress=0, VerStatus="Not Submitted")
            out.append(row)
    # edge cases
    out[0]["Owner"] = None                                   # missing owner
    out[3].update(Status="Complete", EvidLink=None, VerStatus="Not Submitted", Actual=as_of - dt.timedelta(days=4))
    out[5].update(Applic="N/A - Approved", Verifier="DEMO Owner 01")  # approved N/A
    out[6].update(Applic="N/A - Requested")
    out[8].update(Required="No")                              # optional
    out[9].update(VerStatus="Verified", EvidLink=None, Status="Complete", Actual=as_of - dt.timedelta(days=9))
    out[10]["ID"] = out[11]["ID"]                             # duplicate id
    out[12].update(Start=as_of + dt.timedelta(days=30), Target=as_of + dt.timedelta(days=10), Status="Not Started",
                   Progress=0, VerStatus="Not Submitted", EvidLink=None, Actual=None, VerDate=None)
    out[14].update(Owner="DEMO Owner 15")                     # inactive owner
    for r in out[15:20]:
        r.update(Status="In Progress", Progress=0.3, Target=as_of - dt.timedelta(days=40), VerStatus="Not Submitted",
                 EvidLink=None, Actual=None, VerDate=None)    # overdue
    return out


def demo_bronze(bdata, tdata, as_of):
    # mark criteria Verified where all linked tasks are verified (coherent demo)
    ver = {}
    for t in tdata:
        if t.get("Applic") == "N/A - Approved" or t.get("Required") == "No":
            continue
        ok = t.get("Status") == "Complete" and t.get("VerStatus") == "Verified" and t.get("EvidLink")
        ver.setdefault(t["BronzeRef"], []).append(bool(ok))
    for i, b in enumerate(bdata):
        b["Owner"] = f"DEMO Owner {random.randint(3, 9):02d}"
        b["Due"] = as_of + dt.timedelta(days=random.randint(-20, 120))
        v = ver.get(b["ReqID"], [])
        if v and all(v):
            b.update(Verif="Verified", Verifier="DEMO Owner 13", VerDate=as_of - dt.timedelta(days=3),
                     EvidLink=f"https://example.sharepoint.com/sites/DunGPS/Shared/{b['Pillar']}/DEMO_{b['ReqID']}.pdf")
        elif v and any(v):
            b.update(Verif="Self-Assessed",
                     EvidLink=f"https://example.sharepoint.com/sites/DunGPS/Shared/{b['Pillar']}/DEMO_{b['ReqID']}.pdf")
    if bdata:
        bdata[-1]["Applic"] = "N/A - Approved by CHG"


def demo_gaps(as_of, bronze_ids):
    cats = ["Standard / procedure missing", "Training / skills", "Resources / time", "Equipment condition",
            "Data / system (CMMS, Redzone)", "Leadership / accountability", "Documentation / evidence",
            "Safety / LOTO", "Food safety / GMP"]
    out = []
    for i in range(1, 29):
        bid, p, req, _ = random.choice(bronze_ids)
        ident = as_of - dt.timedelta(days=random.randint(5, 300))
        due = ident + dt.timedelta(days=random.randint(20, 120))
        sev = random.choices(["Critical", "High", "Medium", "Low"], [1, 3, 4, 2])[0]
        g = dict(ID=f"GAP-{i:03d}", DateId=ident, Pillar=p, BronzeRef=bid, Requirement=req[:60],
                 Area=random.choice(["Nugget Line", "Packaging", "Baking", "Mixing", "Sanitation", "Site-wide"]),
                 Category=random.choices(cats, [5, 4, 4, 3, 2, 2, 3, 1, 1])[0],
                 Current="DEMO current condition", RequiredC="DEMO required condition",
                 Desc=f"DEMO gap {i:03d} – {req[:40]}", Sev=sev, Owner=f"DEMO Owner {random.randint(3, 12):02d}",
                 Due=due, Leader="Yes" if random.random() < 0.15 else "No")
        u = random.random()
        if u < 0.4 and due < as_of + dt.timedelta(days=30):
            close = min(as_of - dt.timedelta(days=1), due + dt.timedelta(days=random.randint(-15, 10)))
            if close < ident:
                close = ident + dt.timedelta(days=5)
            g.update(Status="Closed - Verified", Verif="Verified", VerifiedBy="DEMO Owner 13", Closure=close,
                     Root="DEMO root cause", CA="DEMO corrective action",
                     EvidLink=f"https://example.sharepoint.com/sites/DunGPS/Shared/{p}/DEMO_GAP_{i:03d}.pdf")
        elif u < 0.55:
            g.update(Status="Pending Verification", Verif="Submitted - Pending Review", Root="DEMO root cause",
                     CA="DEMO corrective action")
        else:
            g.update(Status=random.choice(["Open", "In Progress"]), Verif="Not Submitted")
        if g["Leader"] == "Yes":
            g["Decision"] = "DEMO: approve overtime / resource for this gap"
        out.append(g)
    out[1].update(Owner=None)
    out[2].update(Due=None)
    out[4].update(Status="Closed - Verified", Verif="Submitted - Pending Review", Closure=as_of - dt.timedelta(days=2))
    out[7].update(Status="Cancelled")
    return out


def demo_matrix(as_of):
    st = ["Not Assessed", "N/A", "Not Started", "In Progress", "Implemented", "Verified"]
    rows = []
    areas = ["Pre-weigh", "Mixing", "Dough Processing / Forming", "Gemini Equipment", "Nugget Line",
             "Rolled Dough Lines", "Baking", "Cooling / Freezing", "Packaging", "Sanitation"]
    n = 0
    for a in areas:
        for eq, sh_ in ([("DEMO Line 1", "A"), ("DEMO Line 1", "B")] if a in ("Nugget Line", "Rolled Dough Lines")
                        else [("", "All")]):
            n += 1
            r = dict(RowID=f"PA-{n:02d}", Area=a, Equip=eq, Shift=sh_, Owner=f"DEMO Owner {random.randint(9, 12):02d}",
                     LastVer=as_of - dt.timedelta(days=random.randint(5, 60)))
            for j in range(1, 10):
                r[f"E{j}"] = random.choices(st, [1, 1, 2, 3, 2, 2])[0]
            rows.append(r)
    return rows


def demo_projects(as_of):
    return [
        dict(ID="VO-002", Name="DEMO SMED – Nugget Line changeover", Pillar="VO", BronzeRef="B-VO-06",
             Area="Nugget Line", Type="SMED", Charter="Approved", Sponsor="DEMO Owner 02", Lead="DEMO Owner 08",
             Members="DEMO team", Problem="DEMO: changeover too long", KPI="Changeover time", Unit="min",
             Baseline=95, Target=60, Actual=70, Start=as_of - dt.timedelta(days=120),
             TargetDate=as_of - dt.timedelta(days=10), DMAIC="Control", Status="Implemented", EstOpp=120000,
             VerSav=85000, FinApp="Approved", FinBy="DEMO Owner 14", FinDate=as_of - dt.timedelta(days=5),
             Sustain="30-Day Verified"),
        dict(ID="VO-003", Name="DEMO FIT – Packaging film waste", Pillar="VO", BronzeRef="B-VO-07", Area="Packaging",
             Type="Tiger Team / FIT", Charter="Approved", Sponsor="DEMO Owner 02", Lead="DEMO Owner 12",
             Problem="DEMO: film waste above target", KPI="Film waste", Unit="%", Baseline=4.2, Target=2.5,
             Actual=3.6, Start=as_of - dt.timedelta(days=60), TargetDate=as_of + dt.timedelta(days=45), DMAIC="Analyze",
             Status="Active", EstOpp=60000, FinApp="Not Submitted", Sustain="Not Due"),
        dict(ID="VO-004", Name="DEMO Kaizen – Baking oven startup loss", Pillar="VO", Area="Baking", Type="Kaizen Event",
             Charter="Draft", Sponsor="DEMO Owner 01", Lead="DEMO Owner 10", Problem="DEMO", KPI="Startup scrap",
             Unit="lb", Start=as_of - dt.timedelta(days=90), TargetDate=as_of - dt.timedelta(days=5), DMAIC="Measure",
             Status="Active", EstOpp=40000, VerSav=10000, FinApp="Submitted", Sustain="Not Due"),
        dict(ID="VO-005", Name="DEMO 5S – Parts room", Pillar="PM", BronzeRef="B-PM-01", Area="Maintenance / Parts Room",
             Type="5S", Charter="Approved", Sponsor="DEMO Owner 04", Lead="DEMO Owner 04", KPI="5S audit score",
             Unit="%", Baseline=45, Target=80, Actual=82, Start=as_of - dt.timedelta(days=150),
             TargetDate=as_of - dt.timedelta(days=60), ActualDate=as_of - dt.timedelta(days=55), DMAIC="Control",
             Status="Closed", EstOpp=0, FinApp="Not Submitted", Sustain="90-Day Verified"),
    ]


def demo_docs(as_of, bronze_ids):
    types = ["SOP", "OPL", "CIL Sheet", "Centerline Standard", "LOTO Procedure", "Training Record", "Skill Matrix",
             "Meeting Minutes", "Audit / Checklist"]
    out = []
    for i in range(1, 23):
        bid, p, req, _ = random.choice(bronze_ids)
        lr = as_of - dt.timedelta(days=random.randint(10, 420))
        st = random.choices(["Draft", "In Review", "Approved", "Verified"], [2, 2, 5, 3])[0]
        out.append(dict(ID=f"DOC-{i:03d}", Name=f"DEMO {random.choice(types)} – {req[:35]}", Pillar=p, Linked=bid,
                        EvType=random.choice(types), Owner=f"DEMO Owner {random.randint(3, 12):02d}",
                        Rev=f"{random.randint(0, 4)}", LastReview=lr, Freq=12, Status=st,
                        Link=None if i % 7 == 0 else f"https://example.sharepoint.com/sites/DunGPS/Shared/{p}/DEMO_DOC_{i:03d}.docx"))
    return out


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "."
    build(False, f"{out}/CHG_DunGPS_Deployment_Master.xlsx")
    build(True, f"{out}/CHG_DunGPS_Deployment_Master_DEMO.xlsx")
    print("built")
