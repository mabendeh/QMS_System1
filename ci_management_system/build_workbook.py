#!/usr/bin/env python3
"""
C.H. Guenther & Son - Duncanville
Supervisor Productivity & Continuous Improvement Management System

Workbook generator (openpyxl). Produces:
  output/CHG_Duncanville_CI_Management_System.xlsx      (production - real assignments only)
  output/CHG_Duncanville_CI_System_DEMO_TRAINING.xlsx   (fictitious demo data for training/testing)

Usage:  python3 build_workbook.py [--qa qa_results.json]

All dashboard metrics are formulas. Nothing in the production file is fabricated:
unknown owners, dates, baselines and savings are left blank and flagged.
"""
import argparse
import datetime as dt
import json
import os
import sys

import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.table import Table, TableStyleInfo

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "output")
VERSION = "1.0"
BUILD_DATE = dt.date(2026, 11, 19)

# ----------------------------------------------------------------------------- sheet names
S_DASH = "01 Executive Dashboard"
S_REG = "02 Project Register"
S_CH = "03 Project Charter"
S_CHR = "03A Charter Records"
S_ACT = "04 Action Tracker"
S_FIN = "05 Financial Savings"
S_WOR = "06 Weekly WOR Report"
S_WORU = "06A WOR Weekly Updates"
S_DEV = "07 Supervisor Development"
S_DEVR = "07A Rubric and Summary"
S_TOOLS = "08 CI Tools"
S_RISK = "09 Risks and Escalations"
S_EVD = "10 Evidence Register"
S_MD = "11 Master Data"
S_GUIDE = "12 User Guide"
S_DD = "13 Data Dictionary"
S_QA = "14 QA Report"


def q(sheet):
    return "'" + sheet + "'"


# ----------------------------------------------------------------------------- palette / styles
NAVY = "1F3864"
NAVY2 = "2F5496"
STEEL = "44546A"
GRAY_D = "595959"
GRAY_M = "BFBFBF"
GRAY_L = "D9D9D9"
GRAY_XL = "F2F2F2"
CALC_FILL = "F5F6F8"
INPUT_FILL = "FFFBE6"
INPUT_FONT = "1F4E9F"
WHITE = "FFFFFF"
G_FILL, G_FONT = "E2EFDA", "375623"
A_FILL, A_FONT = "FFF2CC", "7F6000"
R_FILL, R_FONT = "F8D7DA", "9C0006"
P_FILL, P_FONT = "E7E6E6", "404040"

FONT = "Arial"


def F(size=10, bold=False, color="000000", italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def fill(c):
    return PatternFill("solid", start_color=c, end_color=c)


THIN = Side(style="thin", color=GRAY_L)
MED = Side(style="medium", color=NAVY)
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BOTTOM = Border(bottom=THIN)
UNLOCK = Protection(locked=False)
LOCK = Protection(locked=True)
WRAP_TOP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_C = Alignment(horizontal="left", vertical="center", wrap_text=True)

FMT_DATE = "mm/dd/yyyy"
FMT_USD = '$#,##0;($#,##0);"-"'
FMT_USD2 = '$#,##0.00;($#,##0.00);"-"'
FMT_PCT = "0%"
FMT_PCT1 = "0.0%"
FMT_NUM = "#,##0.##"
FMT_INT = "0"

# Excel serials for date validation
D2020 = 43831  # 2020-01-01
D2040 = 51501  # 2040-12-31


def style_input(c, fmt=None, wrap=False):
    c.fill = fill(INPUT_FILL)
    c.font = F(10, color=INPUT_FONT)
    c.border = BORDER
    c.protection = UNLOCK
    c.alignment = Alignment(wrap_text=wrap, vertical="top")
    if fmt:
        c.number_format = fmt


def style_calc(c, fmt=None, wrap=False, bold=False):
    c.fill = fill(CALC_FILL)
    c.font = F(10, bold=bold)
    c.border = BORDER
    c.protection = LOCK
    c.alignment = Alignment(wrap_text=wrap, vertical="top")
    if fmt:
        c.number_format = fmt


def merge_write(ws, rng, value, font=None, fll=None, align=None, fmt=None, border=None):
    ws.merge_cells(rng)
    c = ws[rng.split(":")[0]]
    c.value = value
    if font:
        c.font = font
    if fll:
        for row in ws[rng]:
            for cc in row:
                cc.fill = fll
    if border:
        for row in ws[rng]:
            for cc in row:
                cc.border = border
    if align:
        c.alignment = align
    if fmt:
        c.number_format = fmt
    return c


def banner(ws, title, subtitle, last_col, demo, note=None):
    """Rows 1-3 header band. Text only - no logo is fabricated."""
    lc = get_column_letter(last_col)
    company = "C.H. GUENTHER & SON  |  DUNCANVILLE PLANT"
    if demo:
        company = "DEMONSTRATION / TRAINING COPY - FICTITIOUS DATA - NOT FOR WOR OR FINANCE REPORTING"
    merge_write(ws, f"A1:{lc}1", company, F(13, True, WHITE), fill(R_FONT if demo else NAVY), LEFT_C)
    merge_write(ws, f"A2:{lc}2", title, F(16, True, NAVY), fill(WHITE), LEFT_C)
    merge_write(ws, f"A3:{lc}3", subtitle, F(9, False, GRAY_D, True), fill(WHITE), LEFT_C)
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 26
    ws.row_dimensions[3].height = 30 if note else 18
    for col in range(1, last_col + 1):
        ws.cell(3, col).border = Border(bottom=MED)


def section(ws, rng, text, color=NAVY2):
    merge_write(ws, rng, text, F(11, True, WHITE), fill(color), LEFT_C)


def add_name(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def dv_list(ws, name, rng, prompt=None):
    dv = DataValidation(type="list", formula1=f"={name}", allow_blank=True, showDropDown=False)
    dv.error = "Select a value from the list (lists are maintained in '11 Master Data')."
    dv.errorTitle = "Invalid entry"
    dv.showErrorMessage = True
    if prompt:
        dv.prompt = prompt
        dv.showInputMessage = True
    ws.add_data_validation(dv)
    dv.add(rng)
    return dv


def dv_date(ws, rng):
    dv = DataValidation(type="date", operator="between", formula1=str(D2020), formula2=str(D2040), allow_blank=True)
    dv.error = "Enter a valid date (mm/dd/yyyy) between 2020 and 2040."
    dv.errorTitle = "Invalid date"
    dv.showErrorMessage = True
    ws.add_data_validation(dv)
    dv.add(rng)


def dv_decimal(ws, rng, lo=None, hi=None, msg="Enter a number."):
    if lo is not None and hi is not None:
        dv = DataValidation(type="decimal", operator="between", formula1=str(lo), formula2=str(hi), allow_blank=True)
    else:
        dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="-1E+15", allow_blank=True)
    dv.error = msg
    dv.errorTitle = "Invalid number"
    dv.showErrorMessage = True
    ws.add_data_validation(dv)
    dv.add(rng)


def dv_whole(ws, rng, lo, hi, msg):
    dv = DataValidation(type="whole", operator="between", formula1=str(lo), formula2=str(hi), allow_blank=True)
    dv.error = msg
    dv.errorTitle = "Invalid score"
    dv.showErrorMessage = True
    ws.add_data_validation(dv)
    dv.add(rng)


def protect(ws):
    p = ws.protection
    p.sheet = True
    p.formatColumns = False
    p.formatRows = False
    p.formatCells = True
    p.autoFilter = False
    p.sort = False
    p.selectLockedCells = False
    p.selectUnlockedCells = False
    p.insertRows = True
    p.deleteRows = True


class RowRef(dict):
    """format_map helper: {key} -> A6 style cell ref of same row, {key_c} -> column letter, {r} -> row."""

    def __init__(self, cols, r):
        super().__init__()
        self.cols = cols
        self.r = r

    def __missing__(self, k):
        if k == "r":
            return str(self.r)
        if k.endswith("_c"):
            return self.cols[k[:-2]]
        return f"{self.cols[k]}{self.r}"


HDR_ROW = 5
FIRST = 6


def build_table(wb, ws, tname, prefix, cols, cap, data=None, freeze="C6", hdr_height=48):
    """cols: list of dicts with keys:
         key, h (header), w (width), kind ('in'|'calc'), f (formula template), fmt,
         dv (named list), t ('date'|'pct'|'num'|'score'|'text'), note (header comment), wrap
    Creates an Excel table, column-level defined names <prefix>_<Key>, validation and styling."""
    letters = {}
    for j, c in enumerate(cols):
        letters[c["key"]] = get_column_letter(j + 1)
    last = get_column_letter(len(cols))
    end = HDR_ROW + cap
    for j, c in enumerate(cols):
        L = letters[c["key"]]
        h = ws.cell(HDR_ROW, j + 1, c["h"])
        h.font = F(9, True, WHITE)
        h.fill = fill(NAVY if c["kind"] == "in" else GRAY_D)
        h.alignment = CENTER
        h.border = BORDER
        h.protection = LOCK
        note = c.get("note")
        if c["kind"] == "calc":
            note = ("AUTO (formula, locked). " + (note or "")).strip()
        else:
            note = ("INPUT. " + (note or "")).strip()
        h.comment = Comment(note, "CI System")
        h.comment.width = 320
        h.comment.height = 140
        ws.column_dimensions[L].width = c.get("w", 14)
    ws.row_dimensions[HDR_ROW].height = hdr_height
    data = data or []
    for i in range(cap):
        r = FIRST + i
        rr = RowRef(letters, r)
        rowdata = data[i] if i < len(data) else {}
        for j, c in enumerate(cols):
            cell = ws.cell(r, j + 1)
            fmt = c.get("fmt")
            t = c.get("t")
            if fmt is None:
                fmt = {"date": FMT_DATE, "pct": FMT_PCT, "num": FMT_NUM, "score": FMT_INT, "usd": FMT_USD}.get(t)
            if c["kind"] == "calc":
                cell.value = c["f"].format_map(rr)
                style_calc(cell, fmt, c.get("wrap", False))
            else:
                v = rowdata.get(c["key"])
                if v is not None:
                    cell.value = v
                style_input(cell, fmt, c.get("wrap", False))
    # validation
    for c in cols:
        L = letters[c["key"]]
        rng = f"{L}{FIRST}:{L}{end}"
        if c["kind"] != "in":
            continue
        if c.get("dv"):
            dv_list(ws, c["dv"], rng)
        elif c.get("t") == "date":
            dv_date(ws, rng)
        elif c.get("t") == "pct":
            dv_decimal(ws, rng, 0, 1, "Enter a percentage between 0% and 100%.")
        elif c.get("t") in ("num", "usd"):
            dv_decimal(ws, rng, msg="Enter a number (no text). Leave blank if unknown - do not enter 0 for missing data.")
        elif c.get("t") == "score":
            dv_whole(ws, rng, 1, 4, "Score must be a whole number 1-4 per the rubric. Leave blank if not assessed.")
    # names
    for c in cols:
        L = letters[c["key"]]
        add_name(wb, f"{prefix}_{c['key']}", f"{q(ws.title)}!${L}${FIRST}:${L}${end}")
    tab = Table(displayName=tname, ref=f"A{HDR_ROW}:{last}{end}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False, showColumnStripes=False)
    ws.add_table(tab)
    ws.freeze_panes = freeze
    ws.sheet_view.zoomScale = 90
    # print setup: landscape, fit to width, header repeated, print area grows with the records entered
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_TABLOID if len(cols) > 30 else ws.PAPERSIZE_LETTER
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{HDR_ROW}:{HDR_ROW}"
    ws.page_margins.left = ws.page_margins.right = 0.25
    first_col = letters[cols[0]["key"]]
    ws.defined_names["_xlnm.Print_Area"] = DefinedName(
        "_xlnm.Print_Area",
        attr_text=f"OFFSET({q(ws.title)}!$A$1,0,0,{HDR_ROW}+MAX(1,COUNTIF({q(ws.title)}!${first_col}${FIRST}:${first_col}${end},\"?*\")),{len(cols)})")
    ws.oddFooter.left.text = "C.H. Guenther & Son - Duncanville | " + ws.title
    ws.oddFooter.right.text = "Printed &D | Page &P of &N"
    return letters, end


def legend_row(ws, row, last_col, extra=""):
    lc = get_column_letter(last_col)
    txt = ("LEGEND:  Navy header + yellow cell = INPUT (type or pick from dropdown).   Gray header + gray cell = AUTO formula "
           "(locked - do not overwrite).   Hover a header for its definition.   Do not leave blank rows between records.  " + extra)
    merge_write(ws, f"A{row}:{lc}{row}", txt, F(8, False, GRAY_D, True), fill(GRAY_XL), LEFT_C)
    ws.row_dimensions[row].height = 28


# ============================================================================= MASTER DATA
PEOPLE_CAP = 200
LIST_CAP = 60
MD_HDR = 21
MD_FIRST = 22

LISTS = [
    # key, header, column, values, system(locked existing values)
    ("Dept", "Departments", "I", ["Production", "Maintenance", "Warehouse", "Quality", "Sanitation", "Finance",
                                   "Continuous Improvement", "Cross-Functional"], False),
    ("Area", "Areas / Lines", "J", ["Plant-wide"], False),
    ("Equipment", "Equipment", "K", ["Not equipment-specific"], False),
    ("Shift", "Shifts", "L", ["A", "B", "C", "Cross-Shift"], False),
    ("Category", "Project Categories", "M", ["SAP / Systems Accuracy", "Changeover / SMED", "Waste Reduction",
                                             "Maintenance Reliability", "Warehouse / Logistics", "Quality",
                                             "Safety", "5S / Workplace Organization", "Labor Productivity", "Other"], False),
    ("Stage", "Project Stages", "N", ["1-Define", "2-Measure", "3-Analyze", "4-Improve", "5-Control", "6-Closed"], True),
    ("ProjStatus", "Project Statuses", "O", ["Not Started", "Active", "On Hold", "Blocked", "Complete", "Cancelled",
                                             "Pending Assignment"], True),
    ("ActStatus", "Action Statuses", "P", ["Not Started", "In Progress", "Blocked", "Complete", "Cancelled"], True),
    ("ActCategory", "Action Categories", "Q", ["Data Collection / Measurement", "Root Cause Analysis",
                                               "Countermeasure / Implementation", "Standard Work / Documentation",
                                               "Training", "Verification / Sustainment", "Communication",
                                               "Finance Validation", "Other"], False),
    ("Priority", "Priorities", "R", ["High", "Medium", "Low"], False),
    ("FinStatus", "Finance Approval Statuses", "S", ["Not Submitted", "Under Review", "Approved", "Rejected"], True),
    ("BenefitType", "Benefit Types", "T", ["Hard Savings", "Cost Avoidance", "Productivity Capacity"], True),
    ("CalcMethod", "Savings Calculation Methods", "U", ["Waste - Material Loss", "Changeover - Time Reduction",
                                                        "SAP - Error / Rework / Labor", "Other - Data Supported"], True),
    ("EvidenceType", "Evidence Types", "V", ["Photo - Before", "Photo - After", "Data Extract / Report", "SAP Report",
                                             "Standard Work / SOP", "Training Record", "Audit / Checklist",
                                             "Meeting Minutes", "Finance Validation", "Other"], False),
    ("Verif", "Verification Statuses", "W", ["Pending Review", "Verified", "Needs More Info", "Rejected"], True),
    ("RiskType", "Risk / Blocker Types", "X", ["Blocker", "Risk", "Issue", "Resource Need"], False),
    ("RiskStatus", "Risk Statuses", "Y", ["Open", "Escalated", "Resolved", "Closed"], True),
    ("Impact", "Impact Levels", "Z", ["High", "Medium", "Low"], True),
    ("EscLevel", "Escalation Levels", "AA", ["Supervisor", "CI Manager", "Department Manager", "Plant Leadership",
                                             "Finance"], True),
    ("CharterStatus", "Charter Statuses", "AB", ["Draft", "Submitted", "Approved", "Revision Required"], True),
    ("YesNo", "Yes / No", "AC", ["Yes", "No"], True),
    ("Role", "Roles", "AD", ["Production Supervisor", "Maintenance Supervisor", "Warehouse Supervisor",
                             "Supervisor", "CI Manager", "Finance Reviewer", "Sponsor / Manager", "Team Member"], False),
]


def build_master(wb, ws, demo, people):
    banner(ws, "11  MASTER DATA AND SETTINGS",
           "Editable dropdown lists, system settings and the owner-assignment lookup. Lists grow automatically - add new "
           "values directly below the last entry (no blank rows). SYSTEM lists: existing values are locked because formulas "
           "depend on them; you may add new values below.", 34, demo, note=True)
    # ---------------- settings
    section(ws, "B5:D5", "SYSTEM SETTINGS")
    settings = [
        ("Report As-Of Date override", None, "INPUT. Leave blank to use today's date. Enter a date to freeze the report "
                                             "(e.g., the WOR meeting date).", "date"),
        ("Effective As-Of Date (used by all formulas)", "=IF(ISNUMBER(C6),C6,TODAY())",
         "AUTO. Drives overdue flags, days remaining, RAG and trend windows.", "calc_date"),
        ("Project ID prefix", "DUN-CI-", "INPUT. Project IDs are generated as prefix + 3-digit row number.", "text"),
        ("Operating weeks per year (annualization)", 52,
         "ASSUMPTION - default 52 pending Finance confirmation. Annualized = benefit per period x (weeks per year / "
         "period length in weeks).", "num"),
        ("Overdue actions that trigger RED", 3, "INPUT. A project turns RED at this many overdue actions.", "num"),
        ("Due-soon window (days)", 7, "INPUT. Open actions due within this many days are flagged DUE SOON.", "num"),
        ("Production volume comparability tolerance", 0.2,
         "INPUT. Waste savings lines whose comparison volume differs from baseline by more than this are flagged for "
         "Finance.", "pct"),
        ("Approved evidence folder URL (SharePoint / Teams)", None,
         "INPUT (CI Manager). Paste the approved folder link. Not configured until entered.", "text"),
        ("Evidence folder link", '=IF(C13="","Not configured - CI Manager to enter the approved folder URL above",'
                                 'HYPERLINK(C13,"Open approved evidence folder"))', "AUTO hyperlink.", "calc"),
        ("Site", "Duncanville", "", "text"),
        ("Company", "C.H. Guenther & Son", "", "text"),
        ("Workbook version", f"{VERSION} ({BUILD_DATE.isoformat()})", "Update when the structure changes.", "calc_text"),
    ]
    names = ["AsOfOverride", "AsOfDate", "IDPrefix", "OpWeeks", "RedOverdueThr", "DueSoonDays", "VolTol",
             "EvidenceRoot", None, "SiteName", "CompanyName", None]
    for i, (lab, val, note, kind) in enumerate(settings):
        r = 6 + i
        a = ws.cell(r, 2, lab)
        a.font = F(9, True, NAVY)
        a.alignment = LEFT_C
        a.border = BORDER
        c = ws.cell(r, 3, val)
        if kind.startswith("calc"):
            style_calc(c, FMT_DATE if kind == "calc_date" else None, bold=True)
        else:
            style_input(c, {"date": FMT_DATE, "pct": FMT_PCT, "num": "0"}.get(kind))
        n = ws.cell(r, 4, note)
        n.font = F(8, False, GRAY_D, True)
        n.alignment = LEFT_C
        n.border = BORDER
        ws.row_dimensions[r].height = 30
        if names[i]:
            add_name(wb, names[i], f"{q(ws.title)}!$C${r}")
    if demo:
        ws["C6"] = dt.date(2026, 11, 19)
    dv_date(ws, "C6")
    ws["C14"].font = Font(name=FONT, size=10, color="0563C1", underline="single", bold=True)
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 26
    ws.column_dimensions["D"].width = 46
    ws.column_dimensions["E"].width = 2
    ws.column_dimensions["H"].width = 2

    # ---------------- people table (B:G)
    section(ws, f"B{MD_HDR-1}:G{MD_HDR-1}", "PEOPLE / OWNERS (add new supervisors here - no blank rows)")
    hdrs = ["Name", "Role", "Department", "Shift", "Email (optional)", "Active"]
    for j, h in enumerate(hdrs):
        c = ws.cell(MD_HDR, 2 + j, h)
        c.font = F(9, True, WHITE)
        c.fill = fill(NAVY)
        c.alignment = CENTER
        c.border = BORDER
    for i in range(PEOPLE_CAP):
        r = MD_FIRST + i
        p = people[i] if i < len(people) else None
        for j in range(6):
            c = ws.cell(r, 2 + j, p[j] if p and j < len(p) else None)
            style_input(c)
    pe = MD_FIRST + PEOPLE_CAP - 1
    tab = Table(displayName="tblPeople", ref=f"B{MD_HDR}:G{pe}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
    ws.add_table(tab)
    add_name(wb, "md_PeopleName", f"{q(ws.title)}!$B${MD_FIRST}:$B${pe}")
    add_name(wb, "md_PeopleRole", f"{q(ws.title)}!$C${MD_FIRST}:$C${pe}")
    add_name(wb, "md_PeopleDept", f"{q(ws.title)}!$D${MD_FIRST}:$D${pe}")
    add_name(wb, "md_PeopleShift", f"{q(ws.title)}!$E${MD_FIRST}:$E${pe}")
    add_name(wb, "lst_People",
             f"OFFSET({q(ws.title)}!$B${MD_FIRST},0,0,MAX(1,COUNTA({q(ws.title)}!$B${MD_FIRST}:$B${pe})),1)")
    dv_list(ws, "lst_Role", f"C{MD_FIRST}:C{pe}")
    dv_list(ws, "lst_Dept", f"D{MD_FIRST}:D{pe}")
    dv_list(ws, "lst_Shift", f"E{MD_FIRST}:E{pe}")
    dv_list(ws, "lst_YesNo", f"G{MD_FIRST}:G{pe}")
    for col, w in zip("BCDEFG", [24, 22, 20, 11, 24, 8]):
        ws.column_dimensions[col].width = max(ws.column_dimensions[col].width or 0, w)

    # ---------------- dropdown lists
    section(ws, f"I{MD_HDR-1}:AD{MD_HDR-1}",
            "DROPDOWN LISTS  (navy = editable list; steel = SYSTEM list: existing values locked, new values may be added)")
    le = MD_FIRST + LIST_CAP - 1
    for key, header, col, values, system in LISTS:
        ci = openpyxl.utils.column_index_from_string(col)
        h = ws.cell(MD_HDR, ci, header)
        h.font = F(9, True, WHITE)
        h.fill = fill(STEEL if system else NAVY)
        h.alignment = CENTER
        h.border = BORDER
        h.comment = Comment(("SYSTEM list - formulas reference these exact values; do not rename or delete them. "
                             "New values may be added below." if system else
                             "Editable list - rename, add or remove values. Keep the list contiguous (no blank rows)."),
                            "CI System")
        for i in range(LIST_CAP):
            r = MD_FIRST + i
            c = ws.cell(r, ci, values[i] if i < len(values) else None)
            style_input(c)
            if system and i < len(values):
                c.protection = LOCK
                c.fill = fill(GRAY_XL)
                c.font = F(10, color="000000")
        ws.column_dimensions[col].width = 24
        add_name(wb, f"lst_{key}", f"OFFSET({q(ws.title)}!${col}${MD_FIRST},0,0,MAX(1,COUNTA({q(ws.title)}!${col}${MD_FIRST}:${col}${le})),1)")
        add_name(wb, f"md_{key}", f"{q(ws.title)}!${col}${MD_FIRST}:${col}${le}")
    ws.row_dimensions[MD_HDR].height = 32

    # ---------------- filter lists (auto) AF:AH
    section(ws, f"AF{MD_HDR-1}:AH{MD_HDR-1}", "REPORT FILTER LISTS (AUTO)", GRAY_D)
    for ci, (hdr, src, nm) in zip([32, 33, 34], [("Dept filter", "I", "lst_FDept"), ("Shift filter", "L", "lst_FShift"),
                                                ("Owner filter", "B", "lst_FOwner")]):
        L = get_column_letter(ci)
        h = ws.cell(MD_HDR, ci, hdr)
        h.font = F(9, True, WHITE)
        h.fill = fill(GRAY_D)
        h.alignment = CENTER
        ws.cell(MD_FIRST, ci, "(All)")
        style_calc(ws.cell(MD_FIRST, ci))
        cap = PEOPLE_CAP if src == "B" else LIST_CAP
        for i in range(cap):
            r = MD_FIRST + 1 + i
            c = ws.cell(r, ci, f'=IF({src}{MD_FIRST + i}="","",{src}{MD_FIRST + i})')
            style_calc(c)
        srcend = MD_FIRST + cap - 1
        add_name(wb, nm, f"OFFSET({q(ws.title)}!${L}${MD_FIRST},0,0,1+COUNTIF({q(ws.title)}!${src}${MD_FIRST}:${src}${srcend},\"?*\"),1)")
        ws.column_dimensions[L].width = 20
    ws.column_dimensions["AE"].width = 2

    # ---------------- owner assignment interface F5:J16 -> use columns F..G and I..K? keep F:G and J:L region rows 5-17
    section(ws, "F5:G5", "OWNER ASSIGNMENT LOOKUP")
    ws.column_dimensions["F"].width = max(ws.column_dimensions["F"].width or 0, 26)
    ws.column_dimensions["G"].width = 26
    ws["F6"] = "Select person"
    ws["G6"] = people[0][0] if people else None
    style_input(ws["G6"])
    dv_list(ws, "lst_People", "G6")
    rows = [
        ("Role", '=IFERROR(IF(INDEX(md_PeopleRole,MATCH($G$6,md_PeopleName,0))="","Not set",INDEX(md_PeopleRole,MATCH($G$6,md_PeopleName,0))),"Not in people list")'),
        ("Department", '=IFERROR(IF(INDEX(md_PeopleDept,MATCH($G$6,md_PeopleName,0))="","Not set",INDEX(md_PeopleDept,MATCH($G$6,md_PeopleName,0))),"Not in people list")'),
        ("Shift", '=IFERROR(IF(INDEX(md_PeopleShift,MATCH($G$6,md_PeopleName,0))="","Not set",INDEX(md_PeopleShift,MATCH($G$6,md_PeopleName,0))),"Not in people list")'),
        ("Projects as PRIMARY leader", '=IF($G$6="","",COUNTIFS(reg_Leader,$G$6,reg_ID,"?*"))'),
        ("Projects as co-owner", '=IF($G$6="","",COUNTIFS(reg_CoOwner,$G$6,reg_ID,"?*"))'),
        ("Projects as sponsor / coach / Finance", '=IF($G$6="","",COUNTIFS(reg_Sponsor,$G$6,reg_ID,"?*")+COUNTIFS(reg_Coach,$G$6,reg_ID,"?*")+COUNTIFS(reg_FinRev,$G$6,reg_ID,"?*"))'),
        ("Open actions owned", '=IF($G$6="","",COUNTIFS(act_Owner,$G$6,act_Credit,0,act_Status,"<>Cancelled"))'),
        ("Overdue actions owned", '=IF($G$6="","",COUNTIFS(act_Owner,$G$6,act_Flag,"OVERDUE"))'),
        ("Open risks / escalations owned", '=IF($G$6="","",SUMIFS(risk_IsOpen,risk_Owner,$G$6))'),
        ("Development reviews on file", '=IF($G$6="","",COUNTIF(dev_Sup,$G$6))'),
    ]
    for i, (lab, f) in enumerate(rows):
        r = 7 + i
        ws.cell(r, 6, lab).font = F(9, True, NAVY)
        ws.cell(r, 6).border = BORDER
        c = ws.cell(r, 7, f)
        style_calc(c, bold=True)
    ws["F6"].font = F(9, True, NAVY)
    ws["F17"] = ("Assign owners in '02 Project Register' (Project Leader = single accountable owner; Co-Owner and "
                 "Additional Team Members support).")
    ws["F17"].font = F(8, False, GRAY_D, True)
    ws["F17"].alignment = WRAP_TOP
    ws.merge_cells("F17:G18")

    section(ws, "J5:L5", "PROJECT OWNERSHIP LOOKUP")
    ws["J6"] = "Select project ID"
    ws["J6"].font = F(9, True, NAVY)
    ws["K6"] = "DEMO-001" if demo else "DUN-CI-001"
    style_input(ws["K6"])
    ws.merge_cells("K6:L6")
    dv_list(ws, "lst_ProjectID", "K6")
    prow = [("Project title", "reg_Title"), ("Primary leader (accountable)", "reg_Leader"), ("Co-owner", "reg_CoOwner"),
            ("Additional team members", "reg_Team"), ("Sponsor", "reg_Sponsor"), ("CI coach", "reg_Coach"),
            ("Finance reviewer", "reg_FinRev"), ("Status", "reg_Status"), ("RAG", "reg_RAG"), ("Owner check", "reg_OwnerChk")]
    for i, (lab, nm) in enumerate(prow):
        r = 7 + i
        ws.cell(r, 10, lab).font = F(9, True, NAVY)
        c = ws.cell(r, 11, f'=IFERROR(IF(INDEX({nm},MATCH($K$6,reg_ID,0))&""="","Not assigned",INDEX({nm},MATCH($K$6,reg_ID,0))),"Unknown project ID")')
        style_calc(c, bold=True)
        ws.merge_cells(f"K{r}:L{r}")
    ws.freeze_panes = "A5"
    ws.sheet_view.zoomScale = 85


# ============================================================================= PROJECT REGISTER
def register_cols():
    asof = "AsOfDate"
    return [
        dict(key="ID", h="Project ID", w=13, kind="calc", f='=IF(TRIM({Title})="","",IDPrefix&TEXT(ROW()-5,"000"))',
             note="Generated from the ID prefix (Master Data) and row position once a title is entered. Never re-use or re-sort rows."),
        dict(key="Title", h="Project Title", w=34, kind="in", wrap=True, note="Short, specific title. Entering a title creates the project ID."),
        dict(key="Desc", h="Description", w=44, kind="in", wrap=True),
        dict(key="Category", h="Project Category", w=20, kind="in", dv="lst_Category"),
        dict(key="Dept", h="Department", w=16, kind="in", dv="lst_Dept"),
        dict(key="Area", h="Area / Line", w=14, kind="in", dv="lst_Area"),
        dict(key="Equipment", h="Equipment", w=16, kind="in", dv="lst_Equipment"),
        dict(key="Shift", h="Shift", w=10, kind="in", dv="lst_Shift"),
        dict(key="Leader", h="Project Leader (Primary Owner)", w=20, kind="in", dv="lst_People",
             note="ONE accountable owner. Must be listed in Master Data > People."),
        dict(key="CoOwner", h="Co-Owner", w=18, kind="in", dv="lst_People"),
        dict(key="Team", h="Additional Team Members", w=26, kind="in", wrap=True,
             note="Free text, separate names with semicolons. Supports multiple members without diluting primary ownership."),
        dict(key="Sponsor", h="Sponsor", w=16, kind="in", dv="lst_People"),
        dict(key="Coach", h="CI Coach", w=16, kind="in", dv="lst_People"),
        dict(key="FinRev", h="Finance Reviewer", w=16, kind="in", dv="lst_People"),
        dict(key="Priority", h="Priority", w=10, kind="in", dv="lst_Priority"),
        dict(key="Stage", h="Stage", w=12, kind="in", dv="lst_Stage"),
        dict(key="Status", h="Project Status", w=14, kind="in", dv="lst_ProjStatus"),
        dict(key="Start", h="Start Date", w=12, kind="in", t="date"),
        dict(key="TargetDate", h="Target Date", w=12, kind="in", t="date"),
        dict(key="CompDate", h="Completion Date", w=12, kind="in", t="date"),
        dict(key="KPI", h="Primary KPI / Metric", w=24, kind="in", wrap=True,
             note="The outcome measure, e.g. 'SAP confirmation errors per week'. Actuals are entered weekly in 06A."),
        dict(key="Unit", h="KPI Unit", w=12, kind="in"),
        dict(key="Baseline", h="Baseline", w=11, kind="in", t="num",
             note="Measured baseline. Leave BLANK until measured - blank shows N/A (never assumed zero)."),
        dict(key="Target", h="Target", w=11, kind="in", t="num"),
        dict(key="Notes", h="Register Notes", w=34, kind="in", wrap=True),
        # ---- AUTO
        dict(key="KPIWeek", h="Latest KPI Week", w=12, kind="calc", t="date",
             f='=IF({ID}="","",IF(COUNTIFS(wor_Proj,{ID},wor_KPI,"<>")=0,"",_xlfn.MAXIFS(wor_Week,wor_Proj,{ID},wor_KPI,"<>")))',
             note="Week-ending date of the most recent KPI actual recorded in 06A WOR Weekly Updates."),
        dict(key="Actual", h="Latest KPI Actual", w=11, kind="calc", fmt=FMT_NUM,
             f='=IF({ID}="","",IF({KPIWeek}="","N/A",IFERROR(INDEX(wor_KPI,MATCH({ID}&"|"&{KPIWeek},wor_Key,0)),"N/A")))'),
        dict(key="Attain", h="Outcome Attainment %", w=12, kind="calc", fmt=FMT_PCT,
             f='=IF({ID}="","",IF(OR(NOT(ISNUMBER({Baseline})),NOT(ISNUMBER({Target})),NOT(ISNUMBER({Actual}))),"N/A",IF({Baseline}={Target},"N/A",({Baseline}-{Actual})/({Baseline}-{Target}))))',
             note="(Baseline - Actual) / (Baseline - Target). 100% = target met; works for increase or decrease goals; <0% = worse than baseline. N/A if any value is missing."),
        dict(key="NAct", h="Total Actions", w=9, kind="calc", fmt=FMT_INT,
             f='=IF({ID}="","",COUNTIFS(act_Proj,{ID},act_Status,"<>Cancelled"))', note="Actions in 04 for this project, excluding Cancelled."),
        dict(key="NDone", h="Actions Complete", w=9, kind="calc", fmt=FMT_INT,
             f='=IF({ID}="","",COUNTIFS(act_Proj,{ID},act_Credit,1))', note="Credited = Status Complete AND a completed date entered."),
        dict(key="ActPct", h="Action Completion %", w=11, kind="calc", fmt=FMT_PCT,
             f='=IF({ID}="","",IF({NAct}=0,"N/A",{NDone}/{NAct}))', note="Credited complete / total non-cancelled actions. Count-based - subjective % complete is NOT averaged."),
        dict(key="NMS", h="Milestones", w=9, kind="calc", fmt=FMT_INT,
             f='=IF({ID}="","",COUNTIFS(act_Proj,{ID},act_MS,"Yes",act_Status,"<>Cancelled"))'),
        dict(key="NMSDone", h="Milestones Complete", w=10, kind="calc", fmt=FMT_INT,
             f='=IF({ID}="","",COUNTIFS(act_Proj,{ID},act_MS,"Yes",act_Credit,1))'),
        dict(key="MSPct", h="Milestone Completion %", w=11, kind="calc", fmt=FMT_PCT,
             f='=IF({ID}="","",IF({NMS}=0,"N/A",{NMSDone}/{NMS}))'),
        dict(key="Overdue", h="Overdue Actions", w=9, kind="calc", fmt=FMT_INT,
             f='=IF({ID}="","",COUNTIFS(act_Proj,{ID},act_Flag,"OVERDUE"))'),
        dict(key="NextMSDue", h="Next Milestone Due", w=12, kind="calc", t="date",
             f='=IF({ID}="","",IF(COUNTIFS(act_Proj,{ID},act_OpenMS,1)=0,"",_xlfn.MINIFS(act_Due,act_Proj,{ID},act_OpenMS,1)))'),
        dict(key="NextMS", h="Next Milestone", w=30, kind="calc", wrap=True,
             f='=IF({ID}="","",IF({NextMSDue}="","No open milestone scheduled",IFERROR(INDEX(act_Desc,MATCH({ID}&"|"&{NextMSDue},act_MSKey,0)),"")))',
             note="Earliest-due open action flagged Milestone = Yes in 04 Action Tracker."),
        dict(key="Risks", h="Open Risks / Escalations", w=10, kind="calc", fmt=FMT_INT,
             f='=IF({ID}="","",SUMIFS(risk_IsOpen,risk_Proj,{ID}))'),
        dict(key="ExpHard", h="Expected Savings - Hard ($/yr, forecast)", w=14, kind="calc", fmt=FMT_USD,
             f='=IF({ID}="","",IF(COUNTIFS(fin_Proj,{ID})=0,"Not quantified",SUMIFS(fin_FcstNum,fin_Proj,{ID},fin_Type,"Hard Savings",fin_Status,"<>Rejected")))',
             note="Attributed net annual forecast of Hard Savings lines not Rejected. FORECAST - not verified unless Finance Approved."),
        dict(key="ExpAvoid", h="Expected Cost Avoidance ($/yr, forecast)", w=14, kind="calc", fmt=FMT_USD,
             f='=IF({ID}="","",IF(COUNTIFS(fin_Proj,{ID})=0,"Not quantified",SUMIFS(fin_FcstNum,fin_Proj,{ID},fin_Type,"Cost Avoidance",fin_Status,"<>Rejected")))'),
        dict(key="CapHrs", h="Capacity Hours / yr (non-cash)", w=12, kind="calc", fmt="#,##0.0",
             f='=IF({ID}="","",IF(COUNTIFS(fin_Proj,{ID})=0,"Not quantified",SUMIFS(fin_CapHrsYr,fin_Proj,{ID},fin_Status,"<>Rejected")))'),
        dict(key="VerHard", h="Verified Savings - Hard ($, Finance-approved)", w=14, kind="calc", fmt=FMT_USD,
             f='=IF({ID}="","",SUMIFS(fin_Verified,fin_Proj,{ID},fin_Type,"Hard Savings"))',
             note="Only lines with Finance Status = Approved contribute."),
        dict(key="VerAvoid", h="Verified Cost Avoidance ($, Finance-approved)", w=14, kind="calc", fmt=FMT_USD,
             f='=IF({ID}="","",SUMIFS(fin_Verified,fin_Proj,{ID},fin_Type,"Cost Avoidance"))'),
        dict(key="RAG", h="RAG Status", w=10, kind="calc",
             f='=IF({ID}="","",IF(OR({Status}="Complete",{Status}="Cancelled"),"CLOSED",IF({Status}="Pending Assignment","PENDING",'
               'IF(OR({Status}="Blocked",{Overdue}>=RedOverdueThr,AND(ISNUMBER({TargetDate}),{TargetDate}<AsOfDate)),"RED",'
               'IF(OR({Overdue}>0,{Risks}>0,{Leader}="",NOT(ISNUMBER({TargetDate})),AND({Stage}<>"",{Stage}<>"1-Define",NOT(ISNUMBER({Baseline})))),"AMBER","GREEN")))))',
             note="RED: Blocked, overdue actions >= threshold, or past target date. AMBER: any overdue action, open risk, no primary owner, no target date, or baseline missing after Define. PENDING: awaiting assignment. CLOSED: Complete/Cancelled."),
        dict(key="Reason", h="Attention Reasons", w=36, kind="calc", wrap=True,
             f='=IF({ID}="","",IF(OR({Status}="Complete",{Status}="Cancelled"),"",IF({Status}="Pending Assignment","Awaiting confirmed assignment; ","")'
               '&IF({Status}="Blocked","Project blocked; ","")&IF({Overdue}>0,{Overdue}&" overdue action(s); ","")'
               '&IF(AND(ISNUMBER({TargetDate}),{TargetDate}<AsOfDate),"Past target date; ","")&IF({Leader}="","No primary owner; ","")'
               '&IF(NOT(ISNUMBER({TargetDate})),"Target date not set; ","")&IF(AND({Stage}<>"",{Stage}<>"1-Define",NOT(ISNUMBER({Baseline}))),"Baseline missing; ","")'
               '&IF({Risks}>0,{Risks}&" open risk/escalation(s); ","")))'),
        dict(key="Attn", h="Requires Attention (1/0)", w=10, kind="calc", fmt=FMT_INT,
             f='=IF({ID}="","",IF(OR({RAG}="RED",{RAG}="AMBER",{RAG}="PENDING"),1,0))'),
        dict(key="OwnerChk", h="Owner Check", w=18, kind="calc",
             f='=IF({ID}="","",IF({Leader}="","Primary owner missing",IF({Leader}={CoOwner},"Leader = co-owner (fix)",IF(ISERROR(MATCH({Leader},md_PeopleName,0)),"Leader not in people list","OK"))))'),
        dict(key="WorInc", h="WOR Filter Match (1/0)", w=9, kind="calc", fmt=FMT_INT,
             f='=IF({ID}="",0,IF(AND(OR(WOR_FDept="(All)",WOR_FDept="",{Dept}=WOR_FDept),OR(WOR_FShift="(All)",WOR_FShift="",{Shift}=WOR_FShift),'
               'OR(WOR_FOwner="(All)",WOR_FOwner="",{Leader}=WOR_FOwner,{CoOwner}=WOR_FOwner),{Status}<>"Cancelled",OR(WOR_FIncl="Yes",{Status}<>"Complete")),1,0))',
             note="Helper for 06 Weekly WOR Report filters."),
        dict(key="WorSeq", h="WOR Sequence", w=9, kind="calc", fmt=FMT_INT,
             f='=IF({WorInc}=1,COUNTIF(${WorInc_c}$6:{WorInc_c}{r},1),"")'),
    ]


REG_CAP = 150
ACT_CAP = 1000
FIN_CAP = 300
WORU_CAP = 2000
CHR_CAP = 150
DEV_CAP = 300
RISK_CAP = 300
EVD_CAP = 1000


def action_cols():
    return [
        dict(key="ID", h="Action ID", w=10, kind="calc", f='=IF({Proj}="","","A-"&TEXT(ROW()-5,"0000"))',
             note="Generated when a Project ID is selected. Do not re-sort or delete rows."),
        dict(key="Proj", h="Project ID", w=12, kind="in", dv="lst_ProjectID"),
        dict(key="PTitle", h="Project Title", w=24, kind="calc", wrap=True,
             f='=IF({Proj}="","",IFERROR(INDEX(reg_Title,MATCH({Proj},reg_ID,0)),"Unknown Project ID"))'),
        dict(key="Desc", h="Action Description", w=40, kind="in", wrap=True, note="Verb + object + measurable outcome."),
        dict(key="Cat", h="Category", w=18, kind="in", dv="lst_ActCategory"),
        dict(key="MS", h="Milestone?", w=9, kind="in", dv="lst_YesNo",
             note="Yes = a charter milestone. Milestone completion is reported separately from action completion."),
        dict(key="Owner", h="Accountable Owner", w=18, kind="in", dv="lst_People", note="ONE accountable person."),
        dict(key="Support", h="Supporting Team", w=20, kind="in", wrap=True),
        dict(key="Priority", h="Priority", w=9, kind="in", dv="lst_Priority"),
        dict(key="Assigned", h="Assigned Date", w=11, kind="in", t="date"),
        dict(key="Due", h="Due Date", w=11, kind="in", t="date"),
        dict(key="CompDate", h="Completed Date", w=11, kind="in", t="date", note="Required for completion credit."),
        dict(key="Status", h="Status", w=12, kind="in", dv="lst_ActStatus"),
        dict(key="RepPct", h="Reported % Complete (info only)", w=11, kind="in", t="pct",
             note="Owner's estimate for discussion only. NOT used in any KPI (subjective percentages are never averaged)."),
        dict(key="EvidRef", h="Evidence Reference", w=20, kind="in", wrap=True,
             note="Short reference or link; record full evidence in 10 Evidence Register using this Action ID."),
        dict(key="Blocker", h="Blocker", w=22, kind="in", wrap=True),
        dict(key="Esc", h="Escalation Needed?", w=10, kind="in", dv="lst_YesNo",
             note="If Yes, also log in 09 Risks and Escalations so it is tracked to closure."),
        dict(key="Next", h="Next Step", w=26, kind="in", wrap=True),
        # AUTO
        dict(key="Credit", h="Credited Complete (1/0)", w=9, kind="calc", fmt=FMT_INT,
             f='=IF({Proj}="","",IF(AND({Status}="Complete",ISNUMBER({CompDate})),1,0))',
             note="1 only when Status = Complete AND Completed Date entered."),
        dict(key="Days", h="Days Remaining", w=9, kind="calc", fmt='0;[Red]-0',
             f='=IF(OR({Proj}="",NOT(ISNUMBER({Due}))),"",IF(OR({Credit}=1,{Status}="Cancelled"),"",{Due}-AsOfDate))',
             note="Due date minus the As-Of date. Negative = days overdue."),
        dict(key="Flag", h="Schedule Status", w=13, kind="calc",
             f='=IF({Proj}="","",IF({Status}="Cancelled","Cancelled",IF({Credit}=1,IF(ISNUMBER({Due}),IF({CompDate}<={Due},"Closed On Time","Closed Late"),"Closed"),'
               'IF(NOT(ISNUMBER({Due})),"NO DUE DATE",IF({Due}<AsOfDate,"OVERDUE",IF({Due}-AsOfDate<=DueSoonDays,"DUE SOON","On Track"))))))',
             note="OVERDUE = not credited complete, not cancelled, and due date before the As-Of date."),
        dict(key="OnTime", h="Closed On Time (1/0)", w=9, kind="calc", fmt=FMT_INT,
             f='=IF({Credit}=1,IF(AND(ISNUMBER({Due}),{CompDate}<={Due}),1,0),"")'),
        dict(key="EvCount", h="Evidence Records", w=9, kind="calc", fmt=FMT_INT,
             f='=IF({ID}="","",COUNTIFS(evd_Action,{ID}))'),
        dict(key="EvCheck", h="Evidence Check", w=16, kind="calc",
             f='=IF({ID}="","",IF({Credit}=1,IF(AND({EvCount}=0,{EvidRef}=""),"MISSING EVIDENCE",IF(COUNTIFS(evd_Action,{ID},evd_Verif,"Verified")>0,"Verified","Not yet verified")),""))'),
        dict(key="DCheck", h="Data Check", w=22, kind="calc", wrap=True,
             f='=IF({Proj}="","",IF(ISERROR(MATCH({Proj},reg_ID,0)),"Unknown Project ID",IF(AND({Status}="Complete",NOT(ISNUMBER({CompDate}))),"Enter completed date",'
               'IF(AND(ISNUMBER({CompDate}),{Status}<>"Complete"),"Completed date but status not Complete",IF({Owner}="","No accountable owner",'
               'IF(AND(ISNUMBER({CompDate}),{CompDate}>AsOfDate),"Completed date after As-Of date",IF({Status}="","Status not set","OK")))))))'),
        dict(key="OpenMS", h="Open Milestone (1/0)", w=9, kind="calc", fmt=FMT_INT,
             f='=IF({Proj}="",0,IF(AND({MS}="Yes",{Credit}=0,{Status}<>"Cancelled",ISNUMBER({Due})),1,0))'),
        dict(key="MSKey", h="Open Milestone Key", w=10, kind="calc",
             f='=IF({OpenMS}=1,{Proj}&"|"&{Due},"")'),
        dict(key="MSSeq", h="Charter Milestone Seq", w=9, kind="calc", fmt=FMT_INT,
             f='=IF(AND({Proj}<>"",{MS}="Yes",{Status}<>"Cancelled"),COUNTIFS(${Proj_c}$6:{Proj_c}{r},{Proj},${MS_c}$6:{MS_c}{r},"Yes",${Status_c}$6:{Status_c}{r},"<>Cancelled"),"")'),
        dict(key="ChKey", h="Charter Milestone Key", w=10, kind="calc",
             f='=IF({MSSeq}="","",{Proj}&"|"&{MSSeq})'),
    ]


def finance_cols():
    W = '"Waste - Material Loss"'
    C = '"Changeover - Time Reduction"'
    return [
        dict(key="ID", h="Benefit ID", w=10, kind="calc", f='=IF({Proj}="","","B-"&TEXT(ROW()-5,"0000"))'),
        dict(key="Proj", h="Project ID", w=12, kind="in", dv="lst_ProjectID"),
        dict(key="PTitle", h="Project Title", w=22, kind="calc", wrap=True,
             f='=IF({Proj}="","",IFERROR(INDEX(reg_Title,MATCH({Proj},reg_ID,0)),"Unknown Project ID"))'),
        dict(key="Type", h="Benefit Type", w=15, kind="in", dv="lst_BenefitType",
             note="Hard Savings = cash/P&L reduction. Cost Avoidance = prevented future cost. Productivity Capacity = time/capacity freed (NOT cash unless Finance converts it)."),
        dict(key="Method", h="Calculation Method", w=20, kind="in", dv="lst_CalcMethod"),
        dict(key="Desc", h="Benefit Description and Assumptions", w=34, kind="in", wrap=True,
             note="State every assumption explicitly (what changed, why the benefit is attributable, data used)."),
        dict(key="Source", h="Data Source / Evidence Ref", w=20, kind="in", wrap=True,
             note="Required. SAP report, scale data, Finance file, etc. Benefits without supporting data are flagged."),
        dict(key="Key", h="Benefit Source Key (double-count control)", w=18, kind="in",
             note="Identify the cost pool, e.g. 'LINE3-FILM-SCRAP' or a cost center + loss driver. Two lines with the same key share one pool; their attribution % cannot exceed 100%."),
        dict(key="Attr", h="Attribution % to this Project", w=11, kind="in", t="pct",
             note="Share of the benefit pool claimed by this line. Blank = 100%."),
        dict(key="BasePeriod", h="Baseline Period", w=14, kind="in", note="e.g. 'Jan-Mar 2026 (13 wks)'."),
        dict(key="PWeeks", h="Measurement Period Length (weeks)", w=11, kind="in", t="num",
             note="Length of the period the per-period values below represent. Used for annualization."),
        dict(key="UnitM", h="Measurement Unit", w=10, kind="in", note="e.g. lbs, units, cases, minutes."),
        dict(key="BVol", h="Baseline Production Volume", w=12, kind="in", t="num"),
        dict(key="CVol", h="Comparison Production Volume", w=12, kind="in", t="num",
             note="Volume in the current measurement period. Waste savings are calculated on this volume (volume-normalized)."),
        dict(key="BRate", h="Baseline Loss Rate", w=10, kind="in", t="num", fmt="0.00%",
             note="Loss per unit produced (fraction, e.g. 3.5% = 0.035)."),
        dict(key="CRate", h="Current Loss Rate", w=10, kind="in", t="num", fmt="0.00%"),
        dict(key="UCost", h="Validated Material Cost per Unit ($)", w=11, kind="in", t="num", fmt=FMT_USD2),
        dict(key="COMin", h="Changeover Minutes Reduced (per changeover)", w=11, kind="in", t="num"),
        dict(key="COFreq", h="Changeovers per Period", w=10, kind="in", t="num"),
        dict(key="Demo", h="Capacity / Cost Benefit Demonstrated?", w=12, kind="in", dv="lst_YesNo",
             note="Yes ONLY when Finance agrees recovered minutes produce real output or cost reduction. Otherwise the line is valued at $0 cash and shown as capacity hours."),
        dict(key="RVal", h="Finance-Approved Recoverable Value ($/min)", w=12, kind="in", t="num", fmt=FMT_USD2),
        dict(key="Labor", h="Labor Cost Savings per Period ($)", w=12, kind="in", t="num", fmt=FMT_USD),
        dict(key="Down", h="Downtime Cost Savings per Period ($)", w=12, kind="in", t="num", fmt=FMT_USD),
        dict(key="Rework", h="Error / Rework Cost Savings per Period ($)", w=12, kind="in", t="num", fmt=FMT_USD,
             note="SAP confirmation: verified rework, correction labor or error cost - only when supported by data."),
        dict(key="Impl", h="Implementation Cost ($)", w=12, kind="in", t="num", fmt=FMT_USD),
        dict(key="OneTime", h="One-Time Benefit ($)", w=12, kind="in", t="num", fmt=FMT_USD),
        dict(key="Start", h="Benefit Start Date", w=11, kind="in", t="date", note="Date the improvement took effect (forecast run-rate starts this month)."),
        dict(key="Status", h="Finance Status", w=13, kind="in", dv="lst_FinStatus",
             note="Not Submitted / Under Review / Approved / Rejected. Only Approved lines count as verified savings."),
        dict(key="Reviewer", h="Finance Reviewer", w=14, kind="in", dv="lst_People"),
        dict(key="RevDate", h="Finance Review Date", w=11, kind="in", t="date"),
        dict(key="Realized", h="Realized Savings to Date - Finance Validated ($)", w=13, kind="in", t="num", fmt=FMT_USD,
             note="FINANCE ONLY. Project-attributable amount actually realized and validated by Finance."),
        dict(key="FinComments", h="Finance Comments", w=24, kind="in", wrap=True),
        # AUTO
        dict(key="VolChk", h="Volume Comparability", w=18, kind="calc", wrap=True,
             f='=IF({Proj}="","",IF(AND(ISNUMBER({BVol}),ISNUMBER({CVol}),N({BVol})>0),IF(ABS({CVol}/{BVol}-1)>VolTol,"Volume shift > "&TEXT(VolTol,"0%")&" - Finance to confirm","Comparable"),IF({Method}=' + W + ',"N/A - volumes missing","")))'),
        dict(key="Mat", h="Material Savings / Period ($)", w=12, kind="calc", fmt=FMT_USD,
             f='=IF({Proj}="","",IF({Method}<>' + W + ',"",IF(AND(ISNUMBER({BRate}),ISNUMBER({CRate}),ISNUMBER({CVol}),ISNUMBER({UCost})),({BRate}-{CRate})*{CVol}*{UCost},"N/A")))',
             note="(Baseline loss rate - Current loss rate) x Comparison volume x Validated unit cost."),
        dict(key="CapHrs", h="Capacity Hours Recovered / Period", w=11, kind="calc", fmt="#,##0.0",
             f='=IF({Proj}="","",IF(AND(ISNUMBER({COMin}),ISNUMBER({COFreq})),{COMin}*{COFreq}/60,""))'),
        dict(key="CODollar", h="Changeover Value / Period ($)", w=12, kind="calc", fmt=FMT_USD,
             f='=IF({Proj}="","",IF({Method}<>' + C + ',"",IF(NOT(AND(ISNUMBER({COMin}),ISNUMBER({COFreq}))),"N/A",IF(AND({Demo}="Yes",ISNUMBER({RVal})),{COMin}*{COFreq}*{RVal},0))))',
             note="Minutes reduced x changeovers x Finance-approved $/min, ONLY when benefit demonstrated = Yes. Otherwise $0 (recovered minutes are not automatic cash)."),
        dict(key="Recur", h="Recurring Benefit / Period ($)", w=12, kind="calc", fmt=FMT_USD,
             f='=IF({Proj}="","",IF(COUNT({Mat},{CODollar},{Labor},{Down},{Rework})=0,"N/A",SUM({Mat},{CODollar},{Labor},{Down},{Rework})))'),
        dict(key="PPY", h="Periods per Year", w=9, kind="calc", fmt="0.00",
             f='=IF({Proj}="","",IF(AND(ISNUMBER({PWeeks}),N({PWeeks})>0),OpWeeks/{PWeeks},"N/A"))'),
        dict(key="Annual", h="Annualized Recurring ($/yr)", w=13, kind="calc", fmt=FMT_USD,
             f='=IF({Proj}="","",IF(AND(ISNUMBER({Recur}),ISNUMBER({PPY})),{Recur}*{PPY},"N/A"))'),
        dict(key="Gross", h="Gross Savings - Annual Forecast ($)", w=13, kind="calc", fmt=FMT_USD,
             f='=IF({Proj}="","",IF(AND(NOT(ISNUMBER({Annual})),NOT(ISNUMBER({OneTime}))),"N/A",N({Annual})+N({OneTime})))',
             note="Annualized recurring + one-time benefit (first 12 months)."),
        dict(key="Net", h="Net Savings - Annual Forecast ($)", w=13, kind="calc", fmt=FMT_USD,
             f='=IF({Proj}="","",IF(ISNUMBER({Gross}),{Gross}-N({Impl}),"N/A"))'),
        dict(key="EAttr", h="Effective Attribution", w=9, kind="calc", fmt=FMT_PCT,
             f='=IF({Proj}="","",IF(ISNUMBER({Attr}),{Attr},1))'),
        dict(key="Fcst", h="Attributed Net Forecast ($/yr)", w=13, kind="calc", fmt=FMT_USD,
             f='=IF({Proj}="","",IF(ISNUMBER({Net}),{Net}*{EAttr},"N/A"))', note="FORECAST - unverified unless Finance Status = Approved."),
        dict(key="FcstNum", h="Forecast (numeric helper)", w=10, kind="calc", fmt=FMT_USD,
             f='=IF(ISNUMBER({Fcst}),{Fcst},0)', note="Numeric copy used by summaries (N/A counts as 0 in sums but is flagged)."),
        dict(key="Verified", h="Verified Savings - Finance Approved ($)", w=13, kind="calc", fmt=FMT_USD,
             f='=IF({Proj}="","",IF(AND({Status}="Approved",ISNUMBER({Realized})),{Realized},0))',
             note="= Realized (Finance-validated) ONLY when Finance Status = Approved; otherwise 0. Unapproved benefits never enter verified savings."),
        dict(key="CapHrsYr", h="Capacity Hours / yr (attributed)", w=11, kind="calc", fmt="#,##0.0",
             f='=IF(AND(ISNUMBER({CapHrs}),ISNUMBER({PPY})),{CapHrs}*{PPY}*{EAttr},0)'),
        dict(key="Missing", h="Missing Information / Assumption Flags", w=36, kind="calc", wrap=True,
             f='=IF({Proj}="","",IF(({Type}=""),"Benefit type; ","")&IF({Method}="","Calculation method; ","")&IF(NOT(ISNUMBER({PWeeks})),"Period length; ","")'
               '&IF(AND({Method}=' + W + ',OR(NOT(ISNUMBER({BRate})),NOT(ISNUMBER({CRate})),NOT(ISNUMBER({CVol})),NOT(ISNUMBER({UCost})))),"Waste inputs (loss rates / volume / unit cost); ","")'
               '&IF(AND({Method}=' + C + ',OR(NOT(ISNUMBER({COMin})),NOT(ISNUMBER({COFreq})))),"Changeover minutes / frequency; ","")'
               '&IF(AND({Method}=' + C + ',{Demo}<>"Yes"),"Benefit not demonstrated - valued at $0 cash; ","")'
               '&IF(AND({Method}=' + C + ',{Demo}="Yes",NOT(ISNUMBER({RVal}))),"Recoverable $/min not approved; ","")'
               '&IF({Source}="","Data source / evidence; ","")&IF(NOT(ISNUMBER({Start})),"Benefit start date; ","")'
               '&IF(AND({Status}="Approved",NOT(ISNUMBER({RevDate}))),"Finance review date; ","")'
               '&IF(AND({Status}="Approved",NOT(ISNUMBER({Realized}))),"Realized amount not entered; ","")'
               '&IF({Status}="","Finance status; ",""))'),
        dict(key="Dbl", h="Double-Count Check", w=20, kind="calc", wrap=True,
             f='=IF({Proj}="","",IF({Key}="","No source key - confirm no overlap",IF(SUMIFS(fin_EAttr,fin_Key,{Key},fin_Status,"<>Rejected")>1.0001,"OVER-ATTRIBUTED >100% - resolve",IF(COUNTIFS(fin_Key,{Key})>1,"Shared pool - attribution OK","Unique"))))'),
        dict(key="StartIdx", h="Start Month Index", w=9, kind="calc", fmt="0",
             f='=IF(ISNUMBER({Start}),YEAR({Start})*12+MONTH({Start}),99999)'),
        dict(key="MF", h="Monthly Hard Run-Rate (helper)", w=10, kind="calc", fmt=FMT_USD,
             f='=IF(AND({Proj}<>"",{Type}="Hard Savings",{Status}<>"Rejected",ISNUMBER({Annual})),{Annual}*{EAttr}/12,0)',
             note="Annualized recurring x attribution / 12. Used for the cumulative forecast trend (excludes one-time benefit and implementation cost)."),
    ]


def woru_cols():
    return [
        dict(key="ID", h="Update ID", w=10, kind="calc", f='=IF({Proj}="","","U-"&TEXT(ROW()-5,"0000"))'),
        dict(key="Week", h="Week Ending", w=11, kind="in", t="date", note="WOR week-ending date. One row per project per week."),
        dict(key="Proj", h="Project ID", w=12, kind="in", dv="lst_ProjectID"),
        dict(key="PTitle", h="Project Title", w=24, kind="calc", wrap=True,
             f='=IF({Proj}="","",IFERROR(INDEX(reg_Title,MATCH({Proj},reg_ID,0)),"Unknown Project ID"))'),
        dict(key="By", h="Submitted By", w=16, kind="in", dv="lst_People"),
        dict(key="Done", h="Work Completed This Week", w=40, kind="in", wrap=True),
        dict(key="Next", h="Next-Week Commitments", w=40, kind="in", wrap=True),
        dict(key="KPI", h="KPI Actual (this week)", w=11, kind="in", t="num",
             note="Measured value of the project's primary KPI. Leave blank if not measured (never enter 0 for missing)."),
        dict(key="Risks", h="Risks / Issues", w=28, kind="in", wrap=True),
        dict(key="Support", h="Support Needed", w=28, kind="in", wrap=True),
        dict(key="DecReq", h="Leadership Decision Requested", w=28, kind="in", wrap=True),
        dict(key="DecRec", h="Leadership Decision Recorded", w=28, kind="in", wrap=True),
        dict(key="Key", h="Update Key", w=12, kind="calc",
             f='=IF(OR({Proj}="",NOT(ISNUMBER({Week}))),"",{Proj}&"|"&{Week})'),
        dict(key="Dup", h="Duplicate Check", w=14, kind="calc",
             f='=IF({Key}="","",IF(COUNTIF(wor_Key,{Key})>1,"DUPLICATE week","OK"))'),
    ]


CHARTER_TEXT = [("Problem", "Business Problem"), ("ScopeIn", "Scope - In"), ("ScopeOut", "Scope - Out"),
                ("Current", "Current Condition"), ("SMART", "SMART Objective"), ("Stake", "Key Stakeholders"),
                ("Resp", "Team Responsibilities"), ("Benefits", "Expected Benefits (narrative)"),
                ("Constraints", "Constraints"), ("CRisks", "Key Risks")]


def charter_cols():
    cols = [
        dict(key="Proj", h="Project ID", w=12, kind="in", dv="lst_ProjectID", note="One charter record per project."),
        dict(key="PTitle", h="Project Title", w=24, kind="calc", wrap=True,
             f='=IF({Proj}="","",IFERROR(INDEX(reg_Title,MATCH({Proj},reg_ID,0)),"Unknown Project ID"))'),
        dict(key="Version", h="Charter Version", w=9, kind="in"),
        dict(key="CStatus", h="Charter Status", w=13, kind="in", dv="lst_CharterStatus"),
    ]
    for k, h in CHARTER_TEXT:
        cols.append(dict(key=k, h=h, w=34, kind="in", wrap=True))
    cols += [
        dict(key="LeadDate", h="Project Leader Sign Date", w=11, kind="in", t="date"),
        dict(key="SponsorDate", h="Sponsor Approval Date", w=11, kind="in", t="date"),
        dict(key="CoachDate", h="CI Coach Review Date", w=11, kind="in", t="date"),
        dict(key="FinDate", h="Finance Review Date", w=11, kind="in", t="date"),
        dict(key="Complete", h="Charter Completeness %", w=11, kind="calc", fmt=FMT_PCT,
             f='=IF({Proj}="","",(COUNTA({Problem_c}{r}:{CRisks_c}{r})+COUNTIFS(reg_ID,{Proj},reg_Baseline,"<>")+COUNTIFS(reg_ID,{Proj},reg_Target,"<>")+COUNTIFS(reg_ID,{Proj},reg_TargetDate,"<>"))/13)',
             note="Filled fields / 13: the 10 narrative sections plus baseline, target and target date in the register. Objective measure of charter completeness (quality is scored separately in 07)."),
        dict(key="Dup", h="Duplicate Check", w=12, kind="calc",
             f='=IF({Proj}="","",IF(COUNTIF(chr_Proj,{Proj})>1,"DUPLICATE","OK"))'),
    ]
    return cols


COMPETENCIES = [
    ("Participation", "Participation and Ownership"),
    ("CharterQ", "Charter Quality"),
    ("ProbDef", "Problem Definition"),
    ("RCA", "Root Cause Analysis"),
    ("Data", "Use of Data"),
    ("Teamwork", "Teamwork and Collaboration"),
    ("Closure", "Action Closure Discipline"),
    ("Comms", "Communication (WOR)"),
    ("Sustain", "Sustainability and Standard Work"),
]


def dev_cols():
    cols = [
        dict(key="ID", h="Review ID", w=10, kind="calc", f='=IF({Sup}="","","D-"&TEXT(ROW()-5,"0000"))'),
        dict(key="Sup", h="Supervisor", w=20, kind="in", dv="lst_People"),
        dict(key="Role", h="Role", w=18, kind="calc",
             f='=IF({Sup}="","",IFERROR(IF(INDEX(md_PeopleRole,MATCH({Sup},md_PeopleName,0))="","Not set",INDEX(md_PeopleRole,MATCH({Sup},md_PeopleName,0))),"Not in people list"))'),
        dict(key="Proj", h="Project ID (context)", w=12, kind="in", dv="lst_ProjectID"),
        dict(key="Date", h="Review Date", w=11, kind="in", t="date"),
        dict(key="Reviewer", h="Reviewer / Coach", w=16, kind="in", dv="lst_People"),
    ]
    for k, h in COMPETENCIES:
        cols.append(dict(key=k, h=h + " (1-4)", w=11, kind="in", t="score",
                         note="Score 1-4 per the rubric on '07A Rubric and Summary'. Leave blank if not observed."))
    sc0 = cols[6]["key"]
    sc1 = cols[6 + len(COMPETENCIES) - 1]["key"]
    cols += [
        dict(key="Strengths", h="Observed Strengths", w=26, kind="in", wrap=True),
        dict(key="Coaching", h="Coaching Notes", w=30, kind="in", wrap=True),
        dict(key="DevActions", h="Development Actions", w=30, kind="in", wrap=True),
        dict(key="NextRev", h="Next Review Date", w=11, kind="in", t="date"),
        dict(key="NScored", h="Competencies Scored", w=10, kind="calc", fmt=FMT_INT,
             f='=IF({Sup}="","",COUNT({' + sc0 + '_c}{r}:{' + sc1 + '_c}{r}))'),
        dict(key="Avg", h="Average Competency (1-4)", w=10, kind="calc", fmt="0.0",
             f='=IF({Sup}="","",IF({NScored}=0,"Not assessed",AVERAGE({' + sc0 + '_c}{r}:{' + sc1 + '_c}{r})))',
             note="Simple average of the competencies scored (unscored areas excluded, not treated as zero)."),
        dict(key="Level", h="Overall Level", w=13, kind="calc",
             f='=IF({Sup}="","",IF(ISNUMBER({Avg}),IF({Avg}>=3.5,"4 - Role Model",IF({Avg}>=2.5,"3 - Proficient",IF({Avg}>=1.5,"2 - Developing","1 - Beginning"))),"Not assessed"))'),
        dict(key="Focus", h="Priority Development Area", w=22, kind="calc", wrap=True,
             f='=IF(OR({Sup}="",{NScored}=0,{NScored}=""),"",SUBSTITUTE(INDEX(${' + sc0 + '_c}$5:${' + sc1 + '_c}$5,MATCH(MIN({' + sc0 + '_c}{r}:{' + sc1 + '_c}{r}),{' + sc0 + '_c}{r}:{' + sc1 + '_c}{r},0))," (1-4)",""))',
             note="Lowest-scored competency (first one if tied) - the suggested coaching focus."),
        dict(key="RevFlag", h="Next Review Status", w=13, kind="calc",
             f='=IF({Sup}="","",IF(ISNUMBER({NextRev}),IF({NextRev}<AsOfDate,"REVIEW OVERDUE",IF({NextRev}-AsOfDate<=14,"Due within 14 days","Scheduled")),"Not scheduled"))'),
        dict(key="Key", h="Review Key", w=10, kind="calc",
             f='=IF(OR({Sup}="",NOT(ISNUMBER({Date}))),"",{Sup}&"|"&{Date})'),
    ]
    return cols


def risk_cols():
    return [
        dict(key="ID", h="Risk ID", w=9, kind="calc", f='=IF({Proj}="","","R-"&TEXT(ROW()-5,"0000"))'),
        dict(key="Proj", h="Project ID", w=12, kind="in", dv="lst_ProjectID"),
        dict(key="PTitle", h="Project Title", w=22, kind="calc", wrap=True,
             f='=IF({Proj}="","",IFERROR(INDEX(reg_Title,MATCH({Proj},reg_ID,0)),"Unknown Project ID"))'),
        dict(key="Raised", h="Date Raised", w=11, kind="in", t="date"),
        dict(key="Type", h="Type", w=12, kind="in", dv="lst_RiskType"),
        dict(key="Desc", h="Blocker / Risk Description", w=36, kind="in", wrap=True),
        dict(key="Impact", h="Impact", w=9, kind="in", dv="lst_Impact"),
        dict(key="Owner", h="Responsible Owner", w=18, kind="in", dv="lst_People"),
        dict(key="Support", h="Required Support", w=28, kind="in", wrap=True),
        dict(key="Level", h="Escalation Level", w=16, kind="in", dv="lst_EscLevel"),
        dict(key="EscDate", h="Escalation Date", w=11, kind="in", t="date"),
        dict(key="Target", h="Resolution Target Date", w=11, kind="in", t="date"),
        dict(key="Status", h="Closure Status", w=11, kind="in", dv="lst_RiskStatus"),
        dict(key="Closed", h="Closed Date", w=11, kind="in", t="date"),
        dict(key="Notes", h="Resolution Notes", w=30, kind="in", wrap=True),
        dict(key="IsOpen", h="Open (1/0)", w=8, kind="calc", fmt=FMT_INT,
             f='=IF({Proj}="","",IF(OR({Status}="Resolved",{Status}="Closed"),0,1))', note="Blank status is treated as Open."),
        dict(key="DaysOpen", h="Days Open", w=8, kind="calc", fmt=FMT_INT,
             f='=IF({Proj}="","",IF(ISNUMBER({Raised}),IF(ISNUMBER({Closed}),{Closed},AsOfDate)-{Raised},"N/A"))'),
        dict(key="Past", h="Past Target?", w=11, kind="calc",
             f='=IF({Proj}="","",IF(AND({IsOpen}=1,ISNUMBER({Target})),IF({Target}<AsOfDate,"PAST TARGET","Within target"),IF({IsOpen}=1,"No target date","")))'),
        dict(key="Lead", h="Leadership Action Needed", w=14, kind="calc",
             f='=IF({Proj}="","",IF(AND({IsOpen}=1,OR({Level}="Plant Leadership",{Impact}="High",{Past}="PAST TARGET",{Status}="Escalated")),"LEADERSHIP ACTION",""))',
             note="Open item that is High impact, escalated, escalated to Plant Leadership, or past its resolution target."),
    ]


def evidence_cols():
    return [
        dict(key="ID", h="Evidence ID", w=10, kind="calc", f='=IF({Proj}="","","E-"&TEXT(ROW()-5,"0000"))'),
        dict(key="Proj", h="Project ID", w=12, kind="in", dv="lst_ProjectID"),
        dict(key="Action", h="Action ID", w=10, kind="in", dv="lst_ActionID", note="Optional - link evidence to a specific action."),
        dict(key="Type", h="Evidence Type", w=18, kind="in", dv="lst_EvidenceType"),
        dict(key="Title", h="Document Title", w=30, kind="in", wrap=True),
        dict(key="URL", h="File / SharePoint / Teams Link", w=40, kind="in",
             note="Paste the full https:// link (SharePoint/Teams: use 'Copy link') or a \\\\server\\share path. Files are NOT stored in Excel."),
        dict(key="By", h="Uploaded By", w=16, kind="in", dv="lst_People"),
        dict(key="Date", h="Evidence Date", w=11, kind="in", t="date"),
        dict(key="Verif", h="Verification Status", w=14, kind="in", dv="lst_Verif"),
        dict(key="Reviewer", h="Reviewer", w=16, kind="in", dv="lst_People"),
        dict(key="RevDate", h="Review Date", w=11, kind="in", t="date"),
        dict(key="Notes", h="Notes", w=26, kind="in", wrap=True),
        dict(key="Open", h="Open Link", w=10, kind="calc", f='=IF({URL}="","",HYPERLINK({URL},"Open link"))'),
        dict(key="LinkChk", h="Link Check", w=14, kind="calc",
             f='=IF({Proj}="","",IF({URL}="","NO LINK",IF(OR(LEFT({URL},8)="https://",LEFT({URL},7)="http://",LEFT({URL},2)="\\\\"),"Format OK","Check format")))'),
        dict(key="ActChk", h="Action Check", w=16, kind="calc",
             f='=IF({Action}="","",IF(ISERROR(MATCH({Action},act_ID,0)),"Unknown Action ID",IF(INDEX(act_Proj,MATCH({Action},act_ID,0))<>{Proj},"Action is on another project","OK")))'),
    ]


# ============================================================================= DATA (production)
PEOPLE_PROD = [
    ["Eric Lujan", "Supervisor"], ["Manuela Garcia", "Supervisor"], ["Jhirmaul Carter", "Supervisor"],
    ["Stacy Jackson", "Supervisor"], ["Phillip Harris", "Supervisor"], ["Ivan Lopez", "Supervisor"],
    ["Moe Abendeh", "CI Manager", "Continuous Improvement"], ["Randy", "Finance Reviewer", "Finance"],
]

LAUNCH_NOTE = ("Assignment confirmed at CI launch. Leader = first-named owner (confirm primary owner). Department, area, "
               "sponsor, dates, KPI, baseline and target are blank until confirmed - do not estimate.")

PROJECTS_PROD = [
    dict(Title="SAP Confirmation Process",
         Desc="Review current SAP confirmation practices, identify gaps, and improve accuracy and consistency.",
         Category="SAP / Systems Accuracy", Leader="Eric Lujan", CoOwner="Manuela Garcia", Coach="Moe Abendeh",
         FinRev="Randy", Stage="1-Define", Status="Active", Notes=LAUNCH_NOTE),
    dict(Title="Changeover Standard",
         Desc="Standardize changeover procedures, identify best practices, and reduce downtime.",
         Category="Changeover / SMED", Leader="Jhirmaul Carter", CoOwner="Stacy Jackson", Coach="Moe Abendeh",
         FinRev="Randy", Stage="1-Define", Status="Active", Notes=LAUNCH_NOTE),
    dict(Title="Waste Reduction",
         Desc="Identify waste sources, implement improvements, and track material utilization.",
         Category="Waste Reduction", Leader="Phillip Harris", CoOwner="Ivan Lopez", Coach="Moe Abendeh",
         FinRev="Randy", Stage="1-Define", Status="Active", Notes=LAUNCH_NOTE),
    dict(Title="Maintenance Projects - awaiting confirmed assignments",
         Desc="Placeholder for Maintenance supervisor-led project(s). Owners, scope and dates to be confirmed.",
         Category="Maintenance Reliability", Dept="Maintenance", Coach="Moe Abendeh", Status="Pending Assignment",
         Notes="Awaiting confirmed assignment. Replace the title and complete the record when the owner is confirmed."),
    dict(Title="Warehouse Projects - awaiting confirmed assignments",
         Desc="Placeholder for Warehouse supervisor-led project(s). Owners, scope and dates to be confirmed.",
         Category="Warehouse / Logistics", Dept="Warehouse", Coach="Moe Abendeh", Status="Pending Assignment",
         Notes="Awaiting confirmed assignment. Replace the title and complete the record when the owner is confirmed."),
]


def prod_data():
    return dict(
        people=PEOPLE_PROD,
        projects=PROJECTS_PROD,
        actions=[], finance=[], woru=[], risks=[], evidence=[],
        charters=[dict(Proj=f"DUN-CI-{i:03d}", CStatus="Draft") for i in range(1, 6)],
        dev=[dict(Sup=p[0]) for p in PEOPLE_PROD[:6]],
    )


# ============================================================================= DATA (demo)
def d(y, m, dd):
    return dt.datetime(y, m, dd)


def demo_data():
    people = [["Demo Supervisor A", "Production Supervisor", "Production", "A"],
              ["Demo Supervisor B", "Production Supervisor", "Production", "B"],
              ["Demo Supervisor C", "Production Supervisor", "Production", "C"],
              ["Demo Maint Sup D", "Maintenance Supervisor", "Maintenance", "A"],
              ["Demo Whse Sup E", "Warehouse Supervisor", "Warehouse", "B"],
              ["Demo Supervisor F", "Production Supervisor", "Production", "B"],
              ["Demo CI Coach", "CI Manager", "Continuous Improvement", "Cross-Shift"],
              ["Demo Finance", "Finance Reviewer", "Finance", "Cross-Shift"],
              ["Demo Sponsor", "Sponsor / Manager", "Production", "Cross-Shift"]]
    P = []

    def proj(**k):
        base = dict(Coach="Demo CI Coach", FinRev="Demo Finance", Sponsor="Demo Sponsor", Area="Plant-wide",
                    Equipment="Not equipment-specific", Priority="Medium")
        base.update(k)
        P.append(base)

    proj(Title="DEMO - Line scrap reduction", Desc="Fictitious demo project.", Category="Waste Reduction",
         Dept="Production", Shift="A", Leader="Demo Supervisor A", CoOwner="Demo Supervisor F", Stage="4-Improve",
         Status="Active", Start=d(2026, 8, 3), TargetDate=d(2027, 1, 29), KPI="Scrap rate", Unit="%", Baseline=4.0,
         Target=2.5, Priority="High")
    proj(Title="DEMO - Changeover standard work", Desc="Fictitious demo project.", Category="Changeover / SMED",
         Dept="Production", Shift="B", Leader="Demo Supervisor B", CoOwner="Demo Supervisor C", Stage="3-Analyze",
         Status="Active", Start=d(2026, 9, 1), TargetDate=d(2027, 2, 26), KPI="Avg changeover minutes", Unit="min",
         Baseline=95, Target=60)
    proj(Title="DEMO - SAP confirmation accuracy", Desc="Fictitious demo project.", Category="SAP / Systems Accuracy",
         Dept="Production", Shift="C", Leader="Demo Supervisor C", Stage="2-Measure", Status="Blocked",
         Start=d(2026, 9, 14), TargetDate=d(2026, 11, 13), KPI="Confirmation errors / week", Unit="errors")
    proj(Title="DEMO - PM compliance", Desc="Fictitious demo project.", Category="Maintenance Reliability",
         Dept="Maintenance", Shift="Cross-Shift", Leader="Demo Maint Sup D", Stage="5-Control", Status="Complete",
         Start=d(2026, 5, 4), TargetDate=d(2026, 10, 30), CompDate=d(2026, 10, 28), KPI="PM compliance", Unit="%",
         Baseline=72, Target=90)
    proj(Title="DEMO - Pick accuracy", Desc="Fictitious demo project.", Category="Warehouse / Logistics",
         Dept="Warehouse", Shift="B", Leader="Demo Whse Sup E", Stage="1-Define", Status="Not Started",
         TargetDate=d(2027, 3, 31))
    proj(Title="DEMO - Unassigned idea", Desc="Fictitious demo project.", Category="Other", Status="Pending Assignment")

    A = []

    def act(p, desc, owner, due, status, comp=None, ms="No", assigned=d(2026, 9, 1), evid=None, rep=None):
        A.append(dict(Proj=p, Desc=desc, Cat="Countermeasure / Implementation", MS=ms, Owner=owner, Priority="Medium",
                      Assigned=assigned, Due=due, CompDate=comp, Status=status, EvidRef=evid, RepPct=rep))

    p1, p2, p3, p4, p5 = [f"DEMO-{i:03d}" for i in range(1, 6)]
    act(p1, "Collect 4 weeks scrap data", "Demo Supervisor A", d(2026, 9, 18), "Complete", d(2026, 9, 17), "Yes", evid="Scrap log")
    act(p1, "Pareto scrap by cause", "Demo Supervisor A", d(2026, 10, 2), "Complete", d(2026, 10, 6), "Yes")
    act(p1, "Trial film tension fix", "Demo Supervisor F", d(2026, 10, 30), "Complete", d(2026, 10, 29))
    act(p1, "Update SOP", "Demo Supervisor F", d(2026, 11, 13), "In Progress", rep=0.6)
    act(p1, "Train all shifts", "Demo Supervisor A", d(2026, 11, 24), "Not Started", ms="Yes")
    act(p1, "Verify 4 weeks sustained", "Demo Supervisor A", d(2027, 1, 15), "Not Started", ms="Yes")
    act(p2, "Video 3 changeovers", "Demo Supervisor B", d(2026, 10, 9), "Complete", d(2026, 10, 9), "Yes")
    act(p2, "SMED internal/external split", "Demo Supervisor C", d(2026, 10, 23), "In Progress", rep=0.8)
    act(p2, "Draft changeover checklist", "Demo Supervisor B", d(2026, 11, 6), "Not Started", ms="Yes")
    act(p2, "Pre-stage tooling cart", "Demo Supervisor C", d(2026, 12, 4), "Not Started")
    act(p2, "Cancelled idea", "Demo Supervisor C", d(2026, 10, 1), "Cancelled")
    act(p3, "Map confirmation process", "Demo Supervisor C", d(2026, 10, 2), "Complete", None)  # complete w/o date
    act(p3, "Pull SAP error report", "Demo Supervisor C", d(2026, 10, 16), "Blocked")
    act(p3, "Interview shift leads", "Demo Supervisor C", d(2026, 11, 5), "In Progress")
    act(p3, "Define error categories", "Demo Supervisor C", d(2026, 11, 20), "Not Started", ms="Yes")
    act(p4, "Baseline PM compliance", "Demo Maint Sup D", d(2026, 6, 1), "Complete", d(2026, 5, 29), "Yes")
    act(p4, "PM schedule rebuild", "Demo Maint Sup D", d(2026, 8, 14), "Complete", d(2026, 8, 20), "Yes")
    act(p4, "Control plan sign-off", "Demo Maint Sup D", d(2026, 10, 30), "Complete", d(2026, 10, 27), "Yes")
    act(p5, "Draft charter", "Demo Whse Sup E", None, "Not Started")
    act(p5, "Pick error baseline", "Demo Whse Sup E", d(2026, 12, 11), "Not Started", ms="Yes")

    FN = []
    FN.append(dict(Proj=p1, Type="Hard Savings", Method="Waste - Material Loss", Desc="Film scrap reduction (demo)",
                   Source="Demo scale data", Key="DEMO-FILM", Attr=0.6, BasePeriod="Jun-Jul 2026", PWeeks=4, UnitM="lbs",
                   BVol=100000, CVol=110000, BRate=0.04, CRate=0.03, UCost=1.25, Impl=2000, Start=d(2026, 11, 1),
                   Status="Approved", Reviewer="Demo Finance", RevDate=d(2026, 11, 12), Realized=1500))
    FN.append(dict(Proj=p2, Type="Hard Savings", Method="Waste - Material Loss", Desc="Shared film pool (demo)",
                   Source="Demo scale data", Key="DEMO-FILM", Attr=0.5, PWeeks=4, BVol=100000, CVol=110000,
                   BRate=0.04, CRate=0.035, UCost=1.25, Start=d(2026, 11, 1), Status="Under Review"))
    FN.append(dict(Proj=p2, Type="Productivity Capacity", Method="Changeover - Time Reduction",
                   Desc="35 min/CO not yet demonstrated (demo)", Source="Demo video study", Key="DEMO-CO-L2",
                   PWeeks=1, COMin=35, COFreq=10, Demo="No", Start=d(2026, 12, 1), Status="Not Submitted"))
    FN.append(dict(Proj=p2, Type="Hard Savings", Method="Changeover - Time Reduction",
                   Desc="Overtime eliminated (demo)", Source="Demo OT report", Key="DEMO-CO-OT", PWeeks=1, COMin=20,
                   COFreq=5, Demo="Yes", RVal=4, Impl=500, Start=d(2026, 10, 1), Status="Approved", Reviewer="Demo Finance",
                   RevDate=d(2026, 10, 20), Realized=2400))
    FN.append(dict(Proj=p3, Type="Cost Avoidance", Method="SAP - Error / Rework / Labor", Desc="Rework labor (demo)",
                   PWeeks=1, Rework=300, Labor=200, Status="Rejected", Start=d(2026, 10, 1)))
    FN.append(dict(Proj=p4, Type="Cost Avoidance", Method="Other - Data Supported", Desc="Avoided breakdown (demo)",
                   Source="Demo CMMS", Key="DEMO-PM", PWeeks=52, Down=12000, OneTime=3000, Start=d(2026, 9, 1),
                   Status="Approved", Reviewer="Demo Finance", RevDate=d(2026, 11, 2), Realized=4000))
    FN.append(dict(Proj=p4, Type="Hard Savings", Method="Other - Data Supported", Desc="Unapproved claim (demo)",
                   Source="Demo", Key="DEMO-PM2", PWeeks=4, Labor=1000, Start=d(2026, 9, 1), Status="Under Review",
                   Realized=9999))

    WU = [dict(Week=d(2026, 11, 13), Proj=p1, By="Demo Supervisor A", Done="Trial complete (demo)",
               Next="Finish SOP (demo)", KPI=3.1, Risks="None", Support="Training time on C shift"),
          dict(Week=d(2026, 11, 6), Proj=p1, By="Demo Supervisor A", Done="Trial running (demo)", KPI=3.4),
          dict(Week=d(2026, 11, 13), Proj=p2, By="Demo Supervisor B", Done="SMED split ongoing (demo)",
               Next="Checklist draft", DecReq="Approve tooling cart purchase"),
          dict(Week=d(2026, 10, 30), Proj=p2, By="Demo Supervisor B", KPI=88),
          dict(Week=d(2026, 10, 30), Proj=p4, By="Demo Maint Sup D", Done="Control plan signed", KPI=91)]
    RK = [dict(Proj=p3, Raised=d(2026, 10, 16), Type="Blocker", Desc="No SAP report access (demo)", Impact="High",
               Owner="Demo CI Coach", Support="IT access", Level="Plant Leadership", EscDate=d(2026, 10, 20),
               Target=d(2026, 11, 6), Status="Escalated"),
          dict(Proj=p1, Raised=d(2026, 11, 2), Type="Risk", Desc="C shift training coverage (demo)", Impact="Medium",
               Owner="Demo Supervisor A", Target=d(2026, 12, 1), Status="Open"),
          dict(Proj=p4, Raised=d(2026, 7, 1), Type="Issue", Desc="Parts lead time (demo)", Impact="Low",
               Owner="Demo Maint Sup D", Target=d(2026, 8, 1), Status="Closed", Closed=d(2026, 7, 20))]
    EV = [dict(Proj=p1, Action="A-0001", Type="Data Extract / Report", Title="Demo scrap log",
               URL="https://example.com/demo/scrap-log", By="Demo Supervisor A", Date=d(2026, 9, 17), Verif="Verified",
               Reviewer="Demo CI Coach"),
          dict(Proj=p1, Action="A-0002", Type="Photo - After", Title="Demo pareto", URL="not-a-link",
               By="Demo Supervisor A", Verif="Pending Review"),
          dict(Proj=p2, Action="A-0001", Type="Other", Title="Wrong project link (demo)",
               URL="https://example.com/demo/x", Verif="Pending Review")]
    CH = [dict(Proj=p1, Version="1", CStatus="Approved", Problem="Demo problem statement", SMART="Reduce scrap from 4.0% to 2.5% by 01/29/2027",
               ScopeIn="Line demo", Current="Demo current condition", LeadDate=d(2026, 8, 3)),
          dict(Proj=p2, CStatus="Draft")]
    DV = [dict(Sup="Demo Supervisor A", Proj=p1, Date=d(2026, 10, 1), Reviewer="Demo CI Coach", Participation=3,
               CharterQ=3, ProbDef=2, RCA=2, Data=3, Teamwork=4, Closure=3, Comms=3, Sustain=2, NextRev=d(2026, 11, 1)),
          dict(Sup="Demo Supervisor A", Proj=p1, Date=d(2026, 11, 10), Reviewer="Demo CI Coach", Participation=4,
               CharterQ=3, ProbDef=3, RCA=3, Data=3, Teamwork=4, Closure=3, Comms=4, NextRev=d(2027, 1, 10)),
          dict(Sup="Demo Supervisor B", Proj=p2, Date=d(2026, 10, 15), Reviewer="Demo CI Coach", Participation=2,
               CharterQ=2, ProbDef=1, Data=2, NextRev=d(2026, 11, 28))]
    return dict(people=people, projects=P, actions=A, finance=FN, woru=WU, risks=RK, evidence=EV, charters=CH, dev=DV)


# ============================================================================= CONDITIONAL FORMATS
def cf_text(ws, rng, text, fl, fc, bold=True):
    first = rng.split(":")[0]
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'EXACT({first},"{text}")'], fill=fill(fl),
                                                   font=Font(name=FONT, color=fc, bold=bold)))


def cf_formula(ws, rng, formula, fl=None, fc=None, bold=False, italic=False):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[formula], fill=fill(fl) if fl else None,
                                                   font=Font(name=FONT, color=fc, bold=bold, italic=italic) if fc else None))


def cf_rag(ws, rng):
    for t, fl, fc in [("RED", R_FILL, R_FONT), ("AMBER", A_FILL, A_FONT), ("GREEN", G_FILL, G_FONT),
                      ("PENDING", P_FILL, P_FONT), ("CLOSED", "DDEBF7", NAVY)]:
        cf_text(ws, rng, t, fl, fc)


def cf_na(ws, rng):
    first = rng.split(":")[0]
    for t in ("N/A", "Not quantified", "Not assessed"):
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'EXACT({first},"{t}")'],
                                                       font=Font(name=FONT, color="9C5700", italic=True)))


# ============================================================================= DASHBOARD
def build_dashboard(wb, ws, demo):
    widths = dict(A=2, B=12, C=30, D=17, E=14, F=9, G=11, H=13, I=10, J=11, K=11, L=11, M=9, N=14, O=14, P=36)
    for k, v in widths.items():
        ws.column_dimensions[k].width = v
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 85
    lc = "P"
    co = ("DEMONSTRATION / TRAINING COPY - FICTITIOUS DATA" if demo else "C.H. GUENTHER & SON  |  DUNCANVILLE PLANT")
    merge_write(ws, "B1:N1", co, F(14, True, WHITE), fill(R_FONT if demo else NAVY), LEFT_C)
    merge_write(ws, "B2:N2", "Supervisor Productivity & Continuous Improvement - Executive Dashboard / WOR View",
                F(16, True, NAVY), None, LEFT_C)
    ws["B3"] = '="As of: "&TEXT(AsOfDate,"dddd, mmmm d, yyyy")&"   |   Report week: "&TEXT(AsOfDate-6,"mm/dd")&" - "&TEXT(AsOfDate,"mm/dd/yyyy")&"   |   All figures are live formulas from the registers"'
    ws.merge_cells("B3:N3")
    ws["B3"].font = F(9, False, GRAY_D, True)
    merge_write(ws, "O1:P3", "[Official C.H. Guenther logo:\ninsert supplied file only -\nnot reproduced here]",
                F(8, False, GRAY_D, True), fill(GRAY_XL), CENTER, border=BORDER)
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 26
    for col in range(2, 17):
        ws.cell(3, col).border = Border(bottom=MED)
    # nav
    nav = [(S_REG, "Projects"), (S_CH, "Charter"), (S_ACT, "Actions"), (S_FIN, "Savings"), (S_WOR, "WOR"),
           (S_WORU, "Updates"), (S_DEV, "Develop"), (S_TOOLS, "Tools"), (S_RISK, "Risks"),
           (S_EVD, "Evidence"), (S_MD, "Settings"), (S_GUIDE, "Guide"), (S_DD, "Definitions"), (S_QA, "QA")]
    cols = "BCDEFGHIJKLMNO"
    for i, (s, lab) in enumerate(nav):
        c = ws[f"{cols[i]}4"]
        c.value = f'=HYPERLINK("#\'{s}\'!A1","{lab} >")'
        c.font = Font(name=FONT, size=8, color="0563C1", underline="single")
        c.alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
    ws.row_dimensions[4].height = 16

    section(ws, "B5:P5", "KEY PERFORMANCE INDICATORS   (ACTUAL = recorded data  |  FORECAST = unverified estimate  |  VERIFIED = Finance-approved  |  N/A = data missing)")
    cards = [
        ("B", "C", "ACTIVE PROJECTS", '=SUMPRODUCT((reg_ID<>"")*(reg_Status<>"Complete")*(reg_Status<>"Cancelled")*(reg_Status<>"Pending Assignment"))',
         '=COUNTIFS(reg_ID,"?*",reg_Status,"Pending Assignment")&" pending assignment  |  "&COUNTIF(reg_ID,"?*")&" registered"', "0", "ACTUAL"),
        ("D", "E", "COMPLETED PROJECTS", '=COUNTIFS(reg_ID,"?*",reg_Status,"Complete")', '="Status = Complete"', "0", "ACTUAL"),
        ("F", "G", "ACTION COMPLETION", '=IF(K_ActTotal=0,"N/A",K_ActDone/K_ActTotal)',
         '=K_ActDone&" of "&K_ActTotal&" actions credited"', FMT_PCT, "ACTUAL"),
        ("H", "I", "OVERDUE ACTIONS", '=COUNTIF(act_Flag,"OVERDUE")', '=COUNTIF(act_Flag,"DUE SOON")&" due within "&DueSoonDays&" days"', "0", "ACTUAL"),
        ("J", "K", "PROJECTED ANNUAL SAVINGS", '=SUMIFS(fin_FcstNum,fin_Type,"Hard Savings",fin_Status,"<>Rejected",fin_Proj,"?*")',
         '="FORECAST - hard savings, not Rejected"', FMT_USD, "FORECAST"),
        ("L", "M", "VERIFIED REALIZED SAVINGS", '=SUMIFS(fin_Verified,fin_Type,"Hard Savings")',
         '="Finance-Approved hard savings only"', FMT_USD, "VERIFIED"),
        ("N", "O", "PROJECTS REQUIRING ATTENTION", '=COUNTIF(reg_Attn,1)',
         '=COUNTIF(reg_RAG,"RED")&" red | "&COUNTIF(reg_RAG,"AMBER")&" amber | "&COUNTIF(reg_RAG,"PENDING")&" pending"', "0", "ACTUAL"),
        ("P", "P", "LEADERSHIP ESCALATIONS", '=COUNTIF(risk_Lead,"LEADERSHIP ACTION")',
         '=SUM(risk_IsOpen)&" open risks / blockers"', "0", "ACTUAL"),
    ]
    tagcol = {"ACTUAL": NAVY, "FORECAST": "7F7F7F", "VERIFIED": G_FONT}
    for a, b, lab, f, sub, fmt, tag in cards:
        merge_write(ws, f"{a}6:{b}6", lab, F(8, True, GRAY_D), fill(GRAY_XL), CENTER)
        c = merge_write(ws, f"{a}7:{b}8", f, F(22, True, tagcol[tag], italic=(tag == "FORECAST")), fill(WHITE), CENTER, fmt)
        merge_write(ws, f"{a}9:{b}9", sub, F(7, False, GRAY_D, True), fill(GRAY_XL), CENTER)
        for r in range(6, 10):
            for col in range(openpyxl.utils.column_index_from_string(a), openpyxl.utils.column_index_from_string(b) + 1):
                cell = ws.cell(r, col)
                cell.border = Border(left=THIN if col == openpyxl.utils.column_index_from_string(a) else None,
                                     right=THIN if col == openpyxl.utils.column_index_from_string(b) else None,
                                     top=Side(style="thick", color=tagcol[tag]) if r == 6 else None,
                                     bottom=THIN if r == 9 else None)
    ws.row_dimensions[7].height = 22
    ws.row_dimensions[8].height = 22
    # helper totals (named)
    ws["R6"] = "Helper: total actions"
    ws["S6"] = '=COUNTIFS(act_Proj,"?*",act_Status,"<>Cancelled")'
    ws["R7"] = "Helper: credited actions"
    ws["S7"] = '=COUNTIF(act_Credit,1)'
    add_name(wb, "K_ActTotal", f"{q(ws.title)}!$S$6")
    add_name(wb, "K_ActDone", f"{q(ws.title)}!$S$7")
    for r in (6, 7):
        ws[f"R{r}"].font = F(8, False, GRAY_M)
        ws[f"S{r}"].font = F(8, False, GRAY_M)

    # secondary metrics row 11-12
    sec = [
        ("B", "C", "MILESTONE COMPLETION", '=IF(COUNTIFS(act_MS,"Yes",act_Proj,"?*",act_Status,"<>Cancelled")=0,"N/A",COUNTIFS(act_MS,"Yes",act_Credit,1)/COUNTIFS(act_MS,"Yes",act_Proj,"?*",act_Status,"<>Cancelled"))', FMT_PCT),
        ("D", "E", "OUTCOME: AT/ABOVE TARGET", '=COUNTIF(reg_Attain,">=1")&" of "&COUNT(reg_Attain)&" measured"', None),
        ("F", "G", "COST AVOIDANCE (FCST)", '=SUMIFS(fin_FcstNum,fin_Type,"Cost Avoidance",fin_Status,"<>Rejected",fin_Proj,"?*")', FMT_USD),
        ("H", "I", "COST AVOIDANCE (VERIFIED)", '=SUMIFS(fin_Verified,fin_Type,"Cost Avoidance")', FMT_USD),
        ("J", "K", "CAPACITY HRS/YR (NON-CASH)", '=SUMIFS(fin_CapHrsYr,fin_Status,"<>Rejected")', "#,##0.0"),
        ("L", "M", "MISSING EVIDENCE", '=COUNTIF(act_EvCheck,"MISSING EVIDENCE")', "0"),
        ("N", "O", "SAVINGS LINES W/ MISSING DATA", '=COUNTIFS(fin_Proj,"?*",fin_Missing,"?*")', "0"),
        ("P", "P", "BLOCKED PROJECTS", '=COUNTIFS(reg_ID,"?*",reg_Status,"Blocked")', "0"),
    ]
    for a, b, lab, f, fmt in sec:
        merge_write(ws, f"{a}11:{b}11", lab, F(7, True, GRAY_D), fill(GRAY_XL), CENTER, border=BORDER)
        merge_write(ws, f"{a}12:{b}12", f, F(13, True, NAVY), fill(WHITE), CENTER, fmt, border=BORDER)
    ws.row_dimensions[12].height = 22
    merge_write(ws, "B13:P13",
                "Data integrity rules: action completion = credited complete actions / non-cancelled actions (status-based, never an average of "
                "subjective %). Forecast savings are unverified until Finance Status = Approved. Capacity hours are not cash. N/A = required data "
                "not yet entered (never assumed zero).", F(7, False, GRAY_D, True), None, LEFT_C)
    ws.row_dimensions[13].height = 22

    # project table
    section(ws, "B15:P15", "PROJECT STATUS - WOR MEETING VIEW  (first 15 projects in the register)")
    hdr = ["Project ID", "Project Title", "Leader (Primary)", "Department", "Shift", "Stage", "Status", "RAG",
           "Action Completion", "Milestone Completion", "Outcome Attainment", "Overdue Actions",
           "Forecast Hard $/yr", "Verified Hard $", "Next Milestone / Attention Reasons"]
    for j, h in enumerate(hdr):
        c = ws.cell(16, 2 + j, h)
        c.font = F(8, True, WHITE)
        c.fill = fill(NAVY)
        c.alignment = CENTER
        c.border = BORDER
    ws.row_dimensions[16].height = 30
    srcmap = ["ID", "Title", "Leader", "Dept", "Shift", "Stage", "Status", "RAG", "ActPct", "MSPct", "Attain", "Overdue",
              "ExpHard", "VerHard", None]
    reg_letters = {c["key"]: get_column_letter(i + 1) for i, c in enumerate(register_cols())}
    fmts = [None, None, None, None, None, None, None, None, FMT_PCT, FMT_PCT, FMT_PCT, "0", FMT_USD, FMT_USD, None]
    for k in range(15):
        r = 17 + k
        rr = FIRST + k
        idc = f"{q(S_REG)}!${reg_letters['ID']}{rr}"
        for j, key in enumerate(srcmap):
            c = ws.cell(r, 2 + j)
            if key is None:
                c.value = (f'=IF({idc}="","",IF(OR({q(S_REG)}!{reg_letters["RAG"]}{rr}="RED",{q(S_REG)}!{reg_letters["RAG"]}{rr}="AMBER",'
                           f'{q(S_REG)}!{reg_letters["RAG"]}{rr}="PENDING"),{q(S_REG)}!{reg_letters["Reason"]}{rr},{q(S_REG)}!{reg_letters["NextMS"]}{rr}))')
            else:
                src = f"{q(S_REG)}!{reg_letters[key]}{rr}"
                c.value = f'=IF({idc}="","",IF({src}&""="","-",{src}))'
            c.font = F(8, key in ("ID", "RAG"))
            c.border = BORDER
            c.alignment = Alignment(vertical="center", wrap_text=True, horizontal="center" if j >= 4 and j <= 13 else "left")
            if fmts[j]:
                c.number_format = fmts[j]
        ws.row_dimensions[r].height = 24
    cf_rag(ws, "I17:I31")
    cf_na(ws, "J17:O31")
    ws.conditional_formatting.add("J17:K31", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                                                         color="8EA9DB", showValue=True))
    cf_formula(ws, "M17:M31", 'AND(ISNUMBER(M17),M17>0)', R_FILL, R_FONT, True)
    cf_formula(ws, "B17:P31", 'MOD(ROW(),2)=0', GRAY_XL)

    # by department
    section(ws, "B33:P33", "PROGRESS BY DEPARTMENT")
    dh = ["Department", "", "Projects", "Active", "Complete", "Pending", "Actions", "Actions Done", "Action Completion",
          "Overdue", "Attention", "Forecast Hard $/yr", "Verified Hard $", "Cost Avoid. Fcst $/yr", "Notes"]
    for j, h in enumerate(dh):
        c = ws.cell(34, 2 + j, h)
        c.font = F(8, True, WHITE)
        c.fill = fill(NAVY)
        c.alignment = CENTER
        c.border = BORDER
    ws.merge_cells("B34:C34")
    ws.row_dimensions[34].height = 30

    def group_rows(start, n, keyexpr_list, crit_col, label_kind):
        rows = []
        for k in range(n):
            r = start + k
            rows.append(r)
        return rows

    def write_group(r, label_formula, crit, critrange, blank=False, other=False, total_rows=None):
        """crit: cell ref holding the criterion label; blank -> unset values."""
        if other:
            return
        ws.cell(r, 2, label_formula)
        ws.merge_cells(f"B{r}:C{r}")
        if blank:
            cc = f'{critrange},""'
        else:
            cc = f'{critrange},$B{r}'
        fs = {
            4: f'=IF($B{r}="","",COUNTIFS(reg_ID,"?*",{cc}))',
            5: f'=IF($B{r}="","",SUMPRODUCT((reg_ID<>"")*({critrange}={"" if blank else "$B" + str(r)}){"" if not blank else ""}*(reg_Status<>"Complete")*(reg_Status<>"Cancelled")*(reg_Status<>"Pending Assignment")))',
            6: f'=IF($B{r}="","",COUNTIFS(reg_ID,"?*",{cc},reg_Status,"Complete"))',
            7: f'=IF($B{r}="","",COUNTIFS(reg_ID,"?*",{cc},reg_Status,"Pending Assignment"))',
            8: f'=IF($B{r}="","",SUMIFS(reg_NAct,reg_ID,"?*",{cc}))',
            9: f'=IF($B{r}="","",SUMIFS(reg_NDone,reg_ID,"?*",{cc}))',
            10: f'=IF($B{r}="","",IF(H{r}=0,"N/A",I{r}/H{r}))',
            11: f'=IF($B{r}="","",SUMIFS(reg_Overdue,reg_ID,"?*",{cc}))',
            12: f'=IF($B{r}="","",SUMIFS(reg_Attn,reg_ID,"?*",{cc}))',
            13: f'=IF($B{r}="","",SUMIFS(reg_ExpHard,reg_ID,"?*",{cc}))',
            14: f'=IF($B{r}="","",SUMIFS(reg_VerHard,reg_ID,"?*",{cc}))',
            15: f'=IF($B{r}="","",SUMIFS(reg_ExpAvoid,reg_ID,"?*",{cc}))',
        }
        if blank:
            fs[5] = f'=SUMPRODUCT((reg_ID<>"")*({critrange}="")*(reg_Status<>"Complete")*(reg_Status<>"Cancelled")*(reg_Status<>"Pending Assignment"))'
        for col, f in fs.items():
            c = ws.cell(r, col, f)
        ws.cell(r, 16, '=IF(AND($B{0}<>"",D{0}>0),IF(L{0}>0,L{0}&" project(s) need attention",""),"")'.format(r))
        for col in range(2, 17):
            c = ws.cell(r, col)
            c.font = F(8, col == 2)
            c.border = BORDER
            c.alignment = Alignment(horizontal="left" if col in (2, 16) else "center", vertical="center")
        ws.cell(r, 10).number_format = FMT_PCT
        for col in (13, 14, 15):
            ws.cell(r, col).number_format = FMT_USD

    # 10 dept rows + Not set + Other + total
    for k in range(10):
        r = 35 + k
        write_group(r, f'=IF(INDEX(md_Dept,{k + 1})="","",INDEX(md_Dept,{k + 1}))', None, "reg_Dept")
    write_group(45, "(Department not set)", None, "reg_Dept", blank=True)
    ws.cell(45, 2).font = F(8, True, A_FONT, True)
    # other row: total minus listed
    r = 46
    ws.cell(r, 2, "(Other departments beyond first 10)")
    ws.merge_cells(f"B{r}:C{r}")
    ws.cell(r, 4, '=COUNTIF(reg_ID,"?*")-SUM(D35:D45)')
    ws.cell(r, 16, '=IF(D46>0,"Increase dashboard department rows","")')
    for col in range(2, 17):
        ws.cell(r, col).font = F(8, False, GRAY_D, True)
        ws.cell(r, col).border = BORDER
        ws.cell(r, col).alignment = Alignment(horizontal="center", vertical="center")
    r = 47
    ws.cell(r, 2, "TOTAL")
    ws.merge_cells(f"B{r}:C{r}")
    for col in range(4, 16):
        L = get_column_letter(col)
        if col == 10:
            ws.cell(r, col, f'=IF(H{r}=0,"N/A",I{r}/H{r})').number_format = FMT_PCT
        else:
            ws.cell(r, col, f"=SUM({L}35:{L}46)")
    for col in range(2, 17):
        c = ws.cell(r, col)
        c.font = F(8, True, WHITE)
        c.fill = fill(STEEL)
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center")
    for col in (13, 14, 15):
        ws.cell(r, col).number_format = FMT_USD
    ws.conditional_formatting.add("J35:J47", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                                                         color="8EA9DB", showValue=True))
    cf_na(ws, "J35:J47")

    # by shift
    section(ws, "B49:P49", "PROGRESS BY SHIFT")
    for j, h in enumerate(dh):
        c = ws.cell(50, 2 + j, "Shift" if j == 0 else h)
        c.font = F(8, True, WHITE)
        c.fill = fill(NAVY)
        c.alignment = CENTER
        c.border = BORDER
    ws.merge_cells("B50:C50")
    ws.row_dimensions[50].height = 30
    for k in range(5):
        write_group(51 + k, f'=IF(INDEX(md_Shift,{k + 1})="","",INDEX(md_Shift,{k + 1}))', None, "reg_Shift")
    write_group(56, "(Shift not set)", None, "reg_Shift", blank=True)
    ws.cell(56, 2).font = F(8, True, A_FONT, True)
    r = 57
    ws.cell(r, 2, "TOTAL")
    ws.merge_cells(f"B{r}:C{r}")
    for col in range(4, 16):
        L = get_column_letter(col)
        if col == 10:
            ws.cell(r, col, f'=IF(H{r}=0,"N/A",I{r}/H{r})').number_format = FMT_PCT
        else:
            ws.cell(r, col, f"=SUM({L}51:{L}56)")
    for col in range(2, 17):
        c = ws.cell(r, col)
        c.font = F(8, True, WHITE)
        c.fill = fill(STEEL)
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center")
    for col in (13, 14, 15):
        ws.cell(r, col).number_format = FMT_USD
    ws.conditional_formatting.add("J51:J57", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                                                         color="8EA9DB", showValue=True))
    cf_na(ws, "J51:J57")
    ws.print_area = "B1:P57"

    # ---------------- finance pipeline + reconciliation (rows 59+)
    section(ws, "B59:G59", "FINANCE APPROVAL PIPELINE (all benefit types)")
    ph = ["Finance Status", "", "Lines", "Attributed Forecast $/yr", "Verified $", "Capacity hrs/yr"]
    for j, h in enumerate(ph):
        c = ws.cell(60, 2 + j, h)
        c.font = F(8, True, WHITE)
        c.fill = fill(NAVY)
        c.alignment = CENTER
        c.border = BORDER
    ws.merge_cells("B60:C60")
    for k, st in enumerate(["Not Submitted", "Under Review", "Approved", "Rejected", "(Status not set)"]):
        r = 61 + k
        ws.cell(r, 2, st)
        ws.merge_cells(f"B{r}:C{r}")
        crit = '""' if st.startswith("(") else f"$B{r}"
        ws.cell(r, 4, f'=COUNTIFS(fin_Proj,"?*",fin_Status,{crit})')
        ws.cell(r, 5, f'=SUMIFS(fin_FcstNum,fin_Proj,"?*",fin_Status,{crit})').number_format = FMT_USD
        ws.cell(r, 6, f'=SUMIFS(fin_Verified,fin_Proj,"?*",fin_Status,{crit})').number_format = FMT_USD
        ws.cell(r, 7, f'=SUMIFS(fin_CapHrsYr,fin_Proj,"?*",fin_Status,{crit})').number_format = "#,##0.0"
        for col in range(2, 8):
            ws.cell(r, col).font = F(8)
            ws.cell(r, col).border = BORDER
            ws.cell(r, col).alignment = Alignment(horizontal="center")
    ws.cell(66, 2, "TOTAL")
    ws.merge_cells("B66:C66")
    for col, fmt in zip((4, 5, 6, 7), ("0", FMT_USD, FMT_USD, "#,##0.0")):
        L = get_column_letter(col)
        ws.cell(66, col, f"=SUM({L}61:{L}65)").number_format = fmt
    for col in range(2, 8):
        c = ws.cell(66, col)
        c.font = F(8, True, WHITE)
        c.fill = fill(STEEL)
        c.alignment = Alignment(horizontal="center")
    merge_write(ws, "B67:G68", "Forecast for Rejected lines is shown for transparency only and is excluded from every KPI. Verified $ "
                "can only be non-zero on Approved lines.", F(7, False, GRAY_D, True), None, WRAP_TOP)

    section(ws, "I59:P59", "DATA INTEGRITY & RECONCILIATION CHECKS (live)")
    rh = ["Check", "", "", "Dashboard / Source", "Underlying Table", "Result", "", ""]
    for j, h in enumerate(rh):
        c = ws.cell(60, 9 + j, h)
        c.font = F(8, True, WHITE)
        c.fill = fill(NAVY)
        c.alignment = CENTER
        c.border = BORDER
    ws.merge_cells("I60:K60")
    ws.merge_cells("N60:P60")
    checks = [
        ("Projects by department = registered projects", "=D47", '=COUNTIF(reg_ID,"?*")'),
        ("Projects by shift = registered projects", "=D57", '=COUNTIF(reg_ID,"?*")'),
        ("Actions in tracker = actions linked to projects", "=K_ActTotal", "=SUM(reg_NAct)"),
        ("Credited actions = sum of project completions", "=K_ActDone", "=SUM(reg_NDone)"),
        ("Forecast hard savings = sum of project forecasts", "=J7", "=SUM(reg_ExpHard)"),
        ("Verified hard savings = sum of project verified", "=L7", "=SUM(reg_VerHard)"),
        ("Verified $ on non-Approved lines (must be 0)", '=SUMIFS(fin_Verified,fin_Status,"<>Approved")', "=0"),
        ("Over-attributed benefit pools (must be 0)", '=COUNTIF(fin_Dbl,"OVER*")', "=0"),
        ("Trend: completed (latest week) = credited actions", "=D84", "=K_ActDone"),
        ("Trend: overdue (latest week) = overdue actions", "=E84", "=H7"),
        ("Duplicate WOR weekly updates (must be 0)", '=COUNTIF(wor_Dup,"DUPLICATE week")', "=0"),
        ("Records with unknown Project ID (must be 0)", '=COUNTIF(act_PTitle,"Unknown Project ID")+COUNTIF(fin_PTitle,"Unknown Project ID")+COUNTIF(wor_PTitle,"Unknown Project ID")+COUNTIF(risk_PTitle,"Unknown Project ID")+COUNTIF(evd_ActChk,"Unknown Action ID")', "=0"),
    ]
    for k, (lab, a, b) in enumerate(checks):
        r = 61 + k
        ws.cell(r, 9, lab)
        ws.merge_cells(f"I{r}:K{r}")
        ws.cell(r, 12, a)
        ws.cell(r, 13, b)
        ws.cell(r, 14, f'=IF(ROUND(L{r}-M{r},2)=0,"PASS","CHECK - review source data")')
        ws.merge_cells(f"N{r}:P{r}")
        for col in range(9, 17):
            c = ws.cell(r, col)
            c.font = F(8, col == 14)
            c.border = BORDER
            c.alignment = Alignment(horizontal="left" if col == 9 else "center", vertical="center")
        ws.cell(r, 12).number_format = "#,##0.##"
        ws.cell(r, 13).number_format = "#,##0.##"
    cf_text(ws, "N61:N72", "PASS", G_FILL, G_FONT)
    cf_formula(ws, "N61:N72", 'LEFT(N61,5)="CHECK"', A_FILL, A_FONT, True)
    add_name(wb, "K_ChecksPassed", f"{q(ws.title)}!$N$61:$N$72")

    # ---------------- trends
    section(ws, "B74:P74", "TRENDS  (computed from recorded dates - weekly windows end on the As-Of date; monthly windows are calendar months)")
    th = ["Week Ending", "Planned Actions Due (cum.)", "Actions Completed (cum.)", "Overdue at Week End", "Completion vs. Due (cum.)"]
    hdr_r = 75
    for j, h in enumerate(th):
        c = ws.cell(hdr_r, 2 + j, h)
        c.font = F(8, True, WHITE)
        c.fill = fill(NAVY)
        c.alignment = CENTER
        c.border = BORDER
    ws.row_dimensions[hdr_r].height = 30
    for k in range(12):
        r = hdr_r + 1 + k  # 76..87  -> latest at 87
        ws.cell(r, 2, f"=AsOfDate-7*{11 - k}").number_format = "mm/dd/yy"
        ws.cell(r, 3, f'=COUNTIFS(act_Due,"<="&B{r},act_Proj,"?*",act_Status,"<>Cancelled")')
        ws.cell(r, 4, f'=COUNTIFS(act_CompDate,"<="&B{r},act_Credit,1)')
        ws.cell(r, 5, f'=COUNTIFS(act_Due,"<"&B{r},act_Proj,"?*",act_Status,"<>Cancelled")-COUNTIFS(act_Due,"<"&B{r},act_CompDate,"<="&B{r},act_Credit,1)')
        ws.cell(r, 6, f'=IF(C{r}=0,"N/A",COUNTIFS(act_Due,"<="&B{r},act_CompDate,"<="&B{r},act_Credit,1)/C{r})').number_format = FMT_PCT
        for col in range(2, 7):
            ws.cell(r, col).font = F(8)
            ws.cell(r, col).border = BORDER
            ws.cell(r, col).alignment = Alignment(horizontal="center")
    # fix reconciliation references to latest week row 87
    ws["L69"] = "=D87"
    ws["L70"] = "=E87"

    mh = ["Month End", "Forecast Hard Run-Rate (cum.)", "Verified Hard Savings (cum., by Finance review date)", "Verified Cost Avoidance (cum.)"]
    for j, h in enumerate(mh):
        c = ws.cell(hdr_r, 8 + j, h)
        c.font = F(8, True, WHITE)
        c.fill = fill(NAVY)
        c.alignment = CENTER
        c.border = BORDER
    for k in range(12):
        r = hdr_r + 1 + k
        ws.cell(r, 8, f"=EOMONTH(AsOfDate,{k - 11})").number_format = "mmm yy"
        ws.cell(r, 9, f"=SUMPRODUCT(fin_MF,((YEAR(H{r})*12+MONTH(H{r}))-fin_StartIdx+1)*((YEAR(H{r})*12+MONTH(H{r}))>=fin_StartIdx))").number_format = FMT_USD
        ws.cell(r, 10, f'=SUMIFS(fin_Verified,fin_Type,"Hard Savings",fin_RevDate,"<="&H{r})').number_format = FMT_USD
        ws.cell(r, 11, f'=SUMIFS(fin_Verified,fin_Type,"Cost Avoidance",fin_RevDate,"<="&H{r})').number_format = FMT_USD
        for col in range(8, 12):
            ws.cell(r, col).font = F(8)
            ws.cell(r, col).border = BORDER
            ws.cell(r, col).alignment = Alignment(horizontal="center")
    merge_write(ws, f"M{hdr_r}:P{hdr_r + 3}",
                "How trends are calculated: Planned = actions with due date on/before week end. Completed = credited actions with "
                "completed date on/before week end. Overdue = due before week end and not completed by then. Forecast run-rate = "
                "each Hard Savings line's attributed annualized recurring benefit / 12, accumulated from its benefit start month "
                "(FORECAST, unverified). Verified = Finance-approved realized savings dated by Finance review date.",
                F(7, False, GRAY_D, True), fill(GRAY_XL), WRAP_TOP, border=BORDER)
    cf_na(ws, "F76:F87")

    # charts
    def style_chart(ch, title, ytitle):
        ch.title = title
        ch.style = 2
        ch.height = 7.2
        ch.width = 13.5
        ch.y_axis.title = ytitle
        ch.legend.position = "b"
        ch.roundedCorners = False
        ch.legend.overlay = False
        ch.x_axis.delete = False
        ch.y_axis.delete = False
        ch.y_axis.majorGridlines = None

    lc1 = LineChart()
    style_chart(lc1, "Action Completion Trend (12 weeks)", "Actions (cumulative)")
    lc1.add_data(Reference(ws, min_col=3, max_col=4, min_row=hdr_r, max_row=hdr_r + 12), titles_from_data=True)
    lc1.set_categories(Reference(ws, min_col=2, min_row=hdr_r + 1, max_row=hdr_r + 12))
    lc1.x_axis.number_format = "mm/dd"
    s0, s1 = lc1.series
    s0.graphicalProperties.line.solidFill = "A6A6A6"
    s0.graphicalProperties.line.dashStyle = "dash"
    s1.graphicalProperties.line.solidFill = NAVY
    s1.graphicalProperties.line.width = 28000
    ws.add_chart(lc1, "B90")

    bc = BarChart()
    style_chart(bc, "Overdue Actions at Week End", "Actions")
    bc.add_data(Reference(ws, min_col=5, min_row=hdr_r, max_row=hdr_r + 12), titles_from_data=True)
    bc.set_categories(Reference(ws, min_col=2, min_row=hdr_r + 1, max_row=hdr_r + 12))
    bc.series[0].graphicalProperties.solidFill = "C0504D"
    bc.series[0].graphicalProperties.line.noFill = True
    bc.gapWidth = 60
    bc.x_axis.number_format = "mm/dd"
    ws.add_chart(bc, "G90")

    lc2 = LineChart()
    style_chart(lc2, "Savings: Forecast vs. Finance-Verified (cumulative)", "$")
    lc2.add_data(Reference(ws, min_col=9, max_col=10, min_row=hdr_r, max_row=hdr_r + 12), titles_from_data=True)
    lc2.set_categories(Reference(ws, min_col=8, min_row=hdr_r + 1, max_row=hdr_r + 12))
    lc2.x_axis.number_format = "mmm yy"
    lc2.y_axis.number_format = "$#,##0"
    a0, a1 = lc2.series
    a0.graphicalProperties.line.solidFill = "A6A6A6"
    a0.graphicalProperties.line.dashStyle = "dash"
    a1.graphicalProperties.line.solidFill = G_FONT
    a1.graphicalProperties.line.width = 28000
    ws.add_chart(lc2, "L90")

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.3
    ws.page_margins.top = ws.page_margins.bottom = 0.4
    ws.oddFooter.left.text = "C.H. Guenther & Son - Duncanville | CI Management System"
    ws.oddFooter.right.text = "Printed &D | Page &P of &N"
    ws.freeze_panes = "A5"


# ============================================================================= CONTENT SHEETS
def text_block(ws, r, label, formula, height=42, c1="B", c2="G"):
    lab = ws[f"{c1}{r}"]
    merge_write(ws, f"{c1}{r}:{c2}{r}", label, F(9, True, NAVY), fill(GRAY_XL), LEFT_C, border=BORDER)
    merge_write(ws, f"{c1}{r + 1}:{c2}{r + 1}", formula, F(9), None, WRAP_TOP, border=BORDER)
    ws.row_dimensions[r + 1].height = height
    return r + 2


def build_charter(wb, ws, demo):
    for k, v in dict(A=2, B=20, C=22, D=18, E=22, F=18, G=24, H=2, I=10, J=10).items():
        ws.column_dimensions[k].width = v
    ws.sheet_view.showGridLines = False
    banner(ws, "03  STANDARD PROJECT CHARTER",
           "Select a Project ID. Identification, team, KPI and milestone fields pull from the registers; narrative "
           "sections come from '03A Charter Records'. Print: one page, portrait.", 7, demo)
    ws["B5"] = "Select Project ID"
    ws["B5"].font = F(10, True, NAVY)
    ws["C5"] = "DEMO-001" if demo else "DUN-CI-001"
    style_input(ws["C5"])
    ws["C5"].font = F(11, True, INPUT_FONT)
    dv_list(ws, "lst_ProjectID", "C5")
    add_name(wb, "CH_Sel", f"{q(ws.title)}!$C$5")
    ws["I5"] = "=IFERROR(MATCH(CH_Sel,reg_ID,0),\"\")"
    ws["I6"] = "=IFERROR(MATCH(CH_Sel,chr_Proj,0),\"\")"
    for c in ("I5", "I6"):
        ws[c].font = F(8, False, GRAY_M)
    ws["J5"] = "helper: register row"
    ws["J6"] = "helper: charter row"
    ws["J5"].font = ws["J6"].font = F(7, False, GRAY_M)
    add_name(wb, "CH_Idx", f"{q(ws.title)}!$I$5")
    add_name(wb, "CH_CIdx", f"{q(ws.title)}!$I$6")

    def R(key, blank="[Not set]"):
        return f'=IF(CH_Idx="","[Select a valid Project ID]",IF(INDEX(reg_{key},CH_Idx)&""="","{blank}",INDEX(reg_{key},CH_Idx)))'

    def C(key):
        return (f'=IF(CH_CIdx="","[No charter record - add one in 03A Charter Records]",IF(INDEX(chr_{key},CH_CIdx)&""="",'
                f'"[Not yet defined - complete in 03A Charter Records]",INDEX(chr_{key},CH_CIdx)))')

    ws["D5"] = "Charter status"
    ws["D5"].font = F(9, True, NAVY)
    ws["E5"] = '=IF(CH_CIdx="","No record",IF(INDEX(chr_CStatus,CH_CIdx)="","Draft",INDEX(chr_CStatus,CH_CIdx)))'
    ws["F5"] = "Completeness"
    ws["F5"].font = F(9, True, NAVY)
    ws["G5"] = '=IF(CH_CIdx="","N/A",INDEX(chr_Complete,CH_CIdx))'
    ws["G5"].number_format = FMT_PCT
    for c in ("E5", "G5"):
        ws[c].font = F(10, True)
        ws[c].border = BORDER
    ws.row_dimensions[5].height = 22

    section(ws, "B7:G7", "1. PROJECT IDENTIFICATION")
    ident = [[("Project title", R("Title")), ("Project ID", '=IF(CH_Idx="","[Invalid]",CH_Sel)'), ("Category", R("Category"))],
             [("Department", R("Dept")), ("Area / line", R("Area")), ("Equipment", R("Equipment"))],
             [("Shift", R("Shift")), ("Priority", R("Priority")), ("Stage", R("Stage"))],
             [("Status", R("Status")), ("Start date", R("Start")), ("Target date", R("TargetDate"))]]
    r = 8
    for row in ident:
        for j, (lab, f) in enumerate(row):
            lc = get_column_letter(2 + 2 * j)
            vc = get_column_letter(3 + 2 * j)
            ws[f"{lc}{r}"] = lab
            ws[f"{lc}{r}"].font = F(8, True, GRAY_D)
            ws[f"{lc}{r}"].fill = fill(GRAY_XL)
            ws[f"{lc}{r}"].border = BORDER
            ws[f"{lc}{r}"].alignment = LEFT_C
            ws[f"{vc}{r}"] = f
            ws[f"{vc}{r}"].font = F(9, True)
            ws[f"{vc}{r}"].border = BORDER
            ws[f"{vc}{r}"].alignment = LEFT_C
            ws[f"{vc}{r}"].number_format = FMT_DATE
        ws.row_dimensions[r].height = 28
        r += 1
    r = text_block(ws, r, "Project description (register)", R("Desc"), 30)

    section(ws, f"B{r}:G{r}", "2. TEAM AND ROLES  (one primary owner is accountable; others support)")
    r += 1
    team = [("Project Leader - PRIMARY OWNER (accountable)", "Leader"), ("Co-Owner", "CoOwner"),
            ("Additional team members", "Team"), ("Sponsor", "Sponsor"), ("CI Coach (facilitates, provides tools)", "Coach"),
            ("Finance Reviewer (validates benefits)", "FinRev")]
    for i in range(0, len(team), 2):
        for j in range(2):
            lab, key = team[i + j]
            lc = "B" if j == 0 else "E"
            vc1 = "C" if j == 0 else "F"
            vc2 = "D" if j == 0 else "G"
            ws[f"{lc}{r}"] = lab
            ws[f"{lc}{r}"].font = F(8, True, GRAY_D)
            ws[f"{lc}{r}"].fill = fill(GRAY_XL)
            ws[f"{lc}{r}"].alignment = LEFT_C
            ws[f"{lc}{r}"].border = BORDER
            merge_write(ws, f"{vc1}{r}:{vc2}{r}", R(key, "[Not assigned]"), F(9, True), None, LEFT_C, border=BORDER)
        ws.row_dimensions[r].height = 28
        r += 1
    r = text_block(ws, r, "Team responsibilities", C("Resp"), 36)

    section(ws, f"B{r}:G{r}", "3. PROBLEM, SCOPE AND OBJECTIVE")
    r += 1
    r = text_block(ws, r, "Business problem", C("Problem"), 40)
    ws.merge_cells(f"B{r}:D{r}")
    ws[f"B{r}"] = "Scope - IN"
    ws.merge_cells(f"E{r}:G{r}")
    ws[f"E{r}"] = "Scope - OUT"
    for c in (f"B{r}", f"E{r}"):
        ws[c].font = F(9, True, NAVY)
        ws[c].fill = fill(GRAY_XL)
    merge_write(ws, f"B{r + 1}:D{r + 1}", C("ScopeIn"), F(9), None, WRAP_TOP, border=BORDER)
    merge_write(ws, f"E{r + 1}:G{r + 1}", C("ScopeOut"), F(9), None, WRAP_TOP, border=BORDER)
    ws.row_dimensions[r + 1].height = 36
    r += 2
    r = text_block(ws, r, "Current condition", C("Current"), 40)
    r = text_block(ws, r, "SMART objective (Specific, Measurable, Achievable, Relevant, Time-bound)", C("SMART"), 32)

    section(ws, f"B{r}:G{r}", "4. BASELINE AND TARGET  (ACTUAL from weekly updates; N/A = not yet measured)")
    r += 1
    hd = ["Primary KPI", "Unit", "Baseline", "Target", "Latest Actual", "Outcome Attainment"]
    keys = ["KPI", "Unit", "Baseline", "Target", "Actual", "Attain"]
    for j, h in enumerate(hd):
        c = ws.cell(r, 2 + j, h)
        c.font = F(8, True, WHITE)
        c.fill = fill(NAVY)
        c.alignment = CENTER
        c.border = BORDER
        v = ws.cell(r + 1, 2 + j, R(keys[j], "N/A"))
        v.font = F(10, True)
        v.alignment = CENTER
        v.border = BORDER
        v.number_format = FMT_PCT if keys[j] == "Attain" else FMT_NUM
    ws.row_dimensions[r + 1].height = 30
    cf_na(ws, f"B{r + 1}:G{r + 1}")
    r += 2
    r = text_block(ws, r, "Key stakeholders", C("Stake"), 30)

    section(ws, f"B{r}:G{r}", "5. EXPECTED BENEFITS  (FORECAST until Finance approves)")
    r += 1
    r = text_block(ws, r, "Expected benefits (narrative)", C("Benefits"), 30)
    bh = [("Forecast hard savings $/yr", "ExpHard"), ("Forecast cost avoidance $/yr", "ExpAvoid"),
          ("Verified hard savings $", "VerHard")]
    for j, (lab, key) in enumerate(bh):
        lc = get_column_letter(2 + 2 * j)
        vc = get_column_letter(3 + 2 * j)
        ws[f"{lc}{r}"] = lab
        ws[f"{lc}{r}"].font = F(8, True, GRAY_D)
        ws[f"{lc}{r}"].fill = fill(GRAY_XL)
        ws[f"{lc}{r}"].alignment = LEFT_C
        ws[f"{vc}{r}"] = R(key, "Not quantified")
        ws[f"{vc}{r}"].number_format = FMT_USD
        ws[f"{vc}{r}"].font = F(10, True, G_FONT if key == "VerHard" else GRAY_D, italic=(key != "VerHard"))
        ws[f"{vc}{r}"].border = BORDER
    ws.row_dimensions[r].height = 28
    r += 1

    section(ws, f"B{r}:G{r}", "6. CONSTRAINTS AND RISKS")
    r += 1
    ws.merge_cells(f"B{r}:D{r}")
    ws[f"B{r}"] = "Constraints"
    ws.merge_cells(f"E{r}:G{r}")
    ws[f"E{r}"] = '="Key risks   (open risk / escalation items in 09: "&IF(CH_Idx="",0,INDEX(reg_Risks,CH_Idx))&")"'
    for c in (f"B{r}", f"E{r}"):
        ws[c].font = F(9, True, NAVY)
        ws[c].fill = fill(GRAY_XL)
    merge_write(ws, f"B{r + 1}:D{r + 1}", C("Constraints"), F(9), None, WRAP_TOP, border=BORDER)
    merge_write(ws, f"E{r + 1}:G{r + 1}", C("CRisks"), F(9), None, WRAP_TOP, border=BORDER)
    ws.row_dimensions[r + 1].height = 36
    r += 2

    section(ws, f"B{r}:G{r}", "7. MILESTONES  (actions flagged Milestone = Yes in 04 Action Tracker)")
    r += 1
    mh = ["#", "Milestone", "", "Owner", "Due Date", "Schedule Status"]
    for j, h in enumerate(mh):
        c = ws.cell(r, 2 + j, h)
        c.font = F(8, True, WHITE)
        c.fill = fill(NAVY)
        c.alignment = CENTER
        c.border = BORDER
    ws.merge_cells(f"C{r}:D{r}")
    r += 1
    ms_first = r
    for k in range(1, 9):
        m = f'MATCH(CH_Sel&"|"&{k},act_ChKey,0)'
        ws.cell(r, 2, k).alignment = CENTER
        if k == 1:
            ws.cell(r, 3, f'=IFERROR(INDEX(act_Desc,{m}),"No milestones flagged yet - set Milestone = Yes on key actions in 04")')
        else:
            ws.cell(r, 3, f'=IFERROR(INDEX(act_Desc,{m}),"")')
        ws.merge_cells(f"C{r}:D{r}")
        ws.cell(r, 5, f'=IFERROR(INDEX(act_Owner,{m})&"","")')
        ws.cell(r, 6, f'=IFERROR(IF(INDEX(act_Due,{m})="","No due date",INDEX(act_Due,{m})),"")').number_format = FMT_DATE
        ws.cell(r, 7, f'=IFERROR(INDEX(act_Flag,{m}),"")')
        for col in range(2, 8):
            ws.cell(r, col).font = F(8)
            ws.cell(r, col).border = BORDER
            ws.cell(r, col).alignment = Alignment(vertical="center", wrap_text=True)
        r += 1
    cf_text(ws, f"G{ms_first}:G{r - 1}", "OVERDUE", R_FILL, R_FONT)
    cf_formula(ws, f"G{ms_first}:G{r - 1}", f'LEFT(G{ms_first},6)="Closed"', G_FILL, G_FONT)

    section(ws, f"B{r}:G{r}", "8. APPROVALS")
    r += 1
    ah = ["Role", "Name", "", "Date (03A)", "Signature", ""]
    for j, h in enumerate(ah):
        c = ws.cell(r, 2 + j, h)
        c.font = F(8, True, WHITE)
        c.fill = fill(NAVY)
        c.alignment = CENTER
        c.border = BORDER
    ws.merge_cells(f"C{r}:D{r}")
    ws.merge_cells(f"F{r}:G{r}")
    r += 1
    appr = [("Project Leader", "Leader", "LeadDate"), ("Sponsor", "Sponsor", "SponsorDate"),
            ("CI Coach", "Coach", "CoachDate"), ("Finance Reviewer", "FinRev", "FinDate")]
    for lab, key, dkey in appr:
        ws.cell(r, 2, lab).font = F(9, True, NAVY)
        ws.cell(r, 3, R(key, "[Not assigned]"))
        ws.merge_cells(f"C{r}:D{r}")
        ws.cell(r, 5, f'=IF(CH_CIdx="","",IF(INDEX(chr_{dkey},CH_CIdx)="","Pending",INDEX(chr_{dkey},CH_CIdx)))').number_format = FMT_DATE
        ws.merge_cells(f"F{r}:G{r}")
        for col in range(2, 8):
            ws.cell(r, col).border = BORDER
            ws.cell(r, col).alignment = LEFT_C
        ws.row_dimensions[r].height = 26
        r += 1
    ws[f"B{r}"] = ("Approval dates are entered in 03A Charter Records. Electronic dates in Excel are a record of review, not a "
                   "secure electronic signature - print and sign, or record approval in the controlled document system if required.")
    ws.merge_cells(f"B{r}:G{r + 1}")
    ws[f"B{r}"].font = F(7, False, GRAY_D, True)
    ws[f"B{r}"].alignment = WRAP_TOP
    last = r + 1
    cf_formula(ws, f"B8:G{last}", 'LEFT(B8,1)="["', None, "9C5700", italic=True)
    ws.print_area = f"A1:G{last}"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.oddFooter.left.text = "C.H. Guenther & Son - Duncanville | Standard Project Charter"
    ws.oddFooter.right.text = "Printed &D"


def build_wor(wb, ws, demo, dash_title):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["A"].hidden = True
    for col in "BCDEFGHIJKLM":
        ws.column_dimensions[col].width = 13
    ws.column_dimensions["B"].width = 15
    banner(ws, "06  WEEKLY WOR REPORT",
           "Auto-generated: one summary per project matching the filters (latest update on/before the As-Of date from "
           "06A). Print: landscape, 1 page wide, 5 projects per page.", 13, demo)
    labels = [("B4", "Department"), ("D4", "Shift"), ("F4", "Owner (leader or co-owner)"), ("I4", "Include completed?")]
    for c, t in labels:
        ws[c] = t
        ws[c].font = F(9, True, NAVY)
        ws[c].alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
    for c, v, nm, lst in [("C4", "(All)", "WOR_FDept", "lst_FDept"), ("E4", "(All)", "WOR_FShift", "lst_FShift"),
                          ("G4", "(All)", "WOR_FOwner", "lst_FOwner"), ("J4", "No", "WOR_FIncl", "lst_YesNo")]:
        ws[c] = v
        style_input(ws[c])
        ws[c].font = F(10, True, INPUT_FONT)
        ws[c].alignment = Alignment(vertical="center")
        dv_list(ws, lst, c)
        add_name(wb, nm, f"{q(ws.title)}!${c[0]}${c[1:]}")
    ws.merge_cells("G4:H4")
    ws["K4"] = '="Projects shown: "&MAX(0,MAX(reg_WorSeq))&IF(MAX(0,MAX(reg_WorSeq))>10,"  (first 10 shown - narrow the filters)","")'
    ws.merge_cells("K4:M4")
    ws["K4"].font = F(9, True, NAVY)
    ws["K4"].alignment = LEFT_C
    ws.row_dimensions[4].height = 28
    d = q(dash_title)
    ws["B5"] = (f'="Week ending "&TEXT(AsOfDate,"mm/dd/yyyy")&"   |   Active projects: "&{d}!B7&"   |   Action completion: "&'
                f'IF(ISNUMBER({d}!F7),TEXT({d}!F7,"0%"),"N/A")&"   |   Overdue actions: "&{d}!H7&"   |   Verified hard savings (Finance-approved): "&'
                f'TEXT({d}!L7,"$#,##0")&"   |   Forecast hard savings (unverified): "&TEXT({d}!J7,"$#,##0")&"   |   Leadership escalations: "&{d}!P7')
    ws.merge_cells("B5:M5")
    ws["B5"].font = F(9, True, WHITE)
    ws["B5"].fill = fill(STEEL)
    ws["B5"].alignment = LEFT_C
    ws.row_dimensions[5].height = 30
    ws["B6"] = ("Legend: ACTUAL = latest recorded KPI value | TARGET/BASELINE from the register | Savings: VERIFIED = Finance-approved, "
                "FORECAST = unverified | N/A or '-' = not yet entered | STALE = no update for this week. Enter updates in '06A WOR Weekly Updates'.")
    ws.merge_cells("B6:M6")
    ws["B6"].font = F(7, False, GRAY_D, True)
    ws["B6"].alignment = LEFT_C

    def RG(key, s, blank="-"):
        return f'IF($A${s}="","",IF(INDEX(reg_{key},$A${s})&""="","{blank}",INDEX(reg_{key},$A${s})))'

    def UP(key, s):
        return f'IF($A${s}="","",IF($A${s + 2}="","No update recorded",IF(INDEX(wor_{key},$A${s + 2})&""="","-",INDEX(wor_{key},$A${s + 2}))))'

    for k in range(1, 11):
        s = 8 + (k - 1) * 7
        ws[f"A{s}"] = f'=IFERROR(MATCH({k},reg_WorSeq,0),"")'
        ws[f"A{s + 1}"] = f'=IF($A${s}="","",_xlfn.MAXIFS(wor_Week,wor_Proj,INDEX(reg_ID,$A${s}),wor_Week,"<="&AsOfDate))'
        ws[f"A{s + 2}"] = f'=IF($A${s}="","",IF($A${s + 1}=0,"",IFERROR(MATCH(INDEX(reg_ID,$A${s})&"|"&$A${s + 1},wor_Key,0),"")))'
        # header
        merge_write(ws, f"B{s}:F{s}", f'={RG("ID", s)}&IF($A${s}="",""," | "&{RG("Title", s)})', F(10, True, WHITE), fill(NAVY), LEFT_C)
        merge_write(ws, f"G{s}:I{s}", f'=IF($A${s}="","","Leader: "&{RG("Leader", s, "NOT ASSIGNED")}&IF(INDEX(reg_CoOwner,$A${s})&""="","","  |  Co: "&INDEX(reg_CoOwner,$A${s})))',
                    F(8, True, WHITE), fill(NAVY), LEFT_C)
        merge_write(ws, f"J{s}:K{s}", f'=IF($A${s}="","",{RG("Dept", s)}&"  |  Shift "&{RG("Shift", s)})', F(8, True, WHITE), fill(NAVY), LEFT_C)
        merge_write(ws, f"L{s}:L{s}", f'=IF($A${s}="","",{RG("Stage", s)})', F(8, True, WHITE), fill(NAVY), CENTER)
        merge_write(ws, f"M{s}:M{s}", f'=IF($A${s}="","",{RG("RAG", s)})', F(9, True), fill(WHITE), CENTER, border=BORDER)
        ws.row_dimensions[s].height = 22
        # labels row
        lab = [("B", "G", "Goal / SMART objective"), ("H", "H", "KPI"), ("I", "I", "Baseline"), ("J", "J", "Actual (latest)"),
               ("K", "K", "Target"), ("L", "L", "Outcome attainment"), ("M", "M", "Action completion")]
        for a, b, t in lab:
            merge_write(ws, f"{a}{s + 1}:{b}{s + 1}", f'=IF($A${s}="","","{t}")', F(7, True, GRAY_D), fill(GRAY_XL), LEFT_C)
        goal = (f'=IF($A${s}="","",IFERROR(IF(INDEX(chr_SMART,MATCH(INDEX(reg_ID,$A${s}),chr_Proj,0))&""="","[SMART objective not yet defined] "&INDEX(reg_Desc,$A${s}),'
                f'INDEX(chr_SMART,MATCH(INDEX(reg_ID,$A${s}),chr_Proj,0))),"[No charter record] "&INDEX(reg_Desc,$A${s})))')
        merge_write(ws, f"B{s + 2}:G{s + 2}", goal, F(8), None, WRAP_TOP)
        vals = [("H", f'=IF($A${s}="","",{RG("KPI", s, "KPI not defined")}&IF(INDEX(reg_Unit,$A${s})&""="",""," ("&INDEX(reg_Unit,$A${s})&")"))', None),
                ("I", f'={RG("Baseline", s, "N/A")}', FMT_NUM), ("J", f'={RG("Actual", s, "N/A")}', FMT_NUM),
                ("K", f'={RG("Target", s, "N/A")}', FMT_NUM), ("L", f'={RG("Attain", s, "N/A")}', FMT_PCT),
                ("M", f'={RG("ActPct", s, "N/A")}', FMT_PCT)]
        for col, f, fmt in vals:
            c = ws[f"{col}{s + 2}"]
            c.value = f
            c.font = F(9, True)
            c.alignment = CENTER
            if fmt:
                c.number_format = fmt
        ws.row_dimensions[s + 2].height = 34
        lab2 = [("B", "E", "Work completed this week"), ("F", "I", "Next-week commitments"), ("J", "K", "Risks / support needed"),
                ("L", "M", "Leadership decisions")]
        for a, b, t in lab2:
            merge_write(ws, f"{a}{s + 3}:{b}{s + 3}", f'=IF($A${s}="","","{t}")', F(7, True, GRAY_D), fill(GRAY_XL), LEFT_C)
        merge_write(ws, f"B{s + 4}:E{s + 4}", "=" + UP("Done", s), F(8), None, WRAP_TOP)
        merge_write(ws, f"F{s + 4}:I{s + 4}", "=" + UP("Next", s), F(8), None, WRAP_TOP)
        merge_write(ws, f"J{s + 4}:K{s + 4}",
                    f'=IF($A${s}="","","Risks: "&{UP("Risks", s)}&CHAR(10)&"Support: "&{UP("Support", s)}&CHAR(10)&"Open risk items: "&INDEX(reg_Risks,$A${s}))',
                    F(8), None, WRAP_TOP)
        merge_write(ws, f"L{s + 4}:M{s + 4}",
                    f'=IF($A${s}="","","Requested: "&{UP("DecReq", s)}&CHAR(10)&"Recorded: "&{UP("DecRec", s)})', F(8), None, WRAP_TOP)
        ws.row_dimensions[s + 4].height = 58
        summ = (f'=IF($A${s}="","","Milestones: "&INDEX(reg_NMSDone,$A${s})&"/"&INDEX(reg_NMS,$A${s})&"   |   Overdue actions: "&INDEX(reg_Overdue,$A${s})'
                f'&"   |   VERIFIED savings (Finance-approved): "&TEXT(INDEX(reg_VerHard,$A${s})+INDEX(reg_VerAvoid,$A${s}),"$#,##0")'
                f'&"   |   FORECAST hard $/yr (unverified): "&IF(ISNUMBER(INDEX(reg_ExpHard,$A${s})),TEXT(INDEX(reg_ExpHard,$A${s}),"$#,##0"),INDEX(reg_ExpHard,$A${s}))'
                f'&"   |   Last update: "&IF(N($A${s + 1})=0,"none",TEXT($A${s + 1},"mm/dd/yyyy"))&IF(OR(N($A${s + 1})=0,AsOfDate-N($A${s + 1})>6),"  - STALE: no update this week",""))')
        merge_write(ws, f"B{s + 5}:M{s + 5}", summ, F(7, True, STEEL), None, LEFT_C)
        for rr in range(s, s + 6):
            for col in range(2, 14):
                cell = ws.cell(rr, col)
                cell.border = Border(left=THIN if col == 2 else None, right=THIN if col == 13 else None,
                                     top=THIN if rr == s else None, bottom=THIN if rr == s + 5 else None)
        ws.conditional_formatting.add(f"B{s}:M{s + 5}", FormulaRule(formula=[f'$A${s}=""'], fill=fill(WHITE),
                                                                      font=Font(name=FONT, color=WHITE)))
        cf_rag(ws, f"M{s}:M{s}")
        cf_na(ws, f"I{s + 2}:M{s + 2}")
        cf_formula(ws, f"B{s + 5}:M{s + 5}", f'ISNUMBER(SEARCH("STALE",B{s + 5}))', None, A_FONT, True)
        for col in "A":
            for rr in range(s, s + 3):
                ws[f"A{rr}"].font = F(7, False, GRAY_M)
    ws.print_area = "B1:M77"
    ws.print_title_rows = "1:6"
    ws.row_breaks.append(Break(id=42))
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.3
    ws.page_margins.top = ws.page_margins.bottom = 0.4
    ws.oddFooter.left.text = "C.H. Guenther & Son - Duncanville | Weekly WOR Report"
    ws.oddFooter.right.text = "Printed &D | Page &P of &N"
    ws.freeze_panes = "B7"


RUBRIC = {
    "Participation": ("Attends sporadically; waits to be told what to do.", "Attends WOR; completes assigned tasks with reminders.",
                      "Leads the project actively; drives the team without prompting.", "Champions CI; mentors peers and volunteers for new projects."),
    "CharterQ": ("Charter missing or incomplete.", "Charter drafted; problem, scope or target vague.",
                 "Charter complete with clear scope, SMART objective, baseline and target.", "Charter is a model for others; reviewed and approved on time."),
    "ProbDef": ("Problem stated as a solution or opinion.", "Problem described but not measurable.",
                "Problem stated with facts: what, where, when, magnitude.", "Problem framed with data, customer impact and clear boundaries."),
    "RCA": ("Jumps to solutions.", "Uses 5 Why / fishbone with help; stops at symptoms.",
            "Identifies verified root causes using structured tools.", "Validates root causes with data/trials and teaches the method."),
    "Data": ("Decisions based on opinion.", "Collects some data; inconsistent measures.",
             "Uses baseline, trend and target data to make decisions.", "Builds reliable measurement systems; data drives every decision."),
    "Teamwork": ("Works alone; limited engagement of others.", "Involves team when asked.",
                 "Engages operators, Maintenance, Warehouse and Finance appropriately.", "Builds cross-functional collaboration; recognises others' contributions."),
    "Closure": ("Actions frequently overdue; no follow-up.", "Most actions closed; some late without escalation.",
                "Actions closed on time with evidence; blockers escalated early.", "Consistently on time; helps others clear their blockers."),
    "Comms": ("Updates missing or unclear.", "Provides updates when asked; lacks data.",
              "Concise WOR updates with data, risks and clear asks.", "Communicates crisply to leadership; influences decisions with facts."),
    "Sustain": ("No standard or control plan.", "Standard drafted; not verified on the floor.",
                "Standard work implemented, trained and verified.", "Controls audited, results sustained, lessons shared plant-wide."),
}


def build_rubric(wb, ws, demo):
    ws.sheet_view.showGridLines = False
    widths = dict(A=2, B=26, C=28, D=28, E=28, F=28, G=12, H=12, I=12, J=12, K=12, L=12, M=12, N=13, O=13)
    for k, v in widths.items():
        ws.column_dimensions[k].width = v
    banner(ws, "07A  SUPERVISOR DEVELOPMENT - COMPETENCY RUBRIC AND PARTICIPATION SUMMARY",
           "Development focus: leadership capability, collaboration and execution discipline. Supervisors are NOT ranked, and "
           "savings dollars are intentionally excluded from development assessment.", 15, demo)
    section(ws, "B5:F5", "COMPETENCY RUBRIC (score each area 1-4 in '07 Supervisor Development')")
    for j, h in enumerate(["Competency", "1 - Beginning", "2 - Developing", "3 - Proficient", "4 - Role Model"]):
        c = ws.cell(6, 2 + j, h)
        c.font = F(9, True, WHITE)
        c.fill = fill(NAVY)
        c.alignment = CENTER
        c.border = BORDER
    for i, (k, name) in enumerate(COMPETENCIES):
        r = 7 + i
        ws.cell(r, 2, name).font = F(9, True, NAVY)
        for j, t in enumerate(RUBRIC[k]):
            c = ws.cell(r, 3 + j, t)
            c.font = F(8)
            c.alignment = WRAP_TOP
            c.border = BORDER
        ws.cell(r, 2).border = BORDER
        ws.cell(r, 2).alignment = WRAP_TOP
        ws.row_dimensions[r].height = 36
    ws["B17"] = ("Overall level = simple average of the areas scored (unscored areas excluded): >=3.5 Role Model, >=2.5 Proficient, "
                 ">=1.5 Developing, otherwise Beginning. Use the rubric for coaching conversations, not for ranking or pay decisions.")
    ws.merge_cells("B17:F17")
    ws["B17"].font = F(8, False, GRAY_D, True)
    ws["B17"].alignment = WRAP_TOP
    ws.row_dimensions[17].height = 28

    section(ws, "B19:O19", "PARTICIPATION AND EXECUTION SUMMARY  (listed in Master Data order - not a ranking)")
    hdr = ["Name", "Role", "Projects Led", "Projects Co-Owned", "Actions Owned", "Actions Closed", "Closure %",
           "On-Time Closure %", "Overdue Owned", "WOR Updates Submitted", "Latest Avg Competency", "Latest Review", "Next Review"]
    cols = "BCDEFGHIJKLMN"
    for j, h in enumerate(hdr):
        c = ws[f"{cols[j]}20"]
        c.value = h
        c.font = F(8, True, WHITE)
        c.fill = fill(NAVY)
        c.alignment = CENTER
        c.border = BORDER
    ws.row_dimensions[20].height = 30
    for k in range(25):
        r = 21 + k
        nm = f"$B{r}"
        f = {
            "B": f'=IF(INDEX(md_PeopleName,{k + 1})="","",INDEX(md_PeopleName,{k + 1}))',
            "C": f'=IF({nm}="","",IF(INDEX(md_PeopleRole,{k + 1})="","Not set",INDEX(md_PeopleRole,{k + 1})))',
            "D": f'=IF({nm}="","",COUNTIFS(reg_Leader,{nm},reg_ID,"?*"))',
            "E": f'=IF({nm}="","",COUNTIFS(reg_CoOwner,{nm},reg_ID,"?*"))',
            "F": f'=IF({nm}="","",COUNTIFS(act_Owner,{nm},act_Proj,"?*",act_Status,"<>Cancelled"))',
            "G": f'=IF({nm}="","",COUNTIFS(act_Owner,{nm},act_Credit,1))',
            "H": f'=IF({nm}="","",IF(F{r}=0,"N/A",G{r}/F{r}))',
            "I": f'=IF({nm}="","",IF(G{r}=0,"N/A",SUMIFS(act_OnTime,act_Owner,{nm})/G{r}))',
            "J": f'=IF({nm}="","",COUNTIFS(act_Owner,{nm},act_Flag,"OVERDUE"))',
            "K": f'=IF({nm}="","",COUNTIFS(wor_By,{nm},wor_Proj,"?*"))',
            "L": f'=IF({nm}="","",IF(COUNTIFS(dev_Sup,{nm},dev_Date,"<>")=0,"Not assessed",IFERROR(INDEX(dev_Avg,MATCH({nm}&"|"&M{r},dev_Key,0)),"Not assessed")))',
            "M": f'=IF({nm}="","",IF(COUNTIFS(dev_Sup,{nm},dev_Date,"<>")=0,"None",_xlfn.MAXIFS(dev_Date,dev_Sup,{nm})))',
            "N": f'=IF({nm}="","",IF(COUNTIFS(dev_Sup,{nm},dev_NextRev,"<>")=0,"Not scheduled",_xlfn.MAXIFS(dev_NextRev,dev_Sup,{nm})))',
        }
        for col, v in f.items():
            c = ws[f"{col}{r}"]
            c.value = v
            c.font = F(8, col == "B")
            c.border = BORDER
            c.alignment = Alignment(horizontal="left" if col in "BC" else "center")
        ws[f"H{r}"].number_format = FMT_PCT
        ws[f"I{r}"].number_format = FMT_PCT
        ws[f"L{r}"].number_format = "0.0"
        ws[f"M{r}"].number_format = FMT_DATE
        ws[f"N{r}"].number_format = FMT_DATE
    cf_na(ws, "H21:L45")
    cf_formula(ws, "J21:J45", "AND(ISNUMBER(J21),J21>0)", R_FILL, R_FONT, True)
    cf_formula(ws, "N21:N45", "AND(ISNUMBER(N21),N21<AsOfDate)", R_FILL, R_FONT, True)
    ws.freeze_panes = "A5"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def build_tools(wb, ws, demo):
    ws.sheet_view.showGridLines = False
    widths = dict(A=2, B=24, C=15, D=15, E=15, F=15, G=15, H=15, I=15, J=15, K=2)
    for k, v in widths.items():
        ws.column_dimensions[k].width = v
    banner(ws, "08  CI TOOLS - SIMPLE TEMPLATES FOR SUPERVISORS",
           "Yellow cells are for input. One working copy per project: right-click this tab > Move or Copy > Create a copy, or "
           "print the tool, complete it on the floor, and log the scan/photo in '10 Evidence Register'.", 10, demo)
    r = 5
    breaks = []

    def head(r, title, purpose):
        section(ws, f"B{r}:J{r}", title)
        ws[f"B{r + 1}"] = purpose
        ws.merge_cells(f"B{r + 1}:J{r + 1}")
        ws[f"B{r + 1}"].font = F(8, False, GRAY_D, True)
        ws[f"B{r + 1}"].alignment = WRAP_TOP
        ws.row_dimensions[r + 1].height = 24
        # id strip
        for lab, a, b, dvn, fmt in [("Project ID", "C", "D", "lst_ProjectID", None), ("Owner", "F", "G", "lst_People", None),
                                    ("Date", "I", "J", None, FMT_DATE)]:
            lc = get_column_letter(openpyxl.utils.column_index_from_string(a) - 1)
            ws[f"{lc}{r + 2}"] = lab
            ws[f"{lc}{r + 2}"].font = F(9, True, NAVY)
            ws[f"{lc}{r + 2}"].alignment = Alignment(horizontal="right", vertical="center")
            ws.merge_cells(f"{a}{r + 2}:{b}{r + 2}")
            for cc in (a, b):
                style_input(ws[f"{cc}{r + 2}"], fmt)
            if dvn:
                dv_list(ws, dvn, f"{a}{r + 2}")
            else:
                dv_date(ws, f"{a}{r + 2}")
        return r + 3

    def field(r, label, h=30, a="C", b="J"):
        ws[f"B{r}"] = label
        ws[f"B{r}"].font = F(9, True, NAVY)
        ws[f"B{r}"].alignment = WRAP_TOP
        ws[f"B{r}"].fill = fill(GRAY_XL)
        ws[f"B{r}"].border = BORDER
        ws.merge_cells(f"{a}{r}:{b}{r}")
        for col in range(openpyxl.utils.column_index_from_string(a), openpyxl.utils.column_index_from_string(b) + 1):
            style_input(ws.cell(r, col), wrap=True)
        ws.row_dimensions[r].height = h
        return r + 1

    def thead(r, headers, start="B"):
        s = openpyxl.utils.column_index_from_string(start)
        for j, h in enumerate(headers):
            c = ws.cell(r, s + j, h)
            c.font = F(8, True, WHITE)
            c.fill = fill(NAVY)
            c.alignment = CENTER
            c.border = BORDER
        ws.row_dimensions[r].height = 30

    # 1. 5 WHY
    r = head(r, "TOOL 1 - 5 WHY ROOT CAUSE ANALYSIS",
             "Start from a clear, factual problem statement. Ask 'why' until you reach a cause you can act on and verify (often 3-5 times). "
             "Check each answer with data or by going to the floor.")
    r = field(r, "Problem statement (what, where, when, how much)", 36)
    for i in range(1, 6):
        r = field(r, f"Why {i}?  -  because ...", 28)
    r = field(r, "Root cause (verified?)", 30)
    r = field(r, "Countermeasure", 30)
    r = field(r, "Countermeasure owner / due date", 22)
    r = field(r, "How effectiveness will be verified", 30)
    breaks.append(r)
    r += 1

    # 2. PDCA
    r = head(r, "TOOL 2 - PDCA (PLAN - DO - CHECK - ACT)", "Use for each improvement trial. Keep it short - bullets are fine.")
    for (t1, p1), (t2, p2) in [(("PLAN", "Problem, target, root cause, what will be tried, how it will be measured"),
                                ("DO", "What was done, where, when, by whom; any deviations from the plan")),
                               (("CHECK", "Results vs. target with data; what worked / did not"),
                                ("ACT", "Standardize, adjust and repeat, or abandon; next PDCA cycle"))]:
        merge_write(ws, f"B{r}:E{r}", f"{t1} - {p1}", F(9, True, WHITE), fill(STEEL), LEFT_C)
        merge_write(ws, f"F{r}:J{r}", f"{t2} - {p2}", F(9, True, WHITE), fill(STEEL), LEFT_C)
        ws.row_dimensions[r].height = 28
        ws.merge_cells(f"B{r + 1}:E{r + 1}")
        ws.merge_cells(f"F{r + 1}:J{r + 1}")
        for col in range(2, 11):
            style_input(ws.cell(r + 1, col), wrap=True)
        ws.row_dimensions[r + 1].height = 90
        r += 2
    breaks.append(r)
    r += 1

    # 3. A3
    r = head(r, "TOOL 3 - A3 PROBLEM SOLVING", "Tell the improvement story on one page: left side = understand the problem; right side = solve and sustain it.")
    left = ["1. Background - why this matters", "2. Current condition - facts, data, observation", "3. Goal / target",
            "4. Root cause analysis"]
    right = ["5. Countermeasures", "6. Implementation plan (what / who / when)", "7. Follow-up and results"]
    r0 = r
    for i, t in enumerate(left):
        merge_write(ws, f"B{r0 + 2 * i}:E{r0 + 2 * i}", t, F(9, True, WHITE), fill(STEEL), LEFT_C)
        ws.merge_cells(f"B{r0 + 2 * i + 1}:E{r0 + 2 * i + 1}")
        for col in range(2, 6):
            style_input(ws.cell(r0 + 2 * i + 1, col), wrap=True)
        ws.row_dimensions[r0 + 2 * i + 1].height = 70
    for i, t in enumerate(right):
        rr = r0 + 2 * i
        merge_write(ws, f"F{rr}:J{rr}", t, F(9, True, WHITE), fill(STEEL), LEFT_C)
        ws.merge_cells(f"F{rr + 1}:J{rr + 1}")
        for col in range(6, 11):
            style_input(ws.cell(rr + 1, col), wrap=True)
    merge_write(ws, f"F{r0 + 6}:J{r0 + 6}", "Sign-off: leader / coach / date", F(9, True, WHITE), fill(STEEL), LEFT_C)
    ws.merge_cells(f"F{r0 + 7}:J{r0 + 7}")
    for col in range(6, 11):
        style_input(ws.cell(r0 + 7, col), wrap=True)
    r = r0 + 8
    breaks.append(r)
    r += 1

    # 4. 5S audit
    r = head(r, "TOOL 4 - 5S AUDIT", "Score each item 0-4: 0 = not in place, 1 = poor, 2 = partial, 3 = good, 4 = excellent / sustained. "
                                     "Leave blank if not applicable. Score % = points / (items scored x 4).")
    thead(r, ["5S Element", "Check item", "", "", "", "Score (0-4)", "Comment / action", "", ""])
    ws.merge_cells(f"C{r}:F{r}")
    ws.merge_cells(f"H{r}:J{r}")
    r += 1
    items = [("Sort", "Only needed items are in the area"), ("Sort", "Red-tag area used; unneeded items removed"),
             ("Sort", "No obsolete paperwork, tools or materials"),
             ("Set in Order", "Every item has a labeled home"), ("Set in Order", "Tools / parts within easy reach"),
             ("Set in Order", "Floor markings and signage clear"),
             ("Shine", "Equipment and floor clean; no leaks"), ("Shine", "Cleaning schedule posted and followed"),
             ("Shine", "Sources of contamination identified"),
             ("Standardize", "Visual standards / photos posted"), ("Standardize", "Checklists current and in use"),
             ("Standardize", "Responsibilities assigned"),
             ("Sustain", "Audits performed on schedule"), ("Sustain", "Previous audit actions closed"),
             ("Sustain", "Team can explain the standard")]
    s5 = r
    for s, t in items:
        ws.cell(r, 2, s).font = F(9, True, NAVY)
        ws.cell(r, 2).border = BORDER
        ws.cell(r, 3, t).font = F(9)
        ws.merge_cells(f"C{r}:F{r}")
        style_input(ws.cell(r, 7), "0")
        ws.merge_cells(f"H{r}:J{r}")
        for col in range(8, 11):
            style_input(ws.cell(r, col), wrap=True)
        for col in range(3, 7):
            ws.cell(r, col).border = BORDER
        r += 1
    dv_whole(ws, f"G{s5}:G{r - 1}", 0, 4, "Score 0-4 (blank if not applicable).")
    ws.cell(r, 2, "TOTAL").font = F(9, True)
    ws.cell(r, 6, "Points / possible").font = F(8, True, GRAY_D)
    ws.cell(r, 7, f'=IF(COUNT(G{s5}:G{r - 1})=0,"N/A",SUM(G{s5}:G{r - 1})&" / "&COUNT(G{s5}:G{r - 1})*4)')
    ws.cell(r, 8, "5S score %").font = F(8, True, GRAY_D)
    ws.cell(r, 9, f'=IF(COUNT(G{s5}:G{r - 1})=0,"N/A",SUM(G{s5}:G{r - 1})/(COUNT(G{s5}:G{r - 1})*4))').number_format = FMT_PCT
    for c in (ws.cell(r, 7), ws.cell(r, 9)):
        style_calc(c, c.number_format, bold=True)
    r += 1
    breaks.append(r)
    r += 1

    # 5. SMED
    r = head(r, "TOOL 5 - CHANGEOVER OBSERVATION / SMED",
             "Time each element of one real changeover (video helps). Internal = machine stopped; External = can be done while running. "
             "Downtime = sum of INTERNAL element times. Recovered minutes are capacity, not cash, until Finance validates a benefit.")
    thead(r, ["Element (step)", "Before: time (min)", "Before: type", "Improvement idea", "", "Action type",
              "After: time (min)", "After: type", "Notes"])
    ws.merge_cells(f"E{r}:F{r}")
    r += 1
    sm = r
    for i in range(15):
        style_input(ws.cell(r, 2), wrap=True)
        style_input(ws.cell(r, 3), "0.0")
        style_input(ws.cell(r, 4))
        ws.merge_cells(f"E{r}:F{r}")
        style_input(ws.cell(r, 5), wrap=True)
        style_input(ws.cell(r, 6))
        style_input(ws.cell(r, 7))
        style_input(ws.cell(r, 8), "0.0")
        style_input(ws.cell(r, 9))
        style_input(ws.cell(r, 10), wrap=True)
        r += 1
    for rng in (f"D{sm}:D{r - 1}", f"I{sm}:I{r - 1}"):
        dv = DataValidation(type="list", formula1='"Internal,External"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(rng)
    dv = DataValidation(type="list", formula1='"Eliminate,Externalize,Simplify,Keep"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"G{sm}:G{r - 1}")
    dv_decimal(ws, f"C{sm}:C{r - 1}", 0, 10000, "Minutes (number).")
    dv_decimal(ws, f"H{sm}:H{r - 1}", 0, 10000, "Minutes (number).")
    res = [("Total elapsed time before (min)", f"=SUM(C{sm}:C{r - 1})"),
           ("Downtime before - internal (min)", f'=SUMIF(D{sm}:D{r - 1},"Internal",C{sm}:C{r - 1})'),
           ("Downtime after - internal (min)", f'=IF(COUNT(H{sm}:H{r - 1})=0,"N/A",SUMIF(I{sm}:I{r - 1},"Internal",H{sm}:H{r - 1}))'),
           ("Downtime reduction (min per changeover)", f'=IF(ISNUMBER(D{r + 2}),D{r + 1}-D{r + 2},"N/A")'),
           ("Downtime reduction %", f'=IF(AND(ISNUMBER(D{r + 3}),D{r + 1}>0),D{r + 3}/D{r + 1},"N/A")')]
    for i, (lab, f) in enumerate(res):
        ws.cell(r + i, 2, lab).font = F(9, True, NAVY)
        ws.merge_cells(f"B{r + i}:C{r + i}")
        c = ws.cell(r + i, 4, f)
        style_calc(c, FMT_PCT if "%" in lab else "0.0", bold=True)
    ws.cell(r + 1, 6, "Enter verified minutes reduced and changeover frequency in '05 Financial Savings' only when supported by observations.").font = F(8, False, GRAY_D, True)
    ws.merge_cells(f"F{r + 1}:J{r + 2}")
    ws.cell(r + 1, 6).alignment = WRAP_TOP
    r += len(res)
    breaks.append(r)
    r += 1

    # 6. Pareto
    r = head(r, "TOOL 6 - WASTE PARETO INPUT",
             "List waste categories with quantity for the SAME period and unit. Optional cost per unit converts quantity to $. The table "
             "below sorts automatically (by $ if every row has a cost, otherwise by quantity) and shows the cumulative %.")
    thead(r, ["Waste category", "Quantity", "Unit", "Cost per unit ($, optional)", "Cost ($)", "Sort value (auto)"])
    r += 1
    pi = r
    for i in range(10):
        style_input(ws.cell(r, 2))
        style_input(ws.cell(r, 3), FMT_NUM)
        style_input(ws.cell(r, 4))
        style_input(ws.cell(r, 5), FMT_USD2)
        style_calc(ws.cell(r, 6, f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(E{r})),C{r}*E{r},"")'), FMT_USD)
        r += 1
    pe = r - 1
    dv_decimal(ws, f"C{pi}:C{pe}", 0, 1E12, "Quantity (number).")
    dv_decimal(ws, f"E{pi}:E{pe}", 0, 1E9, "Cost per unit (number).")
    ws.cell(r, 2, "Sort basis").font = F(8, True, GRAY_D)
    ws.cell(r, 3, f'=IF(COUNT(C{pi}:C{pe})=0,"No data",IF(COUNT(F{pi}:F{pe})=COUNT(C{pi}:C{pe}),"Cost ($)","Quantity"))')
    style_calc(ws.cell(r, 3), bold=True)
    basis = f"$C${r}"
    for i in range(10):
        rr = pi + i
        ws.cell(rr, 7, f'=IF(OR(B{rr}="",NOT(ISNUMBER(C{rr}))),"",IF({basis}="Cost ($)",F{rr},C{rr})+ROW()/1000000)')
        style_calc(ws.cell(rr, 7), "0.00")
    r += 2
    thead(r, ["Rank", "Waste category (sorted)", "Value", "% of total", "Cumulative %"])
    r += 1
    ps = r
    for k in range(1, 11):
        rr = r
        ws.cell(rr, 2, k)
        ws.cell(rr, 3, f'=IF(COUNT($G${pi}:$G${pe})<{k},"",INDEX($B${pi}:$B${pe},MATCH(LARGE($G${pi}:$G${pe},{k}),$G${pi}:$G${pe},0)))')
        ws.cell(rr, 4, f'=IF(C{rr}="","",ROUND(LARGE($G${pi}:$G${pe},{k}),2))')
        ws.cell(rr, 5, f'=IF(C{rr}="","",D{rr}/SUM($D${ps}:$D${ps + 9}))')
        ws.cell(rr, 6, f'=IF(C{rr}="","",SUM($D${ps}:D{rr})/SUM($D${ps}:$D${ps + 9}))')
        for col in range(2, 7):
            style_calc(ws.cell(rr, col))
        ws.cell(rr, 4).number_format = FMT_NUM
        ws.cell(rr, 5).number_format = FMT_PCT
        ws.cell(rr, 6).number_format = FMT_PCT
        r += 1
    ch = BarChart()
    ch.title = "Waste Pareto"
    ch.style = 2
    ch.add_data(Reference(ws, min_col=4, min_row=ps - 1, max_row=ps + 9), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=3, min_row=ps, max_row=ps + 9))
    ch.series[0].graphicalProperties.solidFill = NAVY
    ch.y_axis.title = "Value"
    ln = LineChart()
    ln.add_data(Reference(ws, min_col=6, min_row=ps - 1, max_row=ps + 9), titles_from_data=True)
    ln.y_axis.axId = 200
    ln.y_axis.title = "Cumulative %"
    ln.y_axis.number_format = "0%"
    ln.y_axis.scaling.min = 0
    ln.y_axis.scaling.max = 1
    ln.y_axis.crosses = "max"
    ln.y_axis.delete = False
    ln.series[0].graphicalProperties.line.solidFill = "C0504D"
    ch += ln
    ch.height = 7
    ch.width = 14
    ch.legend.position = "b"
    ch.legend.overlay = False
    ch.roundedCorners = False
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.y_axis.majorGridlines = None
    ws.add_chart(ch, f"G{ps - 1}")
    breaks.append(r)
    r += 1

    # 7. Before / after
    r = head(r, "TOOL 7 - BEFORE-AND-AFTER IMPROVEMENT EVIDENCE",
             "Record measured results only. Link photos/data files (SharePoint/Teams) and register them in '10 Evidence Register'.")
    thead(r, ["Metric", "Unit", "Before value", "Before date", "After value", "After date", "Change", "% change", "Evidence link / ref"])
    r += 1
    bs = r
    for i in range(8):
        style_input(ws.cell(r, 2), wrap=True)
        style_input(ws.cell(r, 3))
        style_input(ws.cell(r, 4), FMT_NUM)
        style_input(ws.cell(r, 5), FMT_DATE)
        style_input(ws.cell(r, 6), FMT_NUM)
        style_input(ws.cell(r, 7), FMT_DATE)
        style_calc(ws.cell(r, 8, f'=IF(AND(ISNUMBER(D{r}),ISNUMBER(F{r})),F{r}-D{r},IF(B{r}="","","N/A"))'), FMT_NUM)
        style_calc(ws.cell(r, 9, f'=IF(AND(ISNUMBER(D{r}),ISNUMBER(F{r}),N(D{r})<>0),(F{r}-D{r})/D{r},IF(B{r}="","","N/A"))'), FMT_PCT)
        style_input(ws.cell(r, 10), wrap=True)
        r += 1
    dv_decimal(ws, f"D{bs}:D{r - 1}", msg="Number only.")
    dv_decimal(ws, f"F{bs}:F{r - 1}", msg="Number only.")
    dv_date(ws, f"E{bs}:E{r - 1}")
    dv_date(ws, f"G{bs}:G{r - 1}")
    r = field(r, "Photo / visual evidence (describe and link; Excel does not store the files)", 40)
    breaks.append(r)
    r += 1

    # 8. standard work verification
    r = head(r, "TOOL 8 - STANDARD WORK VERIFICATION",
             "Observe the operator performing the task against the current standard. Observed: Yes = 1, Partial = 0.5, No = 0, N/A excluded.")
    r = field(r, "Standard / SOP name and revision", 22)
    r = field(r, "Area / line / shift observed", 22)
    thead(r, ["Step #", "Standard requirement", "", "", "Observed?", "Deviation / comment", "", "Corrective action", ""])
    ws.merge_cells(f"C{r}:E{r}")
    ws.merge_cells(f"G{r}:H{r}")
    ws.merge_cells(f"I{r}:J{r}")
    r += 1
    sw = r
    for i in range(12):
        style_input(ws.cell(r, 2))
        ws.merge_cells(f"C{r}:E{r}")
        ws.merge_cells(f"G{r}:H{r}")
        ws.merge_cells(f"I{r}:J{r}")
        for col in range(3, 11):
            style_input(ws.cell(r, col), wrap=True)
        r += 1
    dv = DataValidation(type="list", formula1='"Yes,Partial,No,N/A"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"F{sw}:F{r - 1}")
    ws.cell(r, 2, "Compliance %").font = F(9, True, NAVY)
    c = ws.cell(r, 3, f'=IF(COUNTIF(F{sw}:F{r - 1},"Yes")+COUNTIF(F{sw}:F{r - 1},"Partial")+COUNTIF(F{sw}:F{r - 1},"No")=0,"N/A",'
                      f'(COUNTIF(F{sw}:F{r - 1},"Yes")+0.5*COUNTIF(F{sw}:F{r - 1},"Partial"))/(COUNTIF(F{sw}:F{r - 1},"Yes")+COUNTIF(F{sw}:F{r - 1},"Partial")+COUNTIF(F{sw}:F{r - 1},"No")))')
    style_calc(c, FMT_PCT, bold=True)
    r += 1
    r = field(r, "Verified by / date", 22)
    for b in breaks:
        ws.row_breaks.append(Break(id=b))
    ws.print_area = f"A1:J{r}"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    cf_na(ws, f"B5:J{r}")


def write_doc_sheet(ws, demo, title, subtitle, blocks, widths=(2, 30, 110)):
    ws.sheet_view.showGridLines = False
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(i + 1)].width = w
    banner(ws, title, subtitle, len(widths), demo)
    r = 5
    for kind, a, b in blocks:
        if kind == "h":
            section(ws, f"B{r}:{get_column_letter(len(widths))}{r}", a)
            r += 1
        elif kind == "p":
            ws.cell(r, 2, a).font = F(9, True, NAVY)
            ws.cell(r, 2).alignment = WRAP_TOP
            c = ws.cell(r, 3, b)
            c.font = F(9)
            c.alignment = WRAP_TOP
            lines = max(1, int(len(b) / 105) + 1 + b.count("\n"))
            ws.row_dimensions[r].height = min(400, 13 * lines + 4)
            for col in (2, 3):
                ws.cell(r, col).border = BOTTOM
            r += 1
        elif kind == "sp":
            r += 1
    ws.freeze_panes = "A5"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return r


GUIDE = [
    ("h", "PURPOSE AND ROLES", None),
    ("p", "What this workbook is", "A working management system for supervisor-led continuous improvement at Duncanville. It holds the project register, "
          "charters, actions, savings, weekly WOR updates, evidence links, risks and supervisor development records, and produces the dashboard and WOR "
          "report automatically. Collaboration, leadership development and measurable improvement matter more than savings alone."),
    ("p", "Supervisors (Production, Maintenance, Warehouse)", "LEAD and OWN projects. Each project has ONE primary owner (Project Leader) who is accountable; "
          "co-owners and team members support. Supervisors keep their charter, actions, weekly update and evidence current."),
    ("p", "CI Manager (Moe Abendeh)", "Coaches, facilitates and provides tools. Reviews charters, checks data quality, scores development rubrics with "
          "supervisors, maintains Master Data and settings, and prepares the WOR pack."),
    ("p", "Finance (including Randy)", "Validates financial benefits. Only Finance changes Finance Status, Finance Review Date and Realized Savings. "
          "Unapproved benefits are never counted as verified savings."),
    ("h", "QUICK-START (SUPERVISORS) - 10 STEPS", None),
    ("p", "1. Open the file", "Open the master copy from the approved SharePoint/Teams location in desktop Excel (Microsoft 365 recommended). Do not e-mail "
          "copies - one master file avoids conflicting versions."),
    ("p", "2. Find your project", "'02 Project Register'. Filter Project Leader to your name. Yellow cells are inputs; gray cells are formulas (locked)."),
    ("p", "3. Complete the register", "Fill department, area, shift, sponsor, start and target dates, primary KPI, unit, and the MEASURED baseline and target. "
          "Leave unknown values BLANK - blanks show N/A and are flagged; never type 0 for missing data."),
    ("p", "4. Write the charter", "'03A Charter Records': complete the narrative sections for your Project ID (business problem, scope, current condition, "
          "SMART objective, stakeholders, responsibilities, benefits, constraints, risks). View/print it on '03 Project Charter' by selecting the Project ID."),
    ("p", "5. Plan actions", "'04 Action Tracker': one row per action. Pick the Project ID, describe the action, assign ONE accountable owner, set the due date. "
          "Set Milestone = Yes for key charter milestones (they appear on the charter and as 'Next Milestone')."),
    ("p", "6. Update actions", "When an action is done set Status = Complete AND enter the Completed Date - both are required for completion credit. "
          "Use Blocker / Escalation Needed for problems, and log escalations in '09 Risks and Escalations'."),
    ("p", "7. Attach evidence", "Save the file (photo, report, SOP) in the approved SharePoint/Teams project folder, use 'Copy link', and add a row to "
          "'10 Evidence Register' with the Project ID and Action ID. Excel stores the LINK only."),
    ("p", "8. Weekly update (before WOR)", "'06A WOR Weekly Updates': add ONE row per project per week: week-ending date, work completed, next-week "
          "commitments, KPI actual (measured), risks, support needed and decisions requested."),
    ("p", "9. Savings", "Enter benefit lines in '05 Financial Savings' with data source and assumptions, Finance Status = Not Submitted. Ask Finance to review "
          "(Under Review). Do not enter Realized Savings yourself."),
    ("p", "10. Check before the meeting", "Open '06 Weekly WOR Report' and your row on '01 Executive Dashboard'. Fix any OVERDUE, MISSING EVIDENCE, "
          "STALE or N/A flags you can resolve."),
    ("h", "CREATING A NEW PROJECT", None),
    ("p", "Add the project", "Type the title in the first empty row of '02 Project Register' (no blank rows between projects). The Project ID is created "
          "automatically (prefix + row number). Choose a Project Leader from the dropdown - new people must first be added in '11 Master Data' > People."),
    ("p", "Maintenance / Warehouse placeholders", "Rows DUN-CI-004 and DUN-CI-005 are placeholders awaiting confirmed assignments. When an owner is confirmed, "
          "overwrite the title and complete the row, or keep the placeholder and add a new project per confirmed assignment, then set the placeholder "
          "Status to Cancelled."),
    ("p", "Owner assignment", "Assign the primary owner, co-owner, sponsor, CI coach and Finance reviewer with dropdowns in the register. Use "
          "'11 Master Data' > Owner Assignment Lookup to see each person's load (projects, open/overdue actions, risks)."),
    ("h", "PREPARING FOR WOR", None),
    ("p", "Weekly rhythm", "By the day before WOR: supervisors update actions, evidence and 06A. CI Manager sets the Report As-Of Date (Master Data C6) to the "
          "meeting date if a frozen view is wanted, checks the Data Integrity panel on the dashboard (all PASS), and prints the dashboard (one page) and "
          "the WOR report (landscape, 5 projects per page)."),
    ("p", "Filters", "On '06 Weekly WOR Report' choose Department, Shift, Owner and whether to include completed projects. Up to 10 projects are shown."),
    ("p", "In the meeting", "Focus on RED/AMBER projects, overdue actions, leadership escalations and decisions requested. Record decisions in 06A "
          "(Leadership Decision Recorded) and assign follow-up actions in 04."),
    ("h", "CI MANAGER REVIEW WORKFLOW", None),
    ("p", "Charter review", "Review each charter on '03 Project Charter'. Record CI Coach Review Date and Charter Status in 03A. Status Approved only when the "
          "problem, scope, SMART objective, baseline and target are defined (Completeness % helps)."),
    ("p", "Weekly data quality", "Filter '04 Action Tracker' Data Check <> OK, Evidence Check = MISSING EVIDENCE, and verify evidence in '10 Evidence Register' "
          "(Verification Status). Resolve any CHECK on the dashboard integrity panel."),
    ("p", "Development", "Quarterly (or as agreed) score supervisors 1-4 on '07 Supervisor Development' using the rubric on 07A, together with the supervisor. "
          "Record strengths, coaching notes, development actions and the next review date. Never rank supervisors by savings."),
    ("h", "FINANCE REVIEW WORKFLOW", None),
    ("p", "1. Submission", "Supervisor enters the benefit line with method, inputs, data source and assumptions (Not Submitted), then requests review."),
    ("p", "2. Review", "Finance sets Under Review, checks data, unit costs, volume comparability, the source key (double counting) and the benefit type "
          "(Hard Savings vs Cost Avoidance vs Productivity Capacity)."),
    ("p", "3. Decision", "Approved: enter Finance Reviewer, Review Date and Realized Savings to Date (project-attributable, validated). Rejected: record the reason "
          "in Finance Comments. Only Approved lines feed Verified Savings."),
    ("p", "Changeover minutes", "Recovered minutes are capacity, not cash. Set 'Capacity / Cost Benefit Demonstrated?' = Yes and an approved $/min only when "
          "the recovered time produces real output or cost reduction (e.g., eliminated overtime, added sellable volume)."),
    ("h", "MULTI-USER, SECURITY AND EVIDENCE LIMITATIONS", None),
    ("p", "Shared use", "Store the master .xlsx in SharePoint/OneDrive. Microsoft 365 co-authoring lets several people edit at once, but Excel has no row-level "
          "ownership security: anyone with edit access can change any input. Use SharePoint version history to review or restore changes."),
    ("p", "Sheet protection", "Formula cells are locked with sheet protection (no password) to prevent accidental overwrites. It is NOT a security control. "
          "To extend the structure: Review > Unprotect Sheet."),
    ("p", "Evidence storage", "Excel cannot upload, store or secure documents. The Evidence Register holds hyperlinks to files kept in the approved "
          "SharePoint/Teams folder. Viewers need permission to that folder; links do not grant access. There is no live SharePoint synchronization."),
    ("p", "Signatures", "Approval dates in 03A record review only; they are not secure electronic signatures."),
    ("h", "TROUBLESHOOTING", None),
    ("p", "Dropdown is empty / name missing", "Add the value in '11 Master Data' directly below the last entry of the list (no blank rows)."),
    ("p", "Shows N/A", "Required data is missing (baseline, target, actual, or savings inputs). Enter measured data; do not enter zero."),
    ("p", "Action not counted complete", "Status must be Complete AND Completed Date entered."),
    ("p", "Project missing on WOR report", "Check the filters (C4/E4/G4/J4) and that the project is not Cancelled (or Complete with 'Include completed?' = No)."),
    ("p", "Capacity limits", "Register 150 projects, actions 1,000, savings lines 300, weekly updates 2,000, risks 300, evidence 1,000, reviews 300. See "
          "'13 Data Dictionary' > Maintenance to extend."),
]


def dd_blocks():
    K = [
        ("Active Projects", "Registered projects whose status is not Complete, Cancelled or Pending Assignment (blank status counts as active).",
         "02 reg_Status", "Dashboard card"),
        ("Completed Projects", "Projects with Status = Complete.", "02", "Dashboard"),
        ("Projects Pending Assignment", "Status = Pending Assignment (e.g., Maintenance/Warehouse placeholders).", "02", "Dashboard"),
        ("Action Completion %", "Credited complete actions / non-cancelled actions. Credited = Status Complete AND Completed Date entered. Status-based count; the "
         "'Reported % Complete' field is informational and is never averaged.", "04 act_Credit, act_Status", "Dashboard, register, WOR"),
        ("Milestone Completion %", "Credited milestones / non-cancelled actions flagged Milestone = Yes.", "04 act_MS", "Dashboard, register, charter"),
        ("Outcome Attainment %", "(Baseline - Latest Actual) / (Baseline - Target). 100% = target achieved, 0% = at baseline, negative = worse than baseline. "
         "Works for increase and decrease goals. N/A if baseline, target or actual missing, or baseline = target.", "02 baseline/target, 06A KPI actual",
         "Register, charter, WOR"),
        ("Overdue Action", "Not credited complete, not cancelled, due date earlier than the As-Of date.", "04", "All"),
        ("Due Soon", "Open action due within the due-soon window (default 7 days).", "04, Settings", "Actions, dashboard"),
        ("Days Remaining", "Due date - As-Of date for open actions (negative = days overdue).", "04", "Actions"),
        ("RAG Status", "RED: Blocked, overdue actions >= threshold (default 3), or past target date. AMBER: any overdue action, open risk, no primary owner, "
         "no target date, or baseline missing after Define. GREEN otherwise. PENDING = awaiting assignment. CLOSED = Complete/Cancelled.", "02, 04, 09",
         "All"),
        ("Projects Requiring Attention", "Count of projects with RAG = RED, AMBER or PENDING. Reasons listed in 'Attention Reasons'.", "02", "Dashboard"),
        ("Material Savings / Period", "(Baseline loss rate - Current loss rate) x Comparison production volume x Validated unit cost. Normalizes for volume "
         "(loss rates are per unit produced). Flags volume shifts above tolerance.", "05", "Savings"),
        ("Changeover Value / Period", "Minutes reduced x changeovers per period x Finance-approved $/min, ONLY when 'benefit demonstrated' = Yes; otherwise $0 "
         "cash. Capacity hours = minutes x changeovers / 60 are always reported separately (non-cash).", "05", "Savings"),
        ("SAP / Other Benefits", "Labor, downtime and error/rework savings per period entered from supporting data only. Lines without a data source are flagged.",
         "05", "Savings"),
        ("Annualization", "Annualized recurring = Recurring benefit per period x (Operating weeks per year / Period length in weeks). Default 52 weeks is an "
         "assumption pending Finance confirmation (Master Data C9).", "05, Settings", "Savings"),
        ("Gross / Net Savings", "Gross = Annualized recurring + One-time benefit. Net = Gross - Implementation cost. Attributed net = Net x Attribution % "
         "(blank = 100%).", "05", "Savings, register"),
        ("Projected Annual Savings (Forecast)", "Sum of attributed net forecast for Hard Savings lines not Rejected. Unverified estimate.", "05", "Dashboard"),
        ("Verified Realized Savings", "Sum of Finance-validated Realized Savings on lines with Finance Status = Approved. All other statuses contribute 0.",
         "05", "Dashboard, register, WOR"),
        ("Cost Avoidance / Capacity", "Tracked separately from hard savings. Capacity hours/yr are non-cash.", "05", "Dashboard"),
        ("Double-Count Control", "Lines sharing a Benefit Source Key draw from one cost pool; the sum of their attribution % (excluding Rejected) may not exceed "
         "100% (flag OVER-ATTRIBUTED).", "05", "Savings, dashboard check"),
        ("Charter Completeness %", "Filled fields / 13 (10 narrative sections in 03A + baseline, target and target date in 02). Objective completeness, not quality.",
         "03A, 02", "Charter"),
        ("Competency Average", "Average of the 1-4 scores entered (unscored areas excluded). Level bands: >=3.5 Role Model, >=2.5 Proficient, >=1.5 "
         "Developing, else Beginning.", "07", "07, 07A"),
        ("Closure % / On-Time %", "Owner's credited actions / non-cancelled actions owned; on-time = completed on or before due date.", "04", "07A"),
        ("Leadership Action Needed", "Open risk/blocker that is High impact, Escalated, escalated to Plant Leadership, or past its resolution target.", "09",
         "Risks, dashboard"),
        ("Missing Evidence", "Credited complete action with no Evidence Register record and no evidence reference.", "04, 10", "Actions, dashboard"),
        ("Trends", "Weekly points end on the As-Of date (12 weeks). Planned (cum.) = actions due on/before week end; Completed (cum.) = credited actions with "
         "completed date on/before week end; Overdue = due before week end and not completed by then. Monthly forecast run-rate accumulates each Hard Savings "
         "line's annualized recurring / 12 from its start month; verified is dated by Finance review date.", "04, 05", "Dashboard"),
    ]
    blocks = [("h", "KPI DEFINITIONS (formula-driven - no dashboard value is typed manually)", None)]
    for name, defi, src, where in K:
        blocks.append(("p", name, f"{defi}\nSource: {src}   |   Shown on: {where}"))
    blocks += [
        ("h", "DATA STATES", None),
        ("p", "ACTUAL", "A value recorded from measurement or system data (counts, dates, KPI actuals)."),
        ("p", "TARGET / BASELINE", "Entered in the register from measured data and the approved charter."),
        ("p", "FORECAST", "An unverified estimate (savings lines not yet Approved). Shown in gray italic on the dashboard."),
        ("p", "VERIFIED", "Finance-approved realized savings only."),
        ("p", "MISSING (N/A, Not quantified, Not assessed, [Not yet defined])", "Required data not yet entered. Never treated as zero in ratios; sums "
         "exclude it and flag it."),
        ("h", "MAINTENANCE", None),
        ("p", "Adding people / list values", "Type below the last entry of the list in '11 Master Data' (no blank rows). Dropdowns use dynamic named ranges "
         "(OFFSET/COUNTA) so no formula changes are needed."),
        ("p", "System lists", "Stages, statuses, Finance statuses, benefit types, calculation methods, verification/risk statuses, escalation levels and Yes/No "
         "are referenced by formulas. Existing values are locked; new values may be added (they behave as 'other' states)."),
        ("p", "Column named ranges", "Every table column has a workbook name <prefix>_<Column> (reg_, act_, fin_, wor_, chr_, dev_, risk_, evd_) covering rows "
         "6 to capacity. Formulas use these names, so they are readable in the Name Manager (Formulas > Name Manager)."),
        ("p", "Extending capacity", "Unprotect the sheet, select the last table row, copy it down (formulas follow), resize the table (Table Design > Resize), "
         "and edit the matching names in Name Manager to the new last row. Dashboard project table shows the first 15 projects; WOR report shows 10 cards."),
        ("p", "IDs", "IDs are derived from row position. Never sort or delete rows in the data sheets; set Status = Cancelled instead of deleting. Filtering is safe."),
        ("p", "As-Of date", "Master Data C6 (blank = TODAY()). Excel recalculates TODAY() on open; set a fixed date to reproduce a past WOR view."),
        ("p", "Rebuild", "The workbook is generated by ci_management_system/build_workbook.py (Python/openpyxl) in the repository. Structural changes should be "
         "made in the script and regenerated; data must then be copied into the new file (paste values into input columns only)."),
        ("p", "Compatibility", "Desktop Excel 2019 / Microsoft 365 (uses MAXIFS/MINIFS). No macros, no external links, no Power Query. Excel for the web supports "
         "the formulas and dropdowns; printing layouts are designed for desktop Excel."),
    ]
    return blocks


def build_qa(ws, demo, qa):
    ws.sheet_view.showGridLines = False
    for k, v in dict(A=2, B=8, C=34, D=58, E=12, F=50).items():
        ws.column_dimensions[k].width = v
    banner(ws, "14  QUALITY ASSURANCE REPORT",
           "Automated tests run against this build (LibreOffice headless recalculation + independent Python recomputation on the "
           "demonstration dataset). Live integrity checks are on the dashboard (rows 59-72).", 6, demo)
    r = 5
    if not qa:
        ws["B5"] = "QA has not been run for this build."
        return
    ws["B5"] = f"Build: {qa.get('build')}   |   Tests: {qa.get('total')}   |   Passed: {qa.get('passed')}   |   Failed: {qa.get('failed')}"
    ws.merge_cells("B5:F5")
    ws["B5"].font = F(10, True, NAVY)
    r = 7
    for j, h in enumerate(["#", "Test area", "Test performed", "Result", "Detail"]):
        c = ws.cell(r, 2 + j, h)
        c.font = F(9, True, WHITE)
        c.fill = fill(NAVY)
        c.alignment = CENTER
    r += 1
    for i, t in enumerate(qa["tests"]):
        vals = [i + 1, t["area"], t["test"], t["result"], t.get("detail", "")]
        for j, v in enumerate(vals):
            c = ws.cell(r, 2 + j, v)
            c.font = F(8, j == 3)
            c.alignment = WRAP_TOP
            c.border = BORDER
        ws.row_dimensions[r].height = max(15, 12 * (1 + len(str(vals[4])) // 55))
        r += 1
    cf_text(ws, f"E8:E{r}", "PASS", G_FILL, G_FONT)
    cf_text(ws, f"E8:E{r}", "FAIL", R_FILL, R_FONT)
    r += 1
    section(ws, f"B{r}:F{r}", "KNOWN LIMITATIONS")
    r += 1
    for lim in qa.get("limitations", []):
        ws.cell(r, 3, lim).font = F(9)
        ws.merge_cells(f"C{r}:F{r}")
        ws.cell(r, 3).alignment = WRAP_TOP
        ws.row_dimensions[r].height = 28
        r += 1


# ============================================================================= MAIN BUILD
def build(demo, qa=None, path=None):
    data = demo_data() if demo else prod_data()
    wb = openpyxl.Workbook()
    try:
        wb._named_styles["Normal"].font = Font(name=FONT, size=10)
    except Exception:
        pass
    order = [S_DASH, S_REG, S_CH, S_CHR, S_ACT, S_FIN, S_WOR, S_WORU, S_DEV, S_DEVR, S_TOOLS, S_RISK, S_EVD, S_MD,
             S_GUIDE, S_DD, S_QA]
    wss = {}
    for i, n in enumerate(order):
        wss[n] = wb.active if i == 0 else wb.create_sheet()
        wss[n].title = n
    tabcol = {S_DASH: NAVY, S_WOR: NAVY, S_CH: NAVY, S_REG: NAVY2, S_ACT: NAVY2, S_FIN: NAVY2, S_WORU: NAVY2,
              S_CHR: NAVY2, S_DEV: STEEL, S_DEVR: STEEL, S_TOOLS: STEEL, S_RISK: STEEL, S_EVD: STEEL, S_MD: GRAY_D,
              S_GUIDE: GRAY_M, S_DD: GRAY_M, S_QA: GRAY_M}
    for n, c in tabcol.items():
        wss[n].sheet_properties.tabColor = c

    if demo:
        # demo uses its own prefix
        pass
    build_master(wb, wss[S_MD], demo, data["people"])
    if demo:
        wss[S_MD]["C8"] = "DEMO-"

    # register
    ws = wss[S_REG]
    cols = register_cols()
    n = len(cols)
    banner(ws, "02  PROJECT REGISTER", "One row per project. Enter a title to create the Project ID. One primary owner per project; "
           "co-owner and team members support. Gray columns are calculated from the action, savings, risk and update registers.", 12, demo)
    legend_row(ws, 4, 12)
    L, end = build_table(wb, ws, "tblProjects", "reg", cols, REG_CAP, data["projects"], freeze="C6")
    cf_rag(ws, f"{L['RAG']}6:{L['RAG']}{end}")
    cf_na(ws, f"{L['KPIWeek']}6:{L['WorSeq']}{end}")
    cf_formula(ws, f"{L['OwnerChk']}6:{L['OwnerChk']}{end}", f'AND({L["OwnerChk"]}6<>"",{L["OwnerChk"]}6<>"OK")', A_FILL, A_FONT, True)
    cf_formula(ws, f"{L['Status']}6:{L['Status']}{end}", f'{L["Status"]}6="Blocked"', R_FILL, R_FONT, True)
    cf_formula(ws, f"{L['Overdue']}6:{L['Overdue']}{end}", f'AND(ISNUMBER({L["Overdue"]}6),{L["Overdue"]}6>0)', R_FILL, R_FONT, True)
    cf_formula(ws, f"{L['Leader']}6:{L['Leader']}{end}", f'AND({L["ID"]}6<>"",{L["Leader"]}6="")', A_FILL)
    cf_formula(ws, f"{L['Baseline']}6:{L['Target']}{end}", f'AND($A6<>"",{L["Baseline"]}6="")', A_FILL)
    ws.conditional_formatting.add(f"{L['ActPct']}6:{L['ActPct']}{end}", DataBarRule(start_type="num", start_value=0, end_type="num",
                                                                                      end_value=1, color="8EA9DB", showValue=True))

    # actions
    ws = wss[S_ACT]
    cols = action_cols()
    banner(ws, "04  ACTION TRACKER", "One row per action with ONE accountable owner. Completion credit requires Status = Complete AND a "
           "Completed Date. Overdue and days remaining are date-aware (As-Of date in Master Data).", 12, demo)
    legend_row(ws, 4, 12)
    L, end = build_table(wb, ws, "tblActions", "act", cols, ACT_CAP, data["actions"], freeze="E6")
    rng = f"{L['Flag']}6:{L['Flag']}{end}"
    cf_text(ws, rng, "OVERDUE", R_FILL, R_FONT)
    cf_text(ws, rng, "DUE SOON", A_FILL, A_FONT)
    cf_text(ws, rng, "NO DUE DATE", A_FILL, A_FONT)
    cf_formula(ws, rng, f'LEFT({L["Flag"]}6,6)="Closed"', G_FILL, G_FONT)
    cf_text(ws, f"{L['EvCheck']}6:{L['EvCheck']}{end}", "MISSING EVIDENCE", R_FILL, R_FONT)
    cf_text(ws, f"{L['EvCheck']}6:{L['EvCheck']}{end}", "Verified", G_FILL, G_FONT)
    cf_formula(ws, f"{L['DCheck']}6:{L['DCheck']}{end}", f'AND({L["DCheck"]}6<>"",{L["DCheck"]}6<>"OK")', A_FILL, A_FONT, True)
    cf_formula(ws, f"{L['Desc']}6:{L['Desc']}{end}", f'{L["Flag"]}6="OVERDUE"', R_FILL)
    cf_text(ws, f"{L['Status']}6:{L['Status']}{end}", "Blocked", R_FILL, R_FONT)

    # finance
    ws = wss[S_FIN]
    cols = finance_cols()
    banner(ws, "05  FINANCIAL SAVINGS TRACKER", "One row per benefit line. Hard savings, cost avoidance and productivity capacity are tracked "
           "separately. Only Finance-Approved realized amounts count as VERIFIED. Recovered minutes are not automatic cash. Blank = missing "
           "(flagged), never zero.", 12, demo, note=True)
    legend_row(ws, 4, 12, "Finance-only columns: Finance Status, Reviewer, Review Date, Realized Savings, Finance Comments.")
    L, end = build_table(wb, ws, "tblSavings", "fin", cols, FIN_CAP, data["finance"], freeze="D6")
    rs = f"{L['Status']}6:{L['Status']}{end}"
    cf_text(ws, rs, "Approved", G_FILL, G_FONT)
    cf_text(ws, rs, "Rejected", R_FILL, R_FONT)
    cf_text(ws, rs, "Under Review", A_FILL, A_FONT)
    cf_formula(ws, f"{L['Dbl']}6:{L['Dbl']}{end}", f'LEFT({L["Dbl"]}6,4)="OVER"', R_FILL, R_FONT, True)
    cf_formula(ws, f"{L['Missing']}6:{L['Missing']}{end}", f'LEN({L["Missing"]}6)>0', A_FILL, A_FONT)
    cf_formula(ws, f"{L['VolChk']}6:{L['VolChk']}{end}", f'LEFT({L["VolChk"]}6,6)="Volume"', A_FILL, A_FONT)
    cf_na(ws, f"{L['VolChk']}6:{L['MF']}{end}")
    for k in ("Status", "Reviewer", "RevDate", "Realized", "FinComments"):
        c = ws[f"{L[k]}5"]
        c.fill = fill("375623")

    # WOR updates
    ws = wss[S_WORU]
    cols = woru_cols()
    banner(ws, "06A  WOR WEEKLY UPDATES", "One row per project per week (week-ending date). The WOR report and the register's 'Latest KPI "
           "Actual' read the latest row on or before the As-Of date.", 12, demo)
    legend_row(ws, 4, 12)
    L, end = build_table(wb, ws, "tblWORUpdates", "wor", cols, WORU_CAP, data["woru"], freeze="D6")
    cf_text(ws, f"{L['Dup']}6:{L['Dup']}{end}", "DUPLICATE week", R_FILL, R_FONT)

    # charter records
    ws = wss[S_CHR]
    cols = charter_cols()
    banner(ws, "03A  CHARTER RECORDS", "Narrative content for each project's charter (one row per project). View and print on "
           "'03 Project Charter'.", 12, demo)
    legend_row(ws, 4, 12)
    L, end = build_table(wb, ws, "tblCharters", "chr", cols, CHR_CAP, data["charters"], freeze="C6")
    cf_text(ws, f"{L['Dup']}6:{L['Dup']}{end}", "DUPLICATE", R_FILL, R_FONT)
    ws.conditional_formatting.add(f"{L['Complete']}6:{L['Complete']}{end}", DataBarRule(start_type="num", start_value=0,
                                                                                          end_type="num", end_value=1, color="8EA9DB", showValue=True))

    # development
    ws = wss[S_DEV]
    cols = dev_cols()
    banner(ws, "07  SUPERVISOR DEVELOPMENT", "One row per review. Score 1-4 per the rubric on '07A Rubric and Summary'. Development "
           "focus - not a ranking, and savings dollars are not part of the assessment.", 12, demo)
    legend_row(ws, 4, 12)
    L, end = build_table(wb, ws, "tblDevelopment", "dev", cols, DEV_CAP, data["dev"], freeze="C6")
    cf_text(ws, f"{L['RevFlag']}6:{L['RevFlag']}{end}", "REVIEW OVERDUE", R_FILL, R_FONT)
    cf_na(ws, f"{L['Avg']}6:{L['Level']}{end}")
    sc0 = L[COMPETENCIES[0][0]]
    sc1 = L[COMPETENCIES[-1][0]]
    ws.conditional_formatting.add(f"{sc0}6:{sc1}{end}", CellIsRule(operator="equal", formula=["1"], fill=fill(R_FILL)))
    ws.conditional_formatting.add(f"{sc0}6:{sc1}{end}", CellIsRule(operator="equal", formula=["4"], fill=fill(G_FILL)))

    # risks
    ws = wss[S_RISK]
    cols = risk_cols()
    banner(ws, "09  RISKS AND ESCALATIONS", "Record blockers, risks and support needs. Items needing leadership action are highlighted "
           "automatically.", 12, demo)
    legend_row(ws, 4, 12)
    L, end = build_table(wb, ws, "tblRisks", "risk", cols, RISK_CAP, data["risks"], freeze="D6")
    cf_text(ws, f"{L['Lead']}6:{L['Lead']}{end}", "LEADERSHIP ACTION", R_FILL, R_FONT)
    cf_text(ws, f"{L['Past']}6:{L['Past']}{end}", "PAST TARGET", A_FILL, A_FONT)
    cf_text(ws, f"{L['Impact']}6:{L['Impact']}{end}", "High", R_FILL, R_FONT)
    cf_formula(ws, f"{L['Desc']}6:{L['Desc']}{end}", f'{L["Lead"]}6="LEADERSHIP ACTION"', R_FILL)

    # evidence
    ws = wss[S_EVD]
    cols = evidence_cols()
    banner(ws, "10  EVIDENCE REGISTER", "Excel stores LINKS, not files. Save evidence in the approved SharePoint/Teams folder, copy the "
           "link here. Viewers need folder permission; links do not grant access.", 12, demo)
    ws["A4"] = '="Approved evidence folder: "&IF(EvidenceRoot="","NOT CONFIGURED - CI Manager to set it in 11 Master Data (C13)",EvidenceRoot)'
    ws.merge_cells("A4:L4")
    ws["A4"].font = F(9, True, NAVY)
    ws["A4"].fill = fill(GRAY_XL)
    L, end = build_table(wb, ws, "tblEvidence", "evd", cols, EVD_CAP, data["evidence"], freeze="D6")
    cf_formula(ws, f"{L['LinkChk']}6:{L['LinkChk']}{end}", f'OR({L["LinkChk"]}6="NO LINK",{L["LinkChk"]}6="Check format")', A_FILL, A_FONT, True)
    cf_formula(ws, f"{L['ActChk']}6:{L['ActChk']}{end}", f'AND({L["ActChk"]}6<>"",{L["ActChk"]}6<>"OK")', A_FILL, A_FONT, True)
    cf_text(ws, f"{L['Verif']}6:{L['Verif']}{end}", "Verified", G_FILL, G_FONT)
    cf_text(ws, f"{L['Verif']}6:{L['Verif']}{end}", "Rejected", R_FILL, R_FONT)
    for r in range(6, end + 1):
        ws[f"{L['Open']}{r}"].font = Font(name=FONT, size=10, color="0563C1", underline="single")

    # dynamic lists for project/action IDs
    add_name(wb, "lst_ProjectID", f"OFFSET({q(S_REG)}!$A$6,0,0,MAX(1,COUNTIF(reg_ID,\"?*\")),1)")
    add_name(wb, "lst_ActionID", f"OFFSET({q(S_ACT)}!$A$6,0,0,MAX(1,COUNTIF(act_ID,\"?*\")),1)")

    build_dashboard(wb, wss[S_DASH], demo)
    build_charter(wb, wss[S_CH], demo)
    build_wor(wb, wss[S_WOR], demo, S_DASH)
    build_rubric(wb, wss[S_DEVR], demo)
    build_tools(wb, wss[S_TOOLS], demo)
    write_doc_sheet(wss[S_GUIDE], demo, "12  USER GUIDE AND SUPERVISOR QUICK-START",
                    "Plain-language instructions for supervisors, the CI Manager and Finance.", GUIDE)
    write_doc_sheet(wss[S_DD], demo, "13  DATA DICTIONARY AND TECHNICAL NOTES",
                    "Definition and calculation of every KPI, data states, and how to maintain the workbook.", dd_blocks())
    build_qa(wss[S_QA], demo, qa)

    for n in order:
        protect(wss[n])
    wb.calculation.fullCalcOnLoad = True
    wb.active = 0
    path = path or os.path.join(OUT, "CHG_Duncanville_CI_System_DEMO_TRAINING.xlsx" if demo else
                                "CHG_Duncanville_CI_Management_System.xlsx")
    wb.save(path)
    return path


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--qa", default=None)
    ap.add_argument("--only", choices=["prod", "demo"], default=None)
    a = ap.parse_args()
    qa = json.load(open(a.qa)) if a.qa and os.path.exists(a.qa) else None
    os.makedirs(OUT, exist_ok=True)
    for demo in (False, True):
        if a.only == ("demo" if not demo else "prod"):
            continue
        qad = None
        if qa:
            qad = qa.get("demo" if demo else "prod") or qa.get("combined")
        print(build(demo, qad))
