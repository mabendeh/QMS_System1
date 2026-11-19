#!/usr/bin/env python3
"""
Automated QA for the CI Management System workbooks.

1. Recalculates both workbooks headlessly in LibreOffice (formula errors must be zero).
2. Structural checks: sheets, tables, named ranges, data validation, protection, input/formula locking, hyperlinks.
3. Production content checks: only the confirmed assignments are present; unknown fields blank; dashboard reconciles.
4. Independent Python recomputation of every key KPI from the demo dataset, compared with the Excel results.
5. Scenario tests on temporary copies (WOR filters, status change -> dashboard, unapproved savings exclusion).

Writes qa_results.json (consumed by build_workbook.py --qa to populate sheet '14 QA Report').
Usage: python3 test_workbook.py [--recalc /path/to/recalc.py]
"""
import argparse
import datetime as dt
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import warnings

import openpyxl

warnings.filterwarnings("ignore")
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import build_workbook as B  # noqa: E402

PROD = os.path.join(B.OUT, "CHG_Duncanville_CI_Management_System.xlsx")
DEMO = os.path.join(B.OUT, "CHG_Duncanville_CI_System_DEMO_TRAINING.xlsx")
DEFAULT_RECALC = "/root/.claude/skills/synced/88e3ede0-081c-403b-9e1e-7894cee8e53d_e1d1df58-e696-4722-8700-f2c860ff82c4/xlsx/scripts/recalc.py"
ASOF = dt.datetime(2026, 11, 19)

RESULTS = []


def rec(area, test, ok, detail=""):
    RESULTS.append(dict(area=area, test=test, result="PASS" if ok else "FAIL", detail=str(detail)[:600]))
    print(("PASS " if ok else "FAIL ") + area + " | " + test + ("" if ok else "  -> " + str(detail)[:300]))


def recalc(path, recalc_py):
    out = subprocess.run([sys.executable, recalc_py, path, "300"], capture_output=True, text=True)
    txt = out.stdout[out.stdout.find("{"):]
    return json.loads(txt)


def close(a, b, tol=0.01):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return a == b


def letters(cols):
    return {c["key"]: openpyxl.utils.get_column_letter(i + 1) for i, c in enumerate(cols)}


def table_rows(ws, cols, n):
    L = letters(cols)
    out = []
    for i in range(n):
        r = B.FIRST + i
        out.append({k: ws[f"{L[k]}{r}"].value for k in L})
    return out


# ----------------------------------------------------------------------------- structural
def structural(path, label):
    wb = openpyxl.load_workbook(path)
    expected = [B.S_DASH, B.S_REG, B.S_CH, B.S_CHR, B.S_ACT, B.S_FIN, B.S_WOR, B.S_WORU, B.S_DEV, B.S_DEVR, B.S_TOOLS,
                B.S_RISK, B.S_EVD, B.S_MD, B.S_GUIDE, B.S_DD, B.S_QA]
    rec(f"Structure ({label})", "All 17 worksheets present in the specified order (01-14 plus 03A/06A/07A support sheets)",
        wb.sheetnames == expected, wb.sheetnames)
    tables = {t for ws in wb.worksheets for t in ws.tables}
    need = {"tblProjects", "tblActions", "tblSavings", "tblWORUpdates", "tblCharters", "tblDevelopment", "tblRisks",
            "tblEvidence", "tblPeople"}
    rec(f"Structure ({label})", "Excel tables exist for every register (filters/structured layout)", need <= tables, sorted(tables))
    names = set(wb.defined_names.keys())
    # every formula name reference resolves
    ref_names = set()
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    for m in re.findall(r"\b((?:reg|act|fin|wor|chr|dev|risk|evd|md|lst|K|CH|WOR)_[A-Za-z0-9]+)\b", c.value):
                        ref_names.add(m)
    for m in ("AsOfDate", "IDPrefix", "OpWeeks", "RedOverdueThr", "DueSoonDays", "VolTol", "EvidenceRoot"):
        ref_names.add(m)
    missing = sorted(n for n in ref_names if n not in names)
    rec(f"Structure ({label})", f"All {len(ref_names)} named ranges used in formulas are defined", not missing, missing)
    dvn = set()
    dv_total = 0
    for ws in wb.worksheets:
        for dv in ws.data_validations.dataValidation:
            dv_total += 1
            if dv.type == "list" and dv.formula1 and dv.formula1.startswith("="):
                dvn.add(dv.formula1[1:])
    bad = sorted(n for n in dvn if n not in names)
    rec(f"Dropdowns ({label})", f"{dv_total} data-validation rules; every list dropdown points to a defined dynamic list",
        not bad and dv_total > 50, bad)
    required_lists = ["lst_People", "lst_Dept", "lst_Shift", "lst_Equipment", "lst_Stage", "lst_ActStatus", "lst_Priority",
                      "lst_FinStatus", "lst_Category", "lst_ProjStatus", "lst_ProjectID", "lst_ActionID"]
    rec(f"Dropdowns ({label})", "Required editable lists exist (supervisors, departments, shifts, equipment, stages, statuses, "
        "priorities, Finance statuses, categories, IDs)", all(n in names for n in required_lists),
        [n for n in required_lists if n not in names])
    dyn = all("OFFSET(" in wb.defined_names[n].attr_text for n in required_lists if n in names)
    rec(f"Dropdowns ({label})", "Dropdown lists are dynamic (OFFSET/COUNTA) so new supervisors/values need no formula edits", dyn)
    prot = [ws.title for ws in wb.worksheets if not ws.protection.sheet]
    rec(f"Protection ({label})", "Sheet protection enabled on every sheet (formula cells locked, filters allowed)", not prot, prot)
    # locking: formulas locked, inputs unlocked (sample all table sheets)
    bad_lock = []
    n_f = n_i = 0
    for ws in wb.worksheets:
        if not ws.tables:
            continue
        for row in ws.iter_rows(min_row=6, max_row=40):
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    n_f += 1
                    if not c.protection.locked:
                        bad_lock.append(f"{ws.title}!{c.coordinate} formula unlocked")
                elif c.fill and c.fill.start_color and c.fill.start_color.rgb in ("00FFFBE6", "FFFFFBE6"):
                    n_i += 1
                    if c.protection.locked:
                        bad_lock.append(f"{ws.title}!{c.coordinate} input locked")
    rec(f"Protection ({label})", f"Formula cells locked ({n_f} checked) and yellow input cells unlocked ({n_i} checked)",
        not bad_lock and n_f > 100 and n_i > 100, bad_lock[:10])
    frz = [ws.title for ws in wb.worksheets if ws.tables and ws.title != B.S_MD and not ws.freeze_panes]
    rec(f"Usability ({label})", "Frozen header rows on every register", not frz, frz)
    cf_needed = {B.S_ACT: "OVERDUE", B.S_REG: "Blocked", B.S_EVD: "NO LINK", B.S_RISK: "LEADERSHIP ACTION"}
    cfmiss = []
    for s, token in cf_needed.items():
        txt = " ".join(" ".join(r.formula or []) for cf in wb[s].conditional_formatting for r in cf.rules)
        if token not in txt:
            cfmiss.append(s)
    act_txt = " ".join(" ".join(r.formula or []) for cf in wb[B.S_ACT].conditional_formatting for r in cf.rules)
    rec(f"Formatting ({label})", "Conditional formatting present for overdue actions, blocked projects, missing evidence/links and "
        "leadership escalations", not cfmiss and "MISSING EVIDENCE" in act_txt, cfmiss)
    d = wb[B.S_DASH]
    ok = d.page_setup.fitToWidth == 1 and d.page_setup.fitToHeight == 1 and d.page_setup.orientation == "landscape"
    rec(f"Printing ({label})", "Dashboard WOR view prints on one landscape page (fit 1x1)", ok)
    ch = wb[B.S_CH]
    rec(f"Printing ({label})", "Project charter prints on one portrait page (fit 1x1)",
        ch.page_setup.fitToHeight == 1 and ch.page_setup.orientation == "portrait")
    w = wb[B.S_WOR]
    rec(f"Printing ({label})", "WOR report: landscape, 1 page wide, title rows repeat, page break every 5 projects",
        w.page_setup.orientation == "landscape" and w.print_title_rows and len(w.row_breaks.brk) >= 1)
    # hyperlinks
    nav = [d.cell(4, c).value for c in range(2, 16)]
    targets = re.findall(r"#'([^']+)'!A1", " ".join(v for v in nav if v))
    rec(f"Hyperlinks ({label})", f"Dashboard navigation: {len(targets)} internal links all target existing sheets",
        len(targets) == 14 and all(t in wb.sheetnames for t in targets), targets)
    e = wb[B.S_EVD]
    L = letters(B.evidence_cols())
    hf = e[f"{L['Open']}6"].value
    rec(f"Hyperlinks ({label})", "Evidence Register 'Open Link' column uses HYPERLINK() on the entered SharePoint/Teams URL",
        isinstance(hf, str) and hf.startswith("=IF(") and "HYPERLINK(" in hf, hf)
    charts = sum(len(ws._charts) for ws in wb.worksheets)
    rec(f"Dashboard ({label})", f"Trend charts present ({charts}: completion, overdue, savings, Pareto)", charts >= 4)
    # no manually typed dashboard metrics
    typed = []
    for coord in ["B7", "D7", "F7", "H7", "J7", "L7", "N7", "P7", "B12", "D12", "F12", "H12", "J12", "L12", "N12", "P12"]:
        v = d[coord].value
        if not (isinstance(v, str) and v.startswith("=")):
            typed.append(coord)
    for r in range(17, 32):
        for c in range(2, 17):
            v = d.cell(r, c).value
            if not (isinstance(v, str) and v.startswith("=")):
                typed.append(d.cell(r, c).coordinate)
    rec(f"Dashboard ({label})", "Every KPI card and project-status cell is a formula (no typed metrics)", not typed, typed[:10])
    rec(f"Branding ({label})", "No logo image embedded (official logo not supplied - placeholder text only)",
        all(len(ws._images) == 0 for ws in wb.worksheets))
    fonts = set()
    for ws in (wb[B.S_DASH], wb[B.S_REG], wb[B.S_WOR]):
        for row in ws.iter_rows(min_row=1, max_row=30):
            for c in row:
                if c.value is not None:
                    fonts.add(c.font.name)
    rec(f"Formatting ({label})", "Consistent typography (Arial) on dashboard, register and WOR report", fonts == {"Arial"}, fonts)
    macros = any(n.endswith(".bin") for n in __import__("zipfile").ZipFile(path).namelist())
    ext = "externalLink" in " ".join(__import__("zipfile").ZipFile(path).namelist())
    rec(f"Compatibility ({label})", "No macros (VBA) and no external workbook links", not macros and not ext)


# ----------------------------------------------------------------------------- production content
def production_content(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    reg = wb[B.S_REG]
    rows = table_rows(reg, B.register_cols(), 10)
    filled = [r for r in rows if r["ID"]]
    rec("Production data", "Exactly 5 register rows: 3 confirmed projects + Maintenance and Warehouse placeholders",
        len(filled) == 5, [r["ID"] for r in filled])
    exp = [("SAP Confirmation Process", "Eric Lujan", "Manuela Garcia"), ("Changeover Standard", "Jhirmaul Carter", "Stacy Jackson"),
           ("Waste Reduction", "Phillip Harris", "Ivan Lopez")]
    ok = all(filled[i]["Title"] == t and filled[i]["Leader"] == l and filled[i]["CoOwner"] == c for i, (t, l, c) in enumerate(exp))
    rec("Production data", "Confirmed owners preloaded exactly as assigned (leader + co-owner)", ok)
    rec("Production data", "Placeholders for Maintenance / Warehouse have NO owner and status Pending Assignment",
        all(filled[i]["Leader"] is None and filled[i]["Status"] == "Pending Assignment" for i in (3, 4)))
    blanks = ["Start", "TargetDate", "CompDate", "Baseline", "Target", "KPI", "Sponsor", "Shift", "Area", "Equipment"]
    inv = [(r["ID"], k) for r in filled for k in blanks if r[k] is not None]
    rec("Production data", "No invented dates, baselines, targets, KPIs, sponsors or shifts (all blank)", not inv, inv)
    rec("Production data", "Department left blank for the 3 projects where it was not confirmed",
        all(filled[i]["Dept"] is None for i in range(3)))
    for s, cols, nm in [(B.S_ACT, B.action_cols(), "actions"), (B.S_FIN, B.finance_cols(), "savings lines"),
                        (B.S_WORU, B.woru_cols(), "weekly updates"), (B.S_RISK, B.risk_cols(), "risks"),
                        (B.S_EVD, B.evidence_cols(), "evidence records")]:
        rr = table_rows(wb[s], cols, 20)
        rec("Production data", f"No fabricated {nm} (table empty)", all(r["Proj"] is None for r in rr))
    d = wb[B.S_DASH]
    vals = {k: d[k].value for k in ["B7", "D7", "F7", "H7", "J7", "L7", "N7", "P7"]}
    exp = {"B7": 3, "D7": 0, "F7": "N/A", "H7": 0, "J7": 0, "L7": 0, "N7": 5, "P7": 0}
    rec("Production dashboard", "KPI cards: 3 active, 0 complete, action completion N/A (no actions yet), 0 overdue, $0 forecast, "
        "$0 verified, 5 requiring attention, 0 escalations", all(close(vals[k], v) for k, v in exp.items()), vals)
    res = [d.cell(r, 14).value for r in range(61, 73)]
    rec("Production dashboard", "All 12 live reconciliation checks = PASS", all(x == "PASS" for x in res), res)
    rag = [r["RAG"] for r in filled]
    rec("Production dashboard", "RAG: 3 AMBER (target dates not yet set) and 2 PENDING (awaiting assignment)",
        rag == ["AMBER", "AMBER", "AMBER", "PENDING", "PENDING"], rag)
    md = wb[B.S_MD]
    rec("Production data", "Settings: As-Of date blank (uses TODAY()), ID prefix DUN-CI-, 52-week annualization labelled as assumption",
        md["C6"].value is None and md["C8"].value == "DUN-CI-" and md["C9"].value == 52 and "ASSUMPTION" in md["D9"].value)


# ----------------------------------------------------------------------------- independent recomputation (demo)
def expected_demo():
    data = B.demo_data()
    acts = data["actions"]
    for i, a in enumerate(acts):
        a["ID"] = f"A-{i + 1:04d}"
        a["credit"] = 1 if a["Status"] == "Complete" and a.get("CompDate") else 0
        a["overdue"] = (a["credit"] == 0 and a["Status"] != "Cancelled" and a.get("Due") is not None and a["Due"] < ASOF)
    fins = data["finance"]
    for f in fins:
        m = f.get("Method")
        mat = None
        if m == "Waste - Material Loss":
            mat = (f["BRate"] - f["CRate"]) * f["CVol"] * f["UCost"]
        co = None
        if m == "Changeover - Time Reduction":
            co = f["COMin"] * f["COFreq"] * f["RVal"] if f.get("Demo") == "Yes" and f.get("RVal") is not None else 0
        comps = [x for x in (mat, co, f.get("Labor"), f.get("Down"), f.get("Rework")) if x is not None]
        recur = sum(comps) if comps else None
        ppy = 52 / f["PWeeks"] if f.get("PWeeks") else None
        annual = recur * ppy if recur is not None and ppy else None
        gross = None if annual is None and f.get("OneTime") is None else (annual or 0) + (f.get("OneTime") or 0)
        net = gross - (f.get("Impl") or 0) if gross is not None else None
        attr = f.get("Attr") if f.get("Attr") is not None else 1
        f["fcst"] = net * attr if net is not None else 0
        f["annual"] = annual
        f["attr"] = attr
        f["verified"] = f.get("Realized") if f.get("Status") == "Approved" and f.get("Realized") is not None else 0
        caph = f["COMin"] * f["COFreq"] / 60 if f.get("COMin") is not None and f.get("COFreq") is not None else None
        f["caphrs"] = caph * ppy * attr if caph is not None and ppy else 0
    risks = data["risks"]
    for r in risks:
        r["open"] = 0 if r.get("Status") in ("Resolved", "Closed") else 1
    projs = data["projects"]
    E = {}
    for i, p in enumerate(projs):
        pid = f"DEMO-{i + 1:03d}"
        pa = [a for a in acts if a["Proj"] == pid]
        nact = len([a for a in pa if a["Status"] != "Cancelled"])
        ndone = sum(a["credit"] for a in pa)
        ms = [a for a in pa if a["MS"] == "Yes" and a["Status"] != "Cancelled"]
        overdue = sum(1 for a in pa if a["overdue"])
        nrisk = sum(r["open"] for r in risks if r["Proj"] == pid)
        st = p.get("Status")
        td = p.get("TargetDate")
        if st in ("Complete", "Cancelled"):
            rag = "CLOSED"
        elif st == "Pending Assignment":
            rag = "PENDING"
        elif st == "Blocked" or overdue >= 3 or (td is not None and td < ASOF):
            rag = "RED"
        elif overdue > 0 or nrisk > 0 or not p.get("Leader") or td is None or (p.get("Stage") not in (None, "1-Define") and p.get("Baseline") is None):
            rag = "AMBER"
        else:
            rag = "GREEN"
        pf = [f for f in fins if f["Proj"] == pid]
        E[pid] = dict(
            NAct=nact, NDone=ndone, ActPct=(ndone / nact if nact else "N/A"), NMS=len(ms),
            NMSDone=sum(a["credit"] for a in ms), Overdue=overdue, Risks=nrisk, RAG=rag,
            ExpHard=(sum(f["fcst"] for f in pf if f["Type"] == "Hard Savings" and f["Status"] != "Rejected") if pf else "Not quantified"),
            ExpAvoid=(sum(f["fcst"] for f in pf if f["Type"] == "Cost Avoidance" and f["Status"] != "Rejected") if pf else "Not quantified"),
            VerHard=sum(f["verified"] for f in pf if f["Type"] == "Hard Savings"),
        )
    totals = dict(
        active=sum(1 for p in projs if p.get("Status") not in ("Complete", "Cancelled", "Pending Assignment")),
        complete=sum(1 for p in projs if p.get("Status") == "Complete"),
        act_total=sum(1 for a in acts if a["Status"] != "Cancelled"),
        act_done=sum(a["credit"] for a in acts),
        overdue=sum(1 for a in acts if a["overdue"]),
        fcst=sum(f["fcst"] for f in fins if f["Type"] == "Hard Savings" and f["Status"] != "Rejected"),
        verified=sum(f["verified"] for f in fins if f["Type"] == "Hard Savings"),
        avoid_f=sum(f["fcst"] for f in fins if f["Type"] == "Cost Avoidance" and f["Status"] != "Rejected"),
        avoid_v=sum(f["verified"] for f in fins if f["Type"] == "Cost Avoidance"),
        caphrs=sum(f["caphrs"] for f in fins if f["Status"] != "Rejected"),
        attention=sum(1 for v in E.values() if v["RAG"] in ("RED", "AMBER", "PENDING")),
        escal=sum(1 for r in risks if r["open"] and (r.get("Level") == "Plant Leadership" or r.get("Impact") == "High" or
                                                      r.get("Status") == "Escalated" or (r.get("Target") and r["Target"] < ASOF))),
        missing_ev=sum(1 for a in acts if a["credit"] and not a.get("EvidRef") and not any(e.get("Action") == a["ID"] for e in data["evidence"])),
    )
    # trends
    trend = []
    for k in range(12):
        w = ASOF - dt.timedelta(days=7 * (11 - k))
        planned = sum(1 for a in acts if a["Status"] != "Cancelled" and a.get("Due") and a["Due"] <= w)
        comp = sum(1 for a in acts if a["credit"] and a["CompDate"] <= w)
        od = sum(1 for a in acts if a["Status"] != "Cancelled" and a.get("Due") and a["Due"] < w and not (a["credit"] and a["CompDate"] <= w))
        trend.append((planned, comp, od))
    # monthly forecast at current month end
    midx = ASOF.year * 12 + ASOF.month
    mf = 0
    for f in fins:
        if f["Type"] == "Hard Savings" and f["Status"] != "Rejected" and f["annual"] is not None and f.get("Start"):
            si = f["Start"].year * 12 + f["Start"].month
            if midx >= si:
                mf += f["annual"] * f["attr"] / 12 * (midx - si + 1)
    return data, E, totals, trend, fins, mf


def demo_compare(path):
    data, E, T, trend, fins, mf = expected_demo()
    wb = openpyxl.load_workbook(path, data_only=True)
    reg = table_rows(wb[B.S_REG], B.register_cols(), 10)
    keys = ["NAct", "NDone", "ActPct", "NMS", "NMSDone", "Overdue", "Risks", "RAG", "ExpHard", "ExpAvoid", "VerHard"]
    mism = []
    for r in reg:
        if not r["ID"]:
            continue
        for k in keys:
            if not close(r[k], E[r["ID"]][k]):
                mism.append((r["ID"], k, r[k], E[r["ID"]][k]))
    rec("Calculation (demo)", "Register per-project KPIs (actions, completion %, milestones, overdue, risks, RAG, forecast, verified) "
        "match independent Python recomputation for all 6 projects", not mism, mism)
    fr = table_rows(wb[B.S_FIN], B.finance_cols(), len(fins))
    fm = [(r["ID"], r["FcstNum"], f["fcst"], r["Verified"], f["verified"]) for r, f in zip(fr, fins)
          if not (close(r["FcstNum"], f["fcst"]) and close(r["Verified"], f["verified"]))]
    rec("Calculation (demo)", "Savings lines: material, changeover, annualization, gross/net, attribution and verified values match "
        "the documented formulas", not fm, fm)
    rec("Finance integrity (demo)", "Unapproved line with a realized amount entered (Under Review, $9,999) contributes $0 to verified",
        close(fr[6]["Verified"], 0) and fr[6]["Status"] == "Under Review")
    rec("Finance integrity (demo)", "Changeover line NOT demonstrated values $0 cash but reports capacity hours (5.8 h/wk -> 303.3 h/yr)",
        close(fr[2]["CODollar"], 0) and close(fr[2]["CapHrsYr"], 35 * 10 / 60 * 52, 0.1))
    rec("Finance integrity (demo)", "Shared benefit pool attributed 60% + 50% is flagged OVER-ATTRIBUTED (double counting prevented)",
        str(fr[0]["Dbl"]).startswith("OVER") and str(fr[1]["Dbl"]).startswith("OVER"), (fr[0]["Dbl"], fr[1]["Dbl"]))
    rec("Finance integrity (demo)", "Rejected line excluded from forecast KPI and flagged for missing data source",
        fr[4]["Status"] == "Rejected" and "Data source" in str(fr[4]["Missing"]))
    rec("Finance integrity (demo)", "Volume shift above 20% tolerance is flagged / within tolerance reads 'Comparable'",
        fr[0]["VolChk"] == "Comparable", fr[0]["VolChk"])
    d = wb[B.S_DASH]
    got = dict(active=d["B7"].value, complete=d["D7"].value, overdue=d["H7"].value, fcst=d["J7"].value, verified=d["L7"].value,
               attention=d["N7"].value, escal=d["P7"].value, avoid_f=d["F12"].value, avoid_v=d["H12"].value, caphrs=d["J12"].value,
               missing_ev=d["L12"].value, act_total=d["S6"].value, act_done=d["S7"].value)
    mm = {k: (got[k], T[k]) for k in got if not close(got[k], T[k])}
    rec("Dashboard reconciliation (demo)", "All KPI cards equal totals recomputed independently from the source records", not mm, mm)
    rec("Dashboard reconciliation (demo)", "Action completion card = credited / non-cancelled actions",
        close(d["F7"].value, T["act_done"] / T["act_total"]), (d["F7"].value, T["act_done"], T["act_total"]))
    res = {d.cell(r, 9).value: d.cell(r, 14).value for r in range(61, 73)}
    nonpass = {k: v for k, v in res.items() if v != "PASS"}
    rec("Dashboard reconciliation (demo)", "Live integrity panel passes every check except the intentionally seeded over-attribution",
        list(nonpass) == ["Over-attributed benefit pools (must be 0)"], nonpass)
    dept_total = d["D47"].value
    rec("Dashboard reconciliation (demo)", "Department and shift breakdowns sum to total projects, actions and savings",
        dept_total == 6 and d["D57"].value == 6 and d["H47"].value == T["act_total"] and close(d["M47"].value, T["fcst"]) and
        close(d["M57"].value, T["fcst"]) and d["H57"].value == T["act_total"],
        (dept_total, d["D57"].value, d["H47"].value, d["M47"].value))
    tg = [(d.cell(r, 3).value, d.cell(r, 4).value, d.cell(r, 5).value) for r in range(76, 88)]
    rec("Trends (demo)", "12-week planned / completed / overdue trend matches date-based recomputation", tg == trend,
        list(zip(tg, trend))[:4])
    rec("Trends (demo)", "Cumulative forecast run-rate for the current month matches recomputation", close(d["I87"].value, mf),
        (d["I87"].value, mf))
    # action tracker flags
    at = table_rows(wb[B.S_ACT], B.action_cols(), 20)
    rec("Actions (demo)", "'Complete' without completed date is NOT credited and is flagged 'Enter completed date'",
        at[11]["Credit"] == 0 and at[11]["DCheck"] == "Enter completed date", (at[11]["Credit"], at[11]["DCheck"]))
    rec("Actions (demo)", "Days remaining = due - As-Of date; overdue flagged; on-time/late closure classified",
        at[3]["Days"] == -6 and at[3]["Flag"] == "OVERDUE" and at[1]["Flag"] == "Closed Late" and at[0]["Flag"] == "Closed On Time"
        and at[4]["Flag"] == "DUE SOON", [(a["Days"], a["Flag"]) for a in at[:5]])
    rec("Actions (demo)", "Missing evidence flagged on credited actions without evidence; evidence-linked action shows Verified",
        at[0]["EvCheck"] == "Verified" and at[2]["EvCheck"] == "MISSING EVIDENCE" and at[1]["EvCheck"] == "Not yet verified",
        [a["EvCheck"] for a in at[:3]])
    rec("Actions (demo)", "Action with no due date flagged NO DUE DATE", at[18]["Flag"] == "NO DUE DATE")
    ev = table_rows(wb[B.S_EVD], B.evidence_cols(), 3)
    rec("Evidence (demo)", "Link check: https link OK, malformed link flagged; Action ID on another project flagged",
        ev[0]["LinkChk"] == "Format OK" and ev[1]["LinkChk"] == "Check format" and ev[2]["ActChk"] == "Action is on another project",
        [(e["LinkChk"], e["ActChk"]) for e in ev])
    rk = table_rows(wb[B.S_RISK], B.risk_cols(), 3)
    rec("Risks (demo)", "Escalated High-impact blocker flagged LEADERSHIP ACTION; closed item not flagged; days open computed",
        rk[0]["Lead"] == "LEADERSHIP ACTION" and rk[2]["Lead"] in ("", None) and rk[1]["Lead"] in ("", None) and rk[2]["DaysOpen"] == 19,
        [(r["Lead"], r["DaysOpen"]) for r in rk])
    dv = table_rows(wb[B.S_DEV], B.dev_cols(), 3)
    rec("Development (demo)", "Competency average excludes unscored areas (3.375 with 8 of 9 scored); level band and focus area derived",
        close(dv[1]["Avg"], 27 / 8) and dv[1]["Level"] == "3 - Proficient" and dv[2]["Focus"] == "Problem Definition",
        (dv[1]["Avg"], dv[1]["Level"], dv[2]["Focus"]))
    sm = wb[B.S_DEVR]
    rec("Development (demo)", "Participation summary: latest review average and closure % per supervisor (no savings, no ranking)",
        sm["B21"].value == "Demo Supervisor A" and close(sm["L21"].value, 27 / 8) and close(sm["H21"].value, 2 / 4),
        (sm["B21"].value, sm["L21"].value, sm["H21"].value))
    ch = wb[B.S_CH]
    vals = [ch.cell(r, c).value for r in range(1, 60) for c in range(2, 8)]
    rec("Charter (demo)", "Charter pulls register, charter-record and milestone data for the selected project",
        "DEMO - Line scrap reduction" in vals and "Reduce scrap from 4.0% to 2.5% by 01/29/2027" in vals and
        "Collect 4 weeks scrap data" in vals and "Train all shifts" in vals)
    w = wb[B.S_WOR]
    rec("WOR report (demo)", "Default filters show 5 open projects (completed excluded) with goal, KPI actual and latest update",
        str(w["B8"].value).startswith("DEMO-001") and close(w["J10"].value, 3.1) and w["B12"].value == "Trial complete (demo)"
        and w["K4"].value == "Projects shown: 5", (w["B8"].value, w["J10"].value, w["B12"].value, w["K4"].value))
    rec("WOR report (demo)", "Project with no weekly update shows 'No update recorded' and STALE warning",
        "STALE" in str(w["B27"].value) and w["B26"].value == "No update recorded", (w["B26"].value,))


def scenario(src, recalc_py, edits, checks, name):
    tmpd = tempfile.mkdtemp()
    p = os.path.join(tmpd, "scenario.xlsx")
    wb = openpyxl.load_workbook(src)
    for sheet, coord, val in edits:
        wb[sheet][coord] = val
    wb.save(p)
    r = recalc(p, recalc_py)
    wb2 = openpyxl.load_workbook(p, data_only=True)
    ok, detail = checks(wb2)
    rec("Scenario (demo copy)", name, ok and r.get("total_errors") == 0, detail)
    shutil.rmtree(tmpd, ignore_errors=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--recalc", default=DEFAULT_RECALC)
    a = ap.parse_args()
    for p, lab in ((PROD, "production"), (DEMO, "demo")):
        r = recalc(p, a.recalc)
        rec(f"Formulas ({lab})", f"LibreOffice full recalculation of {r.get('total_formulas')} formulas: zero errors "
            "(no REF, VALUE, NAME, DIV/0 or N/A error values)", r.get("status") == "success" and r.get("total_errors") == 0, r.get("error_summary"))
    structural(PROD, "production")
    structural(DEMO, "demo")
    production_content(PROD)
    demo_compare(DEMO)
    actL = letters(B.action_cols())
    # scenario 1: WOR filter by department + include completed
    scenario(DEMO, a.recalc, [(B.S_WOR, "C4", "Maintenance"), (B.S_WOR, "J4", "Yes")],
             lambda wb: (str(wb[B.S_WOR]["B8"].value).startswith("DEMO-004") and wb[B.S_WOR]["A15"].value in (None, ""),
                         (wb[B.S_WOR]["B8"].value, wb[B.S_WOR]["A15"].value)),
             "WOR filter Department = Maintenance + include completed -> only DEMO-004 shown")
    scenario(DEMO, a.recalc, [(B.S_WOR, "G4", "Demo Supervisor C")],
             lambda wb: (str(wb[B.S_WOR]["B8"].value).startswith("DEMO-002") and str(wb[B.S_WOR]["B15"].value).startswith("DEMO-003")
                         and wb[B.S_WOR]["A22"].value in (None, ""), (wb[B.S_WOR]["B8"].value, wb[B.S_WOR]["B15"].value)),
             "WOR owner filter matches leader OR co-owner (Demo Supervisor C -> DEMO-002 co-owner, DEMO-003 leader)")
    scenario(DEMO, a.recalc, [(B.S_WOR, "E4", "B")],
             lambda wb: (str(wb[B.S_WOR]["B8"].value).startswith("DEMO-002") and str(wb[B.S_WOR]["B15"].value).startswith("DEMO-005"),
                         (wb[B.S_WOR]["B8"].value, wb[B.S_WOR]["B15"].value)),
             "WOR shift filter = B -> DEMO-002 and DEMO-005")
    # scenario 2: completing an overdue action updates everything
    scenario(DEMO, a.recalc, [(B.S_ACT, f"{actL['CompDate']}9", dt.datetime(2026, 11, 18)), (B.S_ACT, f"{actL['Status']}9", "Complete")],
             lambda wb: (wb[B.S_DASH]["S7"].value == 8 and wb[B.S_DASH]["H7"].value == 5 and
                         all(wb[B.S_DASH].cell(r, 14).value == "PASS" for r in (63, 64, 69, 70)),
                         (wb[B.S_DASH]["S7"].value, wb[B.S_DASH]["H7"].value)),
             "Entering a completed date on an overdue action: credited 7->8, overdue 6->5, reconciliation still PASS")
    # scenario 3: Finance approves the under-review line
    finL = letters(B.finance_cols())
    scenario(DEMO, a.recalc, [(B.S_FIN, f"{finL['Status']}12", "Approved"), (B.S_FIN, f"{finL['RevDate']}12", dt.datetime(2026, 11, 18))],
             lambda wb: (B and close(wb[B.S_DASH]["L7"].value, 3900 + 9999), wb[B.S_DASH]["L7"].value),
             "Finance status change to Approved moves the realized amount into verified savings (3,900 -> 13,899)")
    # scenario 4: As-Of date change re-evaluates overdue
    scenario(DEMO, a.recalc, [(B.S_MD, "C6", dt.datetime(2026, 9, 1))],
             lambda wb: (wb[B.S_DASH]["H7"].value == 0, wb[B.S_DASH]["H7"].value),
             "Changing the As-Of date to 09/01/2026 re-evaluates overdue actions (6 -> 0)")
    # scenario 5: adding a new supervisor and project without formula edits
    regL = letters(B.register_cols())
    scenario(DEMO, a.recalc, [(B.S_MD, "B31", "New Supervisor X"), (B.S_REG, f"{regL['Title']}12", "New project"),
                              (B.S_REG, f"{regL['Leader']}12", "New Supervisor X"), (B.S_REG, f"{regL['Status']}12", "Active")],
             lambda wb: (wb[B.S_REG]["A12"].value == "DEMO-007" and wb[B.S_DASH]["B7"].value == 5 and
                         wb[B.S_REG][f"{regL['OwnerChk']}12"].value == "OK" and wb[B.S_DEVR]["B30"].value == "New Supervisor X",
                         (wb[B.S_REG]["A12"].value, wb[B.S_DASH]["B7"].value, wb[B.S_REG][f"{regL['OwnerChk']}12"].value)),
             "Adding a supervisor (Master Data) and a project (Register) works with no formula changes: ID auto-generated, owner valid, "
             "dashboard and development summary update")
    passed = sum(1 for r in RESULTS if r["result"] == "PASS")
    out = dict(combined=dict(build=f"v{B.VERSION} {B.BUILD_DATE.isoformat()}", total=len(RESULTS), passed=passed,
                             failed=len(RESULTS) - passed, tests=RESULTS, limitations=LIMITATIONS))
    with open(os.path.join(HERE, "qa_results.json"), "w") as fh:
        json.dump(out, fh, indent=1, default=str)
    print(f"\n{passed}/{len(RESULTS)} passed")
    return 0 if passed == len(RESULTS) else 1


LIMITATIONS = [
    "Recalculation was verified with LibreOffice; the workbook targets desktop Excel 2019 / Microsoft 365 (MAXIFS/MINIFS). Open once in "
    "Excel and confirm the dashboard before first WOR use - Excel recalculates fully on open.",
    "Sheet protection has no password and is not a security control; Excel has no row-level ownership security. Anyone with edit access "
    "can change any input (use SharePoint version history for audit).",
    "Excel cannot upload, store or secure documents, and there is no live SharePoint synchronization. The Evidence Register stores "
    "hyperlinks only; viewers need permission to the target folder. The approved folder URL must be entered by the CI Manager.",
    "Trends are reconstructed from recorded dates (due/completed/Finance review). If dates are edited retroactively the history changes; "
    "for a frozen record, save a PDF of the dashboard each WOR week.",
    "IDs derive from row position: do not sort or delete rows in the data sheets (use Status = Cancelled). Filtering is safe.",
    "Capacity: 150 projects, 1,000 actions, 300 savings lines, 2,000 weekly updates, 300 risks, 1,000 evidence links, 300 reviews; "
    "dashboard lists the first 15 projects and the WOR report 10 per filter.",
    "Annualization uses 52 operating weeks by default - an assumption pending Finance confirmation (Master Data C9).",
    "Data-sheet print areas are dynamic named ranges (grow with the records entered); some non-Microsoft viewers ignore them.",
    "No official C.H. Guenther logo was supplied, so none is embedded; a labelled placeholder marks where to insert it.",
    "Primary owner for each initial project is the first-named person in the assignment; confirm or swap leader/co-owner in the register.",
]

if __name__ == "__main__":
    sys.exit(main())
