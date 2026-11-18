"""Builds CHG_Bronze_Level_Tracker.xlsx – a simple, one-row-per-requirement Bronze tracker with a dashboard.

Usage: python build_bronze_simple.py <out_dir>
"""
import sys
import random
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.chart import BarChart, DoughnutChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from xlcommon import *  # noqa
from build_master import BRONZE, SUGGEST, PILLARS, LAYER, PURPOSE

S_D, S_T, S_O, S_L, S_H = "Dashboard", "Bronze Tracker", "Owners", "Lists", "How To Use"
NROW, NOWN = 45, 60
HEALTH = {"Done": "Green", "On Track": "Green", "Needs Verification": "Yellow", "At Risk": "Yellow",
          "Overdue": "Red", "Not Started": "Gray", "N/A": "Gray"}


def build(demo, path):
    random.seed(11)
    AS_OF = dt.date(2026, 11, 18)
    wb = Workbook()
    wb.remove(wb.active)
    sh = {n: wb.create_sheet(n) for n in [S_D, S_T, S_O, S_L, S_H]}
    for n, c in {S_D: NAVY, S_T: BURG, S_O: "7F7F7F", S_L: "7F7F7F", S_H: "7F7F7F"}.items():
        sh[n].sheet_properties.tabColor = c
    tag = "DEMONSTRATION DATA – fictitious, for training only" if demo else "C.H. Guenther & Son — Duncanville"

    # ---------------- Lists ----------------
    ws = sh[S_L]
    title_band(ws, "Lists & Settings", "Dropdown values and two settings. Add new values at the bottom of a list.", 12,
               legend=False)
    ws["B5"], ws["C5"] = "Status date (blank = today)", AS_OF if demo else None
    ws["B6"], ws["C6"] = "ƒ Status date used", "=IF(ISNUMBER(C5),C5,TODAY())"
    ws["B7"], ws["C7"] = "'At Risk' warning window (days before due)", 14
    for r in (5, 6, 7):
        ws.cell(r, 2).font = f(9, r == 6)
        c = ws.cell(r, 3)
        c.border = BORDER
        c.font = f(9, False, "000000" if r == 6 else "0000FF")
        c.fill = fill(LGRAY if r == 6 else WHITE)
        c.alignment = Alignment(horizontal="center")
    ws["C5"].number_format = ws["C6"].number_format = DATE
    define(wb, "AsOfB", f"{qs(ws)}!$C$6")
    define(wb, "set_RiskDays", f"{qs(ws)}!$C$7")
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 14
    lists = [("Pillar", [p[0] for p in PILLARS[:9]]), ("Pillar Name", [p[1] for p in PILLARS[:9]]),
             ("Status", ["Not Started", "In Progress", "Complete", "N/A"]), ("Yes / No", ["Yes", "No"]),
             ("% Complete", [0, 0.25, 0.5, 0.75, 1]), ("(spacer)", []), ("(owners)", []),
             ("Layer", [LAYER[p[0]] for p in PILLARS[:9]]), ("Pillar Purpose (GPS poster)", [PURPOSE[p[0]] for p in PILLARS[:9]])]
    for j, (h, vals) in enumerate(lists):
        col = 5 + j
        if h in ("(owners)", "(spacer)"):
            continue
        hdr_row(ws, 9, col, [h])
        ws.column_dimensions[L(col)].width = {"Pillar Name": 30, "Layer": 24, "Pillar Purpose (GPS poster)": 52}.get(h, 13)
        for i in range(20):
            c = ws.cell(10 + i, col, vals[i] if i < len(vals) else None)
            c.border = BORDER
            c.font = f(9)
            if h == "% Complete":
                c.number_format = PCT
    for nm, colL in [("bl_Pillar", "E"), ("bl_Status", "G"), ("bl_YesNo", "H"), ("bl_Pct", "I")]:
        dyn_list(wb, nm, ws, colL, 10, 29)
    define(wb, "bl_PillarCode", f"{qs(ws)}!$E$10:$E$29")
    define(wb, "bl_PillarName", f"{qs(ws)}!$F$10:$F$29")
    define(wb, "bl_Layer", f"{qs(ws)}!$L$10:$L$29")
    ws.cell(30, 6, "Pillar names, order, layers and purposes: CHG GPS poster 'How the 9 GPS (TPM) pillars work together'.").font = f(8, False, GRAY, True)
    # active owner helper list
    hdr_row(ws, 9, 11, ["Owner list (auto)"], calc=True)
    ws.column_dimensions["K"].width = 22
    for k in range(NOWN):
        c = ws.cell(10 + k, 11, f'=IFERROR(INDEX(OW_Name,MATCH({k + 1},OW_Seq,0))&"","")')
        c.font = f(8, False, GRAY)
    dyn_list(wb, "bl_Owners", ws, "K", 10, 9 + NOWN, text_only=True)

    # ---------------- Owners ----------------
    ws = sh[S_O]
    title_band(ws, "Owners", "Type each person once. They appear in every Owner dropdown. Set Inactive to remove from the dropdown.", 8)
    os_ = [dict(key="Name", header="Name", width=26), dict(key="Role", header="Role / Position", width=24),
           dict(key="Dept", header="Department", width=18), dict(key="Email", header="Email", width=28),
           dict(key="Active", header="Active?", width=9, dv="bl_ActiveYN", default="Yes")]
    oc = Cols(os_)
    os_.append(dict(key="Seq", header="#", width=5, calc=True, fmt="0",
                    formula=f'=IF(AND({{Name}}<>"",{{Active}}<>"No"),COUNTIFS(${oc.letter["Name"]}$6:{{Name}},"?*",${oc.letter["Active"]}$6:{{Active}},"<>No"),"")'))
    define(wb, "bl_ActiveYN", "bl_YesNo")
    odata = []
    if demo:
        for i, (role, dept) in enumerate([("Plant Manager", "Plant Leadership"), ("AM Pillar Lead", "Production"),
                                          ("Maintenance Manager", "Maintenance"), ("FSQ Manager", "Quality / FSQ"),
                                          ("EHS Manager", "EHS"), ("Training Coordinator", "HR / Training"),
                                          ("CI Engineer", "Engineering"), ("Supply Chain Manager", "Supply Chain"),
                                          ("QA Auditor (verifier)", "Quality / FSQ")], 1):
            odata.append(dict(Name=f"DEMO Owner {i:02d}", Role=role, Dept=dept, Email=f"demo{i:02d}@example.com", Active="Yes"))
    for c in os_:
        if c.get("dv") == "bl_ActiveYN":
            c["dv"] = "bl_YesNo"
    build_table(ws, wb, Cols(os_), 5, NOWN, "tblOwnersB", odata, name_prefix="OW", freeze_col="C")

    # ---------------- Tracker ----------------
    ws = sh[S_T]
    title_band(ws, "Bronze Level Tracker", tag + "  ·  One row per official Bronze requirement. Fill in the white "
               "columns; gray columns calculate automatically.", 22)
    ts = [
        dict(key="Ref", header="Ref", width=8),
        dict(key="Pillar", header="Pillar", width=7, dv="bl_Pillar"),
        dict(key="Layer", header="GPS Layer", width=13, calc=True, wrap=True,
             formula='=IFERROR(INDEX(bl_Layer,MATCH({Pillar},bl_PillarCode,0))&"","")'),
        dict(key="Req", header="Bronze Requirement", width=40, wrap=True),
        dict(key="Evidence", header="Evidence Needed (suggested – confirm with CHG)", width=34, wrap=True),
        dict(key="Owner", header="Owner", width=18, dv="bl_Owners"),
        dict(key="Start", header="Start Date", width=10, dv="date", fmt=DATE),
        dict(key="Due", header="Due Date", width=10, dv="date", fmt=DATE),
        dict(key="Status", header="Status", width=12, dv="bl_Status", default="Not Started"),
        dict(key="Pct", header="% Complete", width=9, dv="bl_Pct", fmt=PCT, default=0),
        dict(key="Link", header="Evidence Link (Teams)", width=22),
        dict(key="Open", header="Open", width=7, calc=True, formula='=IF({Link}="","",HYPERLINK({Link},"Open ↗"))'),
        dict(key="Verified", header="Verified?", width=9, dv="bl_YesNo", default="No"),
        dict(key="VerBy", header="Verified By", width=16, dv="bl_Owners"),
        dict(key="Gap", header="Gap / What's Missing", width=28, wrap=True),
        dict(key="Next", header="Next Step", width=28, wrap=True),
        dict(key="Updated", header="Last Updated", width=10, dv="date", fmt=DATE),
        dict(key="Health", header="Health", width=15, calc=True,
             formula=('=IF({Req}="","",IF({Status}="N/A","N/A",IF({Done}=1,"Done",IF({Overdue}=1,"Overdue",'
                      'IF({Status}="Complete","Needs Verification",IF(AND(ISNUMBER({Due}),{Due}-AsOfB<=set_RiskDays,N({Pct})<0.75),"At Risk",'
                      'IF(AND({Status}="Not Started",N({Pct})=0),"Not Started","On Track")))))))')),
        dict(key="Missing", header="Missing Info", width=20, calc=True, wrap=True,
             formula=('=IF(OR({Req}="",{Status}="N/A",{Done}=1),"",IF({Owner}="","Owner; ","")&IF(NOT(ISNUMBER({Due})),"Due date; ","")'
                      '&IF(AND({Status}="Complete",{Link}=""),"Evidence link; ","")&IF(AND({Status}="Complete",{Verified}<>"Yes"),"Verification; ",""))')),
        dict(key="Prog", header="Progress", width=8, calc=True, fmt=PCT,
             formula='=IF(OR({Req}="",{Status}="N/A"),"",IF({Done}=1,1,IF({Status}="Complete",MAX(0.75,N({Pct})),N({Pct}))))'),
        dict(key="Appl", header="Counts?", width=6, calc=True, fmt="0",
             formula='=IF(AND({Req}<>"",{Status}<>"N/A"),1,0)'),
        dict(key="Done", header="Done (1/0)", width=6, calc=True, fmt="0",
             formula='=IF(AND({Appl}=1,{Status}="Complete",{Verified}="Yes",{Link}<>""),1,0)'),
        dict(key="Overdue", header="Overdue (1/0)", width=6, calc=True, fmt="0",
             formula='=IF(AND({Appl}=1,{Done}=0,ISNUMBER({Due})),IF({Due}<AsOfB,1,0),0)'),
        dict(key="Days", header="Days Late", width=6, calc=True, fmt="0",
             formula='=IF({Overdue}=1,AsOfB-{Due},"")'),
        dict(key="Key", header="Attention Key", width=6, calc=True, fmt="0.0",
             formula=('=IF({Health}="Overdue",100000+N({Days}),IF(OR({Health}="At Risk",{Health}="Needs Verification"),'
                      'IF(ISNUMBER({Due}),5000-MIN(4999,MAX(0,{Due}-AsOfB)),100),0))'
                      '+IF(OR({Health}="Overdue",{Health}="At Risk",{Health}="Needs Verification"),ROW()/100000,0)')),
    ]
    rows = []
    for p, reqs in BRONZE.items():
        for i, r in enumerate(reqs, 1):
            rows.append(dict(Ref=f"{p}-{i}", Pillar=p, Req=r, Evidence=SUGGEST[p][i - 1], Status="Not Started",
                             Pct=0, Verified="No"))
    if demo:
        owner = {"AM": 2, "PM": 3, "FSQ": 4, "EHS": 5, "LEAD": 1, "E&T": 6, "VO": 7, "IM": 1, "SC": 8}
        for i, r in enumerate(rows):
            r["Owner"] = f"DEMO Owner {owner[r['Pillar']]:02d}" if i != 4 else None
            r["Start"] = AS_OF - dt.timedelta(days=random.randint(30, 120))
            r["Due"] = AS_OF + dt.timedelta(days=random.randint(-40, 120))
            r["Updated"] = AS_OF - dt.timedelta(days=random.randint(0, 10))
            u = random.random()
            if u < 0.30:
                r.update(Status="Complete", Pct=1, Verified="Yes", VerBy="DEMO Owner 09",
                         Link=f"https://example.sharepoint.com/sites/DunGPS/Shared/{r['Pillar']}/DEMO_{i}.pdf")
            elif u < 0.38:
                r.update(Status="Complete", Pct=1, Verified="No",
                         Link=f"https://example.sharepoint.com/sites/DunGPS/Shared/{r['Pillar']}/DEMO_{i}.pdf")
            elif u < 0.75:
                r.update(Status="In Progress", Pct=random.choice([0.25, 0.5, 0.75]), Gap="DEMO: what is missing",
                         Next="DEMO: next step")
        rows[-1]["Status"] = "N/A"
    first, last = build_table(ws, wb, Cols(ts), 5, NROW, "tblBronzeSimple", rows, name_prefix="BT", freeze_col="E",
                              row_height=30)
    TL = Cols(ts).letter
    text_cf(ws, f"{TL['Health']}{first}:{TL['Health']}{last}", HEALTH)
    text_cf(ws, f"{TL['Status']}{first}:{TL['Status']}{last}", {"Complete": "Green", "In Progress": "Yellow"})
    bar_cf(ws, f"{TL['Prog']}{first}:{TL['Prog']}{last}")
    ws.conditional_formatting.add(f"{TL['Missing']}{first}:{TL['Missing']}{last}",
                                  FormulaRule(formula=[f'LEN(${TL["Missing"]}{first})>0'], font=Font(color="C55A11", bold=True)))
    ws.cell(4, 2, "Done = Status 'Complete' + Evidence Link + Verified? 'Yes'.   Use 'N/A' only when CHG has approved "
                  "that the requirement does not apply.").font = f(8, True, BURG)
    # hide technical helper columns
    for k in ("Appl", "Done", "Overdue", "Days", "Key"):
        ws.column_dimensions[TL[k]].hidden = True

    # ---------------- Dashboard ----------------
    ws = sh[S_D]
    title_band(ws, "Bronze Level Readiness — Dashboard", "", 14, legend=False)
    ws["B3"] = '="CHG Duncanville  ·  Status as of "&TEXT(AsOfB,"dd-mmm-yyyy")' + (
        '&"  ·  DEMONSTRATION DATA"' if demo else '&"  ·  Source: CHG GPS Pillar Bronze Level Requirements (32)"')
    for i, w in enumerate([2, 10, 26, 12, 10, 10, 10, 10, 10, 11, 11, 13, 10, 10], 1):
        ws.column_dimensions[L(i)].width = w
    ws.merge_cells("B5:N5")
    ws["B5"] = ('=IF(AND(SUM(BT_Appl)>0,SUM(BT_Done)=SUM(BT_Appl)),"✔  BRONZE READY – all "&SUM(BT_Appl)&" requirements done and verified",'
                '"NOT YET BRONZE READY  –  "&SUM(BT_Done)&" of "&SUM(BT_Appl)&" requirements done  ·  "&SUM(BT_Overdue)&" overdue")')
    ws["B5"].font = f(13, True, WHITE)
    ws["B5"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[5].height = 28
    ws.conditional_formatting.add("B5:N5", FormulaRule(formula=['LEFT($B$5,1)="✔"'], fill=fill("2E7D32")))
    ws.conditional_formatting.add("B5:N5", FormulaRule(formula=['LEFT($B$5,1)<>"✔"'], fill=fill(BURG)))
    tiles = [
        ("% DONE (VERIFIED)", '=IF(SUM(BT_Appl)=0,0,SUM(BT_Done)/SUM(BT_Appl))', PCT, '=SUM(BT_Done)&" of "&SUM(BT_Appl)&" requirements"', "2E7D32", 2),
        ("AVERAGE PROGRESS", '=IF(SUM(BT_Appl)=0,0,SUMPRODUCT(BT_Appl,BT_Prog)/SUM(BT_Appl))', PCT, '="all requirements"', NAVY, 2),
        ("IN PROGRESS", '=COUNTIF(BT_Health,"On Track")+COUNTIF(BT_Health,"At Risk")', "0", '=COUNTIF(BT_Health,"At Risk")&" at risk"', NAVY, 2),
        ("NEEDS VERIFICATION", '=COUNTIF(BT_Health,"Needs Verification")', "0", '="complete, not yet verified"', "C9A227", 2),
        ("NOT STARTED", '=COUNTIF(BT_Health,"Not Started")', "0", '="no progress yet"', GRAY, 2),
        ("OVERDUE", "=SUM(BT_Overdue)", "0", '="past due date"', BURG, 2),
        ("NO OWNER", '=COUNTIFS(BT_Appl,1,BT_Done,0,BT_Owner,"")', "0", '="unassigned requirements"', BURG, 1),
    ]
    c = 2
    for lab, fm, fmt, sub, acc, w in tiles:
        tile(ws, 7, c, lab, fm, fmt, sub, w, acc)
        c += w
    ws.row_dimensions[8].height = 30
    ws.row_dimensions[9].height = 22
    # pillar table
    r0 = 11
    section(ws, r0, 2, "Progress by pillar", 13)
    hdr_row(ws, r0 + 1, 2, ["Pillar", "Pillar Name", "Requirements", "Done", "In Progress", "Needs Verif.",
                            "Not Started", "Overdue", "% Done", "Avg Progress", "Pillar Status", "Latest Due", "No Owner"])
    ws.row_dimensions[r0 + 1].height = 30
    LS = qs(sh[S_L])
    for i in range(9):
        r = r0 + 2 + i
        P = f"$B{r}"
        calc_cell(ws, r, 2, f'=IF({LS}!$E${10 + i}="","",{LS}!$E${10 + i}&"")', None, True)
        calc_cell(ws, r, 3, f'=IF({LS}!$F${10 + i}="","",{LS}!$F${10 + i}&"")', None, align="left")
        calc_cell(ws, r, 4, f'=IF({P}="","",COUNTIFS(BT_Pillar,{P},BT_Appl,1))', "0")
        calc_cell(ws, r, 5, f'=IF({P}="","",COUNTIFS(BT_Pillar,{P},BT_Done,1))', "0", True)
        calc_cell(ws, r, 6, f'=IF({P}="","",COUNTIFS(BT_Pillar,{P},BT_Health,"On Track")+COUNTIFS(BT_Pillar,{P},BT_Health,"At Risk"))', "0")
        calc_cell(ws, r, 7, f'=IF({P}="","",COUNTIFS(BT_Pillar,{P},BT_Health,"Needs Verification"))', "0")
        calc_cell(ws, r, 8, f'=IF({P}="","",COUNTIFS(BT_Pillar,{P},BT_Health,"Not Started"))', "0")
        calc_cell(ws, r, 9, f'=IF({P}="","",COUNTIFS(BT_Pillar,{P},BT_Overdue,1))', "0")
        calc_cell(ws, r, 10, f'=IF(N(D{r})=0,"",E{r}/D{r})', PCT)
        calc_cell(ws, r, 11, f'=IF(N(D{r})=0,"",SUMPRODUCT(--(BT_Pillar={P}),BT_Appl,BT_Prog)/D{r})', PCT)
        calc_cell(ws, r, 12, f'=IF({P}="","",IF(N(D{r})=0,"N/A",IF(E{r}=D{r},"Complete",IF(I{r}>0,"Overdue items",IF(H{r}=D{r},"Not Started","In Progress")))))')
        calc_cell(ws, r, 13, f'=IF(N(D{r})=0,"",IFERROR(1/(1/SUMPRODUCT(MAX((BT_Pillar={P})*BT_Appl*(1-BT_Done)*BT_Due))),""))', DATE)
        calc_cell(ws, r, 14, f'=IF({P}="","",COUNTIFS(BT_Pillar,{P},BT_Appl,1,BT_Done,0,BT_Owner,""))', "0")
    rt = r0 + 11
    calc_cell(ws, rt, 2, "TOTAL", None, True)
    calc_cell(ws, rt, 3, "All pillars", None, True, align="left")
    for col in range(4, 10):
        calc_cell(ws, rt, col, f"=SUM({L(col)}{r0 + 2}:{L(col)}{rt - 1})", "0", True)
    calc_cell(ws, rt, 10, f'=IF(N(D{rt})=0,"",E{rt}/D{rt})', PCT, True)
    calc_cell(ws, rt, 11, f'=IF(N(D{rt})=0,"",SUMPRODUCT(BT_Appl,BT_Prog)/D{rt})', PCT, True)
    calc_cell(ws, rt, 12, f'=IF(N(D{rt})=0,"N/A",IF(E{rt}=D{rt},"Complete",IF(I{rt}>0,"Overdue items",IF(H{rt}=D{rt},"Not Started","In Progress"))))', None, True)
    calc_cell(ws, rt, 13, '=IFERROR(1/(1/SUMPRODUCT(MAX(BT_Appl*(1-BT_Done)*BT_Due))),"")', DATE, True)
    calc_cell(ws, rt, 14, f"=SUM(N{r0 + 2}:N{rt - 1})", "0", True)
    for col in range(2, 15):
        ws.cell(rt, col).fill = fill(LGRAY)
    bar_cf(ws, f"J{r0 + 2}:J{rt}", "2E7D32")
    bar_cf(ws, f"K{r0 + 2}:K{rt}")
    text_cf(ws, f"L{r0 + 2}:L{rt}", {"Complete": "Green", "In Progress": "Yellow", "Overdue items": "Red",
                                     "Not Started": "Gray"})
    ws.conditional_formatting.add(f"I{r0 + 2}:I{rt}", FormulaRule(formula=[f"I{r0 + 2}>0"], font=Font(color="C00000", bold=True)))
    ws.conditional_formatting.add(f"N{r0 + 2}:N{rt}", FormulaRule(formula=[f"N{r0 + 2}>0"], font=Font(color="C00000", bold=True)))
    # three layers
    rl = rt + 2
    section(ws, rl, 2, "The 3 layers that create reliability", 13)
    hdr_row(ws, rl + 1, 2, ["Layer", "", "Pillars", "", "Requirements", "Done", "Overdue", "% Done", "Avg Progress", "What it does", "", "", ""])
    for a, b in [(2, 3), (4, 5), (11, 14)]:
        ws.merge_cells(start_row=rl + 1, start_column=a, end_row=rl + 1, end_column=b)
    for k, (lay, pil, txt) in enumerate([("Foundation", "LEAD · EHS · FSQ", "Creates stable & safe conditions"),
                                         ("Equipment Focus", "AM · PM · VO", "Improves equipment health"),
                                         ("Capability & Governance", "E&T · IM · SC", "Builds capability & drives execution")]):
        r = rl + 2 + k
        for a, b in [(2, 3), (4, 5), (11, 14)]:
            ws.merge_cells(start_row=r, start_column=a, end_row=r, end_column=b)
        calc_cell(ws, r, 2, lay, None, True, align="left")
        calc_cell(ws, r, 4, pil, None)
        calc_cell(ws, r, 6, f'=COUNTIFS(BT_Layer,$B{r},BT_Appl,1)', "0")
        calc_cell(ws, r, 7, f'=COUNTIFS(BT_Layer,$B{r},BT_Done,1)', "0", True)
        calc_cell(ws, r, 8, f'=COUNTIFS(BT_Layer,$B{r},BT_Overdue,1)', "0")
        calc_cell(ws, r, 9, f'=IF(N(F{r})=0,"",G{r}/F{r})', PCT)
        calc_cell(ws, r, 10, f'=IF(N(F{r})=0,"",SUMPRODUCT(--(BT_Layer=$B{r}),BT_Appl,BT_Prog)/F{r})', PCT)
        calc_cell(ws, r, 11, txt, None, align="left")
    bar_cf(ws, f"I{rl + 2}:I{rl + 4}", "2E7D32")
    bar_cf(ws, f"J{rl + 2}:J{rl + 4}")
    ws.conditional_formatting.add(f"H{rl + 2}:H{rl + 4}", FormulaRule(formula=[f"H{rl + 2}>0"], font=Font(color="C00000", bold=True)))
    # needs attention list
    r1 = rl + 6
    section(ws, r1, 2, "Needs attention (overdue first, then at risk / awaiting verification)", 13)
    hdr_row(ws, r1 + 1, 2, ["Ref", "Requirement", "", "", "Owner", "", "Due Date", "Health", "What's missing", "", "", "Next step", ""])
    for a, b in [(3, 5), (6, 7), (10, 12), (13, 14)]:
        ws.merge_cells(start_row=r1 + 1, start_column=a, end_row=r1 + 1, end_column=b)
    for k in range(8):
        r = r1 + 2 + k
        m = f"MATCH(LARGE(BT_Key,{k + 1}),BT_Key,0)"
        g = f"LARGE(BT_Key,{k + 1})<=0"
        for a, b in [(3, 5), (6, 7), (10, 12), (13, 14)]:
            ws.merge_cells(start_row=r, start_column=a, end_row=r, end_column=b)
        calc_cell(ws, r, 2, f'=IF({g},"",INDEX(BT_Ref,{m})&"")', None, True, size=8)
        calc_cell(ws, r, 3, f'=IF({g},"",INDEX(BT_Req,{m})&"")', None, align="left", size=8)
        calc_cell(ws, r, 6, f'=IF({g},"",IF(INDEX(BT_Owner,{m})="","(no owner)",INDEX(BT_Owner,{m})&""))', None, align="left", size=8)
        calc_cell(ws, r, 8, f'=IF({g},"",IF(INDEX(BT_Due,{m})="","",INDEX(BT_Due,{m})))', DATE, size=8)
        calc_cell(ws, r, 9, f'=IF({g},"",INDEX(BT_Health,{m})&"")', None, size=8)
        calc_cell(ws, r, 10, f'=IF({g},"",INDEX(BT_Missing,{m})&IF(INDEX(BT_Gap,{m})="",""," "&INDEX(BT_Gap,{m})))', None, align="left", size=8)
        calc_cell(ws, r, 13, f'=IF({g},"",INDEX(BT_Next,{m})&"")', None, align="left", size=8)
        ws.row_dimensions[r].height = 24
    text_cf(ws, f"I{r1 + 2}:I{r1 + 9}", HEALTH)
    ws.conditional_formatting.add(f"F{r1 + 2}:F{r1 + 9}", FormulaRule(formula=[f'F{r1 + 2}="(no owner)"'], font=Font(color="C00000", bold=True)))
    # owner summary
    r2 = r1 + 11
    section(ws, r2, 2, "By owner", 6)
    hdr_row(ws, r2 + 1, 2, ["Owner", "", "Assigned", "Done", "Overdue", "% Done"])
    ws.merge_cells(start_row=r2 + 1, start_column=2, end_row=r2 + 1, end_column=3)
    for k in range(13):
        r = r2 + 2 + k
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        if k < 12:
            nm = f"INDEX(bl_OwnerAll,{k + 1})"
            calc_cell(ws, r, 2, f'={nm}&""', None, align="left", size=8)
            O = f"$B{r}"
            guard = f'{O}=""'
        else:
            calc_cell(ws, r, 2, "(no owner)", None, align="left", size=8)
            O = '""'
            guard = "FALSE"
        calc_cell(ws, r, 4, f'=IF({guard},"",COUNTIFS(BT_Owner,{O},BT_Appl,1))', "0", size=8)
        calc_cell(ws, r, 5, f'=IF({guard},"",COUNTIFS(BT_Owner,{O},BT_Done,1))', "0", size=8)
        calc_cell(ws, r, 6, f'=IF({guard},"",COUNTIFS(BT_Owner,{O},BT_Overdue,1))', "0", size=8)
        calc_cell(ws, r, 7, f'=IF(N(D{r})=0,"",E{r}/D{r})', PCT, size=8)
    define(wb, "bl_OwnerAll", f"{qs(sh[S_L])}!$K$10:$K${9 + NOWN}")
    bar_cf(ws, f"G{r2 + 2}:G{r2 + 14}", "2E7D32")
    ws.conditional_formatting.add(f"F{r2 + 2}:F{r2 + 14}", FormulaRule(formula=[f"N(F{r2 + 2})>0"], font=Font(color="C00000", bold=True)))
    # charts
    ch = BarChart()
    ch.type = "bar"
    ch.add_data(Reference(ws, min_col=10, min_row=r0 + 1, max_row=r0 + 10), titles_from_data=True)
    ch.add_data(Reference(ws, min_col=11, min_row=r0 + 1, max_row=r0 + 10), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=r0 + 2, max_row=r0 + 10))
    color_series(ch.series[0], "2E7D32")
    color_series(ch.series[1], NAVY)
    ch.x_axis.scaling.orientation = "maxMin"
    ch.y_axis.scaling.min = 0
    ch.y_axis.scaling.max = 1
    ch.y_axis.number_format = "0%"
    style_chart(ch, "% done vs average progress by pillar", w=12.5, h=7.5)
    ws.add_chart(ch, f"I{r2}")
    # doughnut data (small, gray) next to owner table
    labels = ["Done", "In progress", "Needs verification", "Not started", "Overdue"]
    forms = ['=COUNTIF(BT_Health,"Done")', '=COUNTIF(BT_Health,"On Track")+COUNTIF(BT_Health,"At Risk")',
             '=COUNTIF(BT_Health,"Needs Verification")', '=COUNTIF(BT_Health,"Not Started")', '=COUNTIF(BT_Health,"Overdue")']
    rd = r2 + 17
    ws.cell(rd, 2, "Status mix (chart data)").font = f(7, False, GRAY)
    for i, (lab, fm) in enumerate(zip(labels, forms)):
        ws.cell(rd + 1 + i, 2, lab).font = f(8)
        ws.cell(rd + 1 + i, 3, fm).font = f(8)
    dc = DoughnutChart()
    dc.add_data(Reference(ws, min_col=3, min_row=rd + 1, max_row=rd + 5))
    dc.set_categories(Reference(ws, min_col=2, min_row=rd + 1, max_row=rd + 5))
    for i, cl in enumerate(["2E7D32", "F2C94C", "C9A227", "BFBFBF", "C62828"]):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = cl
        dc.series[0].dPt.append(pt)
    dc.holeSize = 55
    style_chart(dc, "Requirement status", w=8, h=6.5, legend_pos="r")
    ws.add_chart(dc, f"D{rd}")
    landscape(ws, f"A1:N{rd + 14}", one_page=True)
    protect(ws)

    # ---------------- How To Use ----------------
    ws = sh[S_H]
    title_band(ws, "How To Use", "Five steps. About 10 minutes a week per owner.", 6, legend=False)
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 100
    steps = [
        ("1. Add owners", "On 'Owners', type each person's name (Active? = Yes). They appear in the Owner and Verified By dropdowns."),
        ("2. Assign", "On 'Bronze Tracker', pick an Owner, Start Date and Due Date for each of the 32 requirements."),
        ("3. Update weekly", "Owners update Status, % Complete, Gap / What's Missing, Next Step and Last Updated."),
        ("4. Prove it", "When finished: Status = Complete, paste the Teams link to the evidence (Teams → DunGPS → General → Shared → pillar folder → Copy link)."),
        ("5. Verify", "The verifier opens the evidence and sets Verified? = Yes and Verified By. Only then is the requirement 'Done'."),
        ("Review", "Open 'Dashboard' in the weekly meeting. It prints on one landscape page."),
        ("", ""),
        ("Health – Done", "Complete + evidence link + Verified? = Yes."),
        ("Health – Needs Verification", "Marked Complete but not yet verified (or no evidence link)."),
        ("Health – Overdue", "Past the due date and not Done."),
        ("Health – At Risk", "Due within the warning window (default 14 days) and less than 75% complete."),
        ("Health – On Track", "Started and not at risk."),
        ("Health – Not Started", "Status Not Started and 0%."),
        ("Health – N/A", "Only when CHG has approved that the requirement does not apply. N/A rows are not counted."),
        ("% Done", "Requirements Done ÷ requirements that count (excluding N/A)."),
        ("Average Progress", "Average of % Complete (Done = 100%; Complete but unverified counts at least 75%)."),
        ("Bronze Ready", "The banner turns green only when every requirement that counts is Done."),
        ("", ""),
        ("Source", "The 32 requirements are copied word for word from CHG's 'GPS Pillar Bronze Level Requirements'. "
                   "The 'Evidence Needed' column is a suggestion – replace with CHG's official evidence when available."),
        ("Settings", "'Lists' sheet: status date (blank = today), the At-Risk window, and dropdown values."),
        ("Adding rows", "13 spare rows are pre-formatted at the bottom of the tracker if CHG adds requirements."),
    ]
    r = 5
    for a, b in steps:
        ws.cell(r, 2, a).font = f(10, True, NAVY)
        c = ws.cell(r, 3, b)
        c.font = f(10)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 18 if len(b) < 110 else 32
        r += 1
    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "."
    build(False, f"{out}/CHG_Bronze_Level_Tracker.xlsx")
    build(True, f"{out}/CHG_Bronze_Level_Tracker_DEMO.xlsx")
    print("built")
