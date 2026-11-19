"""Independent check of CHG_Supervisor_90-Day_Projects (run after LibreOffice recalc)."""
import sys
import datetime as dt
import openpyxl

path = sys.argv[1]
wb = openpyxl.load_workbook(path, data_only=True)
fails = []


def chk(n, g, e, tol=1e-6):
    ok = (abs((g or 0) - (e or 0)) <= tol) if isinstance(e, (int, float)) and not isinstance(e, bool) else g == e
    print(("PASS " if ok else "FAIL ") + f"{n}: {g!r} vs {e!r}")
    if not ok:
        fails.append(n)


def d(v):
    return v.date() if isinstance(v, dt.datetime) else v


asof = d(wb["Lists"]["C6"].value)
menu = {r[1]: r for r in wb["Project Menu"].iter_rows(min_row=6, values_only=True) if r[1]}
pct_col = [c.value for c in wb["Project Menu"][5]].index("Default target reduction")
tot_sav = 0
for g in ["NL", "MR", "PK"]:
    for i in range(1, 5):
        ws = wb[f"{g}-{i}"]
        code = ws["C6"].value
        chk(f"{g}-{i} project name from menu", ws["C7"].value, menu[code][3])
        base, own, vol, price = ws["C15"].value, ws["C16"].value, ws["C18"].value, ws["C19"].value
        target = own if isinstance(own, (int, float)) else (base * (1 - menu[code][pct_col]) if isinstance(base, (int, float)) else None)
        if target is not None:
            chk(f"{g}-{i} target", ws["C17"].value, target)
        acts = [(r - 24, ws.cell(r, 6).value) for r in range(25, 38)]
        have = [(w, a) for w, a in acts if isinstance(a, (int, float))]
        sav = 0.0
        for w, a in acts:
            r = 24 + w
            if isinstance(base, (int, float)) and isinstance(target, (int, float)):
                chk(f"{g}-{i} wk{w} target path", ws.cell(r, 5).value, base - (base - target) * min(1, max(0, (w - 1) / 7)))
            if isinstance(a, (int, float)) and all(isinstance(x, (int, float)) for x in (base, vol, price)):
                s = (base - a) * vol * price
                sav += s
                chk(f"{g}-{i} wk{w} savings", ws.cell(r, 9).value, s, 1e-4)
        chk(f"{g}-{i} savings to date", ws["H11"].value or 0, sav, 1e-3)
        tot_sav += sav
        if have:
            last4 = [a for _, a in have[-4:]]
            l4 = sum(last4) / len(last4)
            chk(f"{g}-{i} latest 4-wk avg", ws["J7"].value, l4, 1e-6)
            ach = (base - l4) / (base - target)
            chk(f"{g}-{i} improvement achieved", ws["L7"].value, ach, 1e-6)
            lastw = have[-1][0]
            start = d(ws["C14"].value)
            cur = min(13, max(0, (asof - start).days // 7 + 1))
            exp = min(1, max(0, (lastw - 1) / 7))
            st = ("Data Missing" if cur - lastw >= 2 else "On Track" if ach >= exp - 0.15 else
                  "Behind" if ach >= exp - 0.35 else "Off Track")
            chk(f"{g}-{i} status", ws["L11"].value, st)
        else:
            chk(f"{g}-{i} status (no data)", ws["L11"].value, "Not Started")
D = wb["Dashboard"]
vals = {D.cell(5, c).value: D.cell(6, c).value for c in range(2, 16) if D.cell(5, c).value}
chk("dashboard savings to date", vals["SAVINGS TO DATE"], tot_sav, 1e-3)
w = openpyxl.load_workbook(path)
gated = sum(1 for n in w.sheetnames if n[:2] in ("NL", "MR", "PK") for x in w[n].data_validations.dataValidation
            if "LEN(" in (x.formula1 or ""))
chk("evidence-gated Done dropdown on every project sheet", gated, 12)
print("RESULT:", "ALL PASS" if not fails else f"{len(fails)} FAIL {fails[:8]}")
sys.exit(1 if fails else 0)
