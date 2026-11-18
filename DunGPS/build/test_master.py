"""Independent verification of the DunGPS master workbook calculations (run after LibreOffice recalc).

Re-computes key metrics in Python straight from the input columns and compares them with the
values Excel/LibreOffice calculated. Usage: python test_master.py <recalculated.xlsx> [expect_filter_pillar]
"""
import sys
import datetime as dt
import openpyxl

path = sys.argv[1]
flt = sys.argv[2] if len(sys.argv) > 2 else None
wb = openpyxl.load_workbook(path, data_only=True)
fails = []


def check(name, got, exp, tol=1e-6):
    ok = (abs((got or 0) - (exp or 0)) <= tol) if isinstance(exp, (int, float)) else (got == exp)
    print(("PASS " if ok else "FAIL ") + f"{name}: got={got!r} expected={exp!r}")
    if not ok:
        fails.append(name)


def table(ws, hdr_row):
    hdr = [c.value for c in ws[hdr_row]]
    rows = []
    for r in ws.iter_rows(min_row=hdr_row + 1, values_only=True):
        rows.append({(h or "").replace("ƒ ", ""): v for h, v in zip(hdr, r)})
    return rows


def d(v):
    if isinstance(v, dt.datetime):
        return v.date()
    return v


cfg = wb["Configuration & Dropdown Lists"]
asof = d(cfg["C8"].value)
cap = cfg["C9"].value
print("As of", asof, "cap", cap)
T = [r for r in table(wb["Pillar Deployment Tracker"], 5) if r["Pillar"]]
G = [r for r in table(wb["Gap Register"], 5) if r["Pillar"]]
B = [r for r in table(wb["Bronze Star Readiness"], 17) if r["Req ID"]]
O = [r for r in table(wb["Owner Directory"], 5) if r["Employee Name"]]


def inscope(t):
    return t["Required?"] != "No" and t["Applicability"] != "N/A - Approved"


def verified(t):
    return (t["Applicability"] != "N/A - Approved" and t["Status"] == "Complete"
            and t["Verification Status"] == "Verified" and bool(t["Evidence Link (Teams / SharePoint URL)"]))


def credit(t):
    if verified(t):
        return 1.0
    if t["Status"] == "Complete":
        return cap
    return min(cap, t["Work Progress %"] or 0)


def planned(t):
    s, e = d(t["Start Date"]), d(t["Target Date"])
    if not e:
        return 0
    if not s:
        return 1 if asof >= e else 0
    return max(0, min(1, (asof - s).days / max(1, (e - s).days)))


def overdue(t):
    e = d(t["Target Date"])
    return t["Applicability"] != "N/A - Approved" and not verified(t) and e is not None and e < asof


# ---- row-level checks
for t in T:
    check(f"row {t['Task ID']} verified flag", t["Verified (1/0)"], 1 if verified(t) else 0)
    check(f"row {t['Task ID']} overdue flag", t["Overdue (1/0)"], 1 if overdue(t) else 0)
    if t["Applicability"] != "N/A - Approved":
        check(f"row {t['Task ID']} credit", t["Completion %"], credit(t))
        check(f"row {t['Task ID']} planned", t["Planned % to Date"], planned(t), 1e-4)

# ---- pillar roll-ups from Calc Engine
calc = wb["Calc Engine"]
hdr = {calc.cell(6, c).value: c for c in range(10, 40)}
pillar_rows = {calc.cell(r, 10).value: r for r in range(7, 20) if calc.cell(r, 10).value}
sel = (lambda t: True) if not flt else (lambda t: t["Pillar"] == flt)
for p, r in pillar_rows.items():
    rows = [t for t in T if (t["Pillar"] == p if p != "TOTAL" else True) and inscope(t)]
    n = len(rows)
    v = sum(verified(t) for t in rows)
    check(f"{p} in-scope", calc.cell(r, hdr["In Scope"]).value, n)
    check(f"{p} verified count", calc.cell(r, hdr["Verified Complete"]).value, v)
    if n:
        check(f"{p} verified %", calc.cell(r, hdr["Verified %"]).value, v / n)
        check(f"{p} credit %", calc.cell(r, hdr["Completion Credit %"]).value, sum(credit(t) for t in rows) / n)
        check(f"{p} planned %", calc.cell(r, hdr["Planned %"]).value, sum(planned(t) for t in rows) / n, 1e-4)
        check(f"{p} evidence submitted %", calc.cell(r, hdr["Evidence Submitted %"]).value,
              sum(bool(t["Evidence Link (Teams / SharePoint URL)"]) for t in rows) / n)
    else:
        check(f"{p} verified % (zero denominator)", calc.cell(r, hdr["Verified %"]).value, None)
    check(f"{p} overdue", calc.cell(r, hdr["Overdue"]).value, sum(overdue(t) for t in rows))
    na = [t for t in T if (t["Pillar"] == p or p == "TOTAL") and t["Applicability"] == "N/A - Approved"]
    check(f"{p} N/A approved (excluded)", calc.cell(r, hdr["N/A Approved"]).value, len(na))


# ---- gaps
def g_closed(g):
    return (g["Status"] == "Closed - Verified" and g["Verification"] == "Verified" and g["Closure Date"] is not None)


counted = [g for g in G if g["Status"] != "Cancelled"]
open_ = [g for g in counted if not g_closed(g)]
closed = [g for g in counted if g_closed(g)]
od = [g for g in open_ if g["Due Date"] and d(g["Due Date"]) < asof]
names = {}
for r in range(6, 60):
    if calc.cell(r, 45).value:
        names[calc.cell(r, 45).value] = calc.cell(r, 46).value
check("gaps total identified", names["k_GapTotal"], len(counted))
check("gaps open backlog", names["k_GapOpen"], len(open_))
check("gaps closed verified", names["k_GapClosedV"], len(closed))
check("gaps overdue", names["k_GapOverdue"], len(od))
check("gap closure rate", names["k_GapClosure"] if counted else None, (len(closed) / len(counted)) if counted else None)

# ---- Bronze
cnt = [b for b in B if b["Mandatory?"] != "No" and b["Applicable?"] != "N/A - Approved by CHG"]


def b_ready(b):
    linked = [t for t in T if t["Bronze Req ID"] == b["Req ID"] and inscope(t)]
    gaps = [g for g in open_ if g["Bronze Req ID"] == b["Req ID"]]
    return (b["Verification Status"] == "Verified" and bool(b["Evidence Link"]) and
            (not linked or all(verified(t) for t in linked)) and not gaps)


ready = [b for b in cnt if b_ready(b)]
check("bronze mandatory", names["k_BronzeMand"], len(cnt))
check("bronze ready", names["k_BronzeReady"], len(ready))
banner = names["k_BronzeBanner"]
exp_ready = len(cnt) > 0 and len(ready) == len(cnt)
check("bronze banner never says READY unless all verified", banner.startswith("BRONZE STAR READY"), exp_ready)

# ---- owner dropdown helper (active only)
active = [o["Employee Name"] for o in O if o["Active / Inactive"] == "Active"]
lst = [calc.cell(r, 7).value for r in range(6, 6 + len(active) + 2)]
check("active-owner dropdown list", [x for x in lst if x], active)
check("unassigned tasks", names["k_Unassigned"], sum(1 for t in T if inscope(t) and not t["Primary Owner"]))

# ---- data quality flags present for edge cases
flags = {t["Task ID"]: (t["Data Quality Flags"] or "") for t in T}
print("Sample flags:", [f for f in flags.values() if f][:8])

# ---- filtered pillar block (Management Review)
if flt:
    fr = {calc.cell(r, 10).value: r for r in range(24, 38) if calc.cell(r, 10).value}
    rows = [t for t in T if inscope(t) and t["Pillar"] == flt]
    check(f"MR filtered TOTAL in-scope = {flt}", calc.cell(fr["TOTAL"], hdr["In Scope"]).value, len(rows))
    other = [p for p in fr if p not in ("TOTAL", flt)][0]
    check(f"MR filtered other pillar {other} = 0", calc.cell(fr[other], hdr["In Scope"]).value, 0)

print("\nRESULT:", "ALL PASS" if not fails else f"{len(fails)} FAILURES: {fails[:10]}")
sys.exit(1 if fails else 0)
