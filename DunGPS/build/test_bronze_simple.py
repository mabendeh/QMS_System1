"""Independent check of CHG_Bronze_Level_Tracker (run after LibreOffice recalc).

Re-computes requirement status, progress (from working logs), per-pillar completion rates (gauges),
overdue and evidence rules straight from the input cells of each pillar sheet.
"""
import sys
import datetime as dt
import openpyxl

ORDER = ["LEAD", "EHS", "FSQ", "AM", "PM", "VO", "E&T", "IM", "SC"]
OFFICIAL = {"LEAD": 1, "EHS": 2, "FSQ": 9, "AM": 3, "PM": 3, "VO": 7, "E&T": 3, "IM": 1, "SC": 3}
path = sys.argv[1]
wb = openpyxl.load_workbook(path, data_only=True)
fails = []


def chk(n, g, e, tol=1e-9):
    ok = (abs((g or 0) - (e or 0)) <= tol) if isinstance(e, (int, float)) and not isinstance(e, bool) else g == e
    print(("PASS " if ok else "FAIL ") + f"{n}: {g!r} vs {e!r}")
    if not ok:
        fails.append(n)


def d(v):
    return v.date() if isinstance(v, dt.datetime) else v


asof = d(wb["Lists"]["C6"].value)


def read_table(ws, header_text):
    """Find a header row whose first cell (col B) equals header_text, return list of dict rows until blank key."""
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 2).value == header_text:
            hdr = [str(c.value).replace("ƒ ", "") if c.value is not None else "" for c in ws[r]]
            out = []
            for row in ws.iter_rows(min_row=r + 1, values_only=True):
                if row[1] is None and (len(out) > 0 and all(v in (None, "") for v in row[1:4])):
                    break
                out.append(dict(zip(hdr, row)))
            return out
    raise KeyError(header_text)


tot_done = tot_appl = 0
for p in ORDER:
    ws = wb[p]
    reqs = [r for r in read_table(ws, "Ref") if r.get("Bronze Requirement (official)")]
    log = [r for r in read_table(ws, "Log ID") if r.get("Bronze Ref (select)")]
    chk(f"{p} official requirement count", len(reqs), OFFICIAL[p])

    def ldone(x):
        return x["Status"] == "Complete" and x["Verified?"] == "Yes" and bool(x["Evidence Link (paste Teams link)"])

    appl = [r for r in reqs if r["Status"] != "N/A"]
    done = [r for r in appl if r["Status"] == "Complete" and r["Verified?"] == "Yes" and r["Evidence Link (paste Teams link)"]]
    tot_done += len(done)
    tot_appl += len(appl)
    for r in appl:
        lr = [x for x in log if x["Bronze Ref (select)"] == r["Ref"] and x["Status"] != "N/A"]
        if r in done:
            exp = 1
        elif lr:
            exp = sum(ldone(x) for x in lr) / len(lr)
        elif r["Status"] == "Complete":
            exp = max(0.75, r["% Complete (if no log rows)"] or 0)
        else:
            exp = r["% Complete (if no log rows)"] or 0
        chk(f"{p} {r['Ref']} progress", r["Progress"], exp)
        hrs = sum((x["Scheduled Time (hrs)"] or 0) for x in log if x["Bronze Ref (select)"] == r["Ref"])
        chk(f"{p} {r['Ref']} scheduled hours", r["Scheduled Hours (from log)"] or 0, hrs)
        if r["Status"] == "Complete" and not r["Evidence Link (paste Teams link)"]:
            chk(f"{p} {r['Ref']} evidence missing flagged", r["Health"], "Evidence Missing")
        if r not in done and d(r["Due Date"]) and d(r["Due Date"]) < asof and r["Health"] != "Evidence Missing":
            chk(f"{p} {r['Ref']} overdue flagged", r["Health"], "Overdue")
    # gauge data row on Lists
    L_ = wb["Lists"]
    for row in range(14, 25):
        if L_.cell(row, 23).value == p:
            chk(f"{p} gauge rate", L_.cell(row, 26).value, (len(done) / len(appl)) if appl else 0)
            seg = [L_.cell(row, c).value for c in range(27, 32)]
            chk(f"{p} gauge segments sum to 2 (half circle hidden)", round(sum(seg), 9), 2)
    t = ws.cell(6, 2).value  # completion tile
    chk(f"{p} sheet completion tile", t, (len(done) / len(appl)) if appl else 0)

D = wb["Dashboard"]
chk("overall gauge / % done tile", D.cell(8, 2).value, (tot_done / tot_appl) if tot_appl else 0)
chk("banner honest", str(D["B5"].value).startswith("✔"), tot_appl > 0 and tot_done == tot_appl)
w = openpyxl.load_workbook(path)
chk("9 gauges + overall on dashboard", len(w["Dashboard"]._charts), 10)
dv = sum(1 for p in ORDER for x in w[p].data_validations.dataValidation if "LEN(" in (x.formula1 or ""))
chk("evidence-gated dropdowns (4 per pillar sheet)", dv, 36)
print("RESULT:", "ALL PASS" if not fails else f"{len(fails)} FAIL {fails[:8]}")
sys.exit(1 if fails else 0)
