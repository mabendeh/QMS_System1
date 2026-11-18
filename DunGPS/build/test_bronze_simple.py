"""Independent check of CHG_Bronze_Level_Tracker (after recalc)."""
import sys, datetime as dt, openpyxl
wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
asof = wb["Lists"]["C6"].value
asof = asof.date() if isinstance(asof, dt.datetime) else asof
ws = wb["Bronze Tracker"]
hdr = [(c.value or "").replace("ƒ ", "") for c in ws[5]]
R = [dict(zip(hdr, r)) for r in ws.iter_rows(min_row=6, values_only=True)]
R = [r for r in R if r["Bronze Requirement"]]
fails = []
def chk(n, g, e):
    ok = g == e or (isinstance(e, float) and abs((g or 0) - e) < 1e-9)
    print(("PASS " if ok else "FAIL ") + f"{n}: {g!r} vs {e!r}"); ok or fails.append(n)
def d(v): return v.date() if isinstance(v, dt.datetime) else v
appl = [r for r in R if r["Status"] != "N/A"]
done = [r for r in appl if r["Status"] == "Complete" and r["Verified?"] == "Yes" and r["Evidence Link (Teams)"]]
over = [r for r in appl if r not in done and d(r["Due Date"]) and d(r["Due Date"]) < asof]
chk("32 official requirements loaded", len(R), 32)
chk("pillar order", list(dict.fromkeys(r["Pillar"] for r in R)), ["LEAD","EHS","FSQ","AM","PM","VO","E&T","IM","SC"])
D = wb["Dashboard"]
tiles = {D.cell(7, c).value: D.cell(8, c).value for c in range(2, 15) if D.cell(7, c).value}
chk("% done", tiles["% DONE (VERIFIED)"], (len(done) / len(appl)) if appl else 0)
chk("overdue", tiles["OVERDUE"], len(over))
chk("no owner", tiles["NO OWNER"], sum(1 for r in appl if r not in done and not r["Owner"]))
for row in range(13, 22):
    p = D.cell(row, 2).value
    chk(f"{p} done", D.cell(row, 5).value, sum(1 for r in done if r["Pillar"] == p))
    chk(f"{p} requirements", D.cell(row, 4).value, sum(1 for r in appl if r["Pillar"] == p))
ready = D["B5"].value
chk("banner honest", ready.startswith("✔"), len(done) == len(appl) and len(appl) > 0)
print("RESULT:", "ALL PASS" if not fails else f"FAIL {fails}")
