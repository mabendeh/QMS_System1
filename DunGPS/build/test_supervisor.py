"""Independent check of the Supervisor Productivity Program workbook (after recalc)."""
import sys, openpyxl
wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
fails = []
def check(n, g, e, tol=1e-6):
    ok = abs((g or 0) - (e or 0)) <= tol if isinstance(e, (int, float)) else g == e
    print(("PASS " if ok else "FAIL ") + f"{n}: got={g!r} exp={e!r}")
    if not ok: fails.append(n)
ws = wb["Project Assignments"]
hdr = [(c.value or "").replace("ƒ ", "") for c in ws[5]]
rows = [dict(zip(hdr, r)) for r in ws.iter_rows(min_row=6, values_only=True)]
rows = [r for r in rows if r["Supervisor"] and r["Project Code (select)"] and r["Status"] != "Cancelled"]
W = [0.15, 0.2, 0.3, 0.2, 0.15]
def ms(r): return sum(w * (1 if r[f"M{j+1}"] == "Done" else 0.5 if r[f"M{j+1}"] == "In Progress" else 0) for j, w in enumerate(W))
def ver(r): return (r["Status"] == "Complete" and all(r[f"M{j}"] == "Done" for j in range(1, 6))
                    and r["Coach Verification"] == "Verified" and bool(r["Evidence Link"]))
for r in rows:
    check(f"{r['Assignment ID']} milestone%", r["Milestone %"], ms(r))
    check(f"{r['Assignment ID']} verified", r["Verified (1/0)"], 1 if ver(r) else 0)
    check(f"{r['Assignment ID']} project name from library", bool(r["Project"]), True)
d = wb["Program Dashboard"]
vals = {d.cell(8, c).value: d.cell(9, c).value for c in range(2, 19) if d.cell(8, c).value}
check("projects assigned", vals["PROJECTS ASSIGNED"], len(rows))
check("verified complete", vals["VERIFIED COMPLETE"], sum(ver(r) for r in rows))
check("completion rate", vals["COMPLETION RATE"], sum(ver(r) for r in rows) / len(rows) if rows else 0)
sc = {d.cell(r, 2).value: [d.cell(r, c).value for c in range(2, 18)] for r in range(14, 38) if d.cell(r, 2).value}
for s, v in sc.items():
    mine = [r for r in rows if r["Supervisor"] == s]
    check(f"scorecard {s} assigned", v[3], len(mine))
    check(f"scorecard {s} verified", v[5], sum(ver(r) for r in mine))
    if not mine: check(f"scorecard {s} RAG no project", v[14], "No project")
print("RESULT:", "ALL PASS" if not fails else f"{len(fails)} FAIL {fails[:8]}")
