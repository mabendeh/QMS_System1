import { pageUser } from "@/lib/auth";
import { loadPlant } from "@/lib/data/plant";
import { coverageFor } from "@/lib/data/analytics";
import { Alert, Badge, Card, PageHeader, Stat } from "@/components/ui";
import { TargetsEditor } from "./TargetsEditor";

export const metadata = { title: "Shift Coverage" };

export default async function CoveragePage({ searchParams }: { searchParams: Promise<{ edit?: string }> }) {
  const u = await pageUser("coverage:view");
  const { edit } = await searchParams;
  const plant = await loadPlant();
  const { shifts, rows } = await coverageFor(plant);
  const canEdit = u.permissions.includes("coverage:manage");
  const gaps = rows.flatMap((r) => r.cells.filter((c) => c.status === "GAP"));
  return (
    <div className="grid gap-5">
      <PageHeader
        title="Manufacturing coverage"
        subtitle="Qualified = every competency in the department is currently Competent and not expired. Counted per shift – people on other shifts are not assumed available."
        actions={<>
          <a href="/api/export/coverage?format=xlsx" className="inline-flex h-10 items-center rounded-md border border-line bg-white px-4 text-sm font-semibold hover:bg-canvas">Excel</a>
          <a href="/api/export/coverage?format=pdf" className="inline-flex h-10 items-center rounded-md border border-line bg-white px-4 text-sm font-semibold hover:bg-canvas">PDF</a>
          {canEdit && !edit && <a href="/coverage?edit=1" className="inline-flex h-10 items-center rounded-md bg-navy-800 px-4 text-sm font-semibold text-white">Edit staffing targets</a>}
        </>}
      />
      <section className="grid grid-cols-2 gap-3 md:grid-cols-4">
        <Stat label="Departments covered" value={`${rows.filter((r) => r.status === "COVERED").length} / ${rows.filter((r) => r.status !== "NO_TARGET").length}`} tone="blue" hint="All shifts at or above target" />
        <Stat label="Shift gaps" value={gaps.length} tone={gaps.length ? "red" : "gray"} hint="Department-shift slots below target" />
        <Stat label="Scheduled qualified" value="Unknown" tone="gray" hint="No scheduling system integrated" />
        <Stat label="Available qualified" value="Unknown" tone="gray" hint="No attendance data integrated" />
      </section>
      <Alert tone="info">Scheduled and available headcount are shown as <b>Unknown</b> because no scheduling or attendance system is connected. The figures below are total qualified headcount by home shift.</Alert>
      {edit && canEdit ? (
        <TargetsEditor shifts={shifts.map((s) => ({ id: s.id, code: s.code }))} rows={rows.map((r) => ({ areaId: r.areaId, areaName: r.areaName, targets: r.cells.map((c) => c.target) }))} />
      ) : (
        <Card className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="bg-canvas text-left text-xs font-bold uppercase tracking-wide text-muted">
                <th className="px-4 py-3">Department</th>
                {shifts.map((s) => <th key={s.id} className="px-3 py-3 text-center">Shift {s.code}<div className="font-normal normal-case">qualified / target</div></th>)}
                <th className="px-3 py-3 text-center">Total qualified</th><th className="px-3 py-3 text-center">Target</th><th className="px-4 py-3">Status</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.areaId} className="border-t border-line">
                  <td className="px-4 py-3 font-semibold">{r.areaName}</td>
                  {r.cells.map((c) => (
                    <td key={c.shiftId} className="px-3 py-3 text-center">
                      <span className={`inline-flex min-w-[64px] justify-center rounded px-2 py-1 font-bold ${c.status === "GAP" ? "bg-red-50 text-st-red" : c.status === "COVERED" ? "bg-blue-50 text-st-blue" : "bg-gray-100 text-muted"}`}
                        aria-label={`${c.qualified} qualified, target ${c.target ?? "not set"}, ${c.status === "GAP" ? "gap" : c.status === "COVERED" ? "covered" : "no target"}`}>
                        {c.qualified} / {c.target ?? "–"}
                      </span>
                    </td>
                  ))}
                  <td className="px-3 py-3 text-center font-bold">{r.totalQualified}</td>
                  <td className="px-3 py-3 text-center">{r.totalTarget ?? "–"}</td>
                  <td className="px-4 py-3">{r.status === "GAP" ? <Badge tone="red">Gap</Badge> : r.status === "COVERED" ? <Badge tone="blue">Covered</Badge> : <Badge>No target</Badge>}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </Card>
      )}
    </div>
  );
}
