import Link from "next/link";
import { pageUser } from "@/lib/auth";
import { loadPlant } from "@/lib/data/plant";
import { computeAnalytics } from "@/lib/data/analytics";
import { Alert, Badge, Card, CardHeader, Empty, PageHeader, Stat, Table, Td, Th } from "@/components/ui";
import { Charts } from "./Charts";
import { pct } from "@/lib/utils";

export const metadata = { title: "CI Analytics" };

export default async function AnalyticsPage() {
  await pageUser("analytics:view");
  const plant = await loadPlant();
  const a = await computeAnalytics(plant);
  const t = a.trend;
  const last = t[t.length - 1];
  const prev = t[t.length - 2];
  const delta = last && prev ? last.activeQualifications - prev.activeQualifications : 0;
  const totalAssigned = a.byArea.reduce((n, r) => n + r.assigned, 0);
  const totalAssignedQ = a.byArea.reduce((n, r) => n + r.assignedQualified, 0);
  const open = (a.pipeline.byStatus.SUBMITTED ?? 0) + (a.pipeline.byStatus.MANAGER_APPROVED ?? 0) + (a.pipeline.byStatus.HR_VERIFIED ?? 0);
  return (
    <div className="grid gap-5">
      <PageHeader
        title="Continuous improvement analytics"
        subtitle="Skill coverage, cross-training and training progress – counted from qualification records only."
        actions={<>
          <a href="/api/export/monthly?format=xlsx" className="inline-flex h-10 items-center rounded-md bg-navy-800 px-4 text-sm font-semibold text-white">Monthly report (Excel)</a>
          <a href="/api/export/monthly?format=pdf" className="inline-flex h-10 items-center rounded-md border border-line bg-white px-4 text-sm font-semibold hover:bg-canvas">PDF</a>
        </>}
      />
      <section className="grid grid-cols-2 gap-3 md:grid-cols-3 xl:grid-cols-6">
        <Stat label="Assigned dept. qualified" value={`${pct(totalAssignedQ, totalAssigned)}%`} hint={`${totalAssignedQ} of ${totalAssigned} people qualified in their own department`} tone="blue" />
        <Stat label="Cross-trained" value={`${pct(a.cross.crossTrained, a.cross.total)}%`} hint={`${a.cross.crossTrained} of ${a.cross.total} qualified in 2+ areas`} tone="green" />
        <Stat label="Active qualifications" value={last?.activeQualifications ?? 0} hint={`${delta >= 0 ? "▲" : "▼"} ${Math.abs(delta)} vs last month`} />
        <Stat label="Single-person skills" value={a.single.length} tone={a.single.length ? "red" : "gray"} hint="Only one qualified person" />
        <Stat label="No one qualified" value={a.none.length} tone={a.none.length ? "red" : "gray"} hint="Required by a job, nobody Competent" />
        <Stat label="Promotion pipeline" value={open} hint={`${a.pipeline.readyForReview} meet next-level requirements`} href="/promotions" />
      </section>
      <Charts byArea={a.byArea} dist={a.cross.dist} trend={a.trend} />
      <div className="grid gap-5 xl:grid-cols-2">
        <Card>
          <CardHeader title="Department training priorities" subtitle="Shift gaps first, then single-person dependencies, then open skill gaps" />
          {a.priorities.length === 0 ? <Empty title="No priorities – all departments covered" /> : (
            <Table>
              <thead><tr><Th>Department</Th><Th>Shift gaps</Th><Th>Single-person skills</Th><Th>Open required-skill gaps</Th></tr></thead>
              <tbody>{a.priorities.slice(0, 10).map((p) => (
                <tr key={p.id}><Td className="font-semibold">{p.name}</Td><Td>{p.shiftGaps ? <Badge tone="red">{p.shiftGaps}</Badge> : 0}</Td><Td>{p.single || 0}</Td><Td>{p.gaps}</Td></tr>
              ))}</tbody>
            </Table>
          )}
        </Card>
        <Card>
          <CardHeader title="Critical single-person dependencies" subtitle="Competencies only one person can currently perform" />
          {a.single.length + a.none.length === 0 ? <Empty title="None" /> : (
            <Table>
              <thead><tr><Th>Competency</Th><Th>Department</Th><Th>Only qualified person</Th></tr></thead>
              <tbody>
                {a.none.slice(0, 10).map((n) => <tr key={n.id}><Td>{n.name}</Td><Td>{n.area}</Td><Td><Badge tone="red">No one</Badge></Td></tr>)}
                {a.single.slice(0, 15).map((s) => <tr key={s.id}><Td>{s.name}</Td><Td>{s.area}</Td><Td>{s.holder} · Shift {s.shift}</Td></tr>)}
              </tbody>
            </Table>
          )}
        </Card>
      </div>
      <Card>
        <CardHeader title="Training effectiveness" subtitle="Measured only from recorded evaluations and completions" />
        <div className="grid gap-4 p-5 text-sm md:grid-cols-3">
          <div><p className="text-xs font-bold uppercase text-muted">First-attempt practical pass rate</p><p className="text-2xl font-bold">{a.effectiveness.firstAttemptPassRate == null ? "—" : `${a.effectiveness.firstAttemptPassRate}%`}</p><p className="text-muted">{a.effectiveness.evaluations} employee-skill evaluations</p></div>
          <div><p className="text-xs font-bold uppercase text-muted">Average days, start to completion</p><p className="text-2xl font-bold">{a.effectiveness.avgDaysToComplete ?? "—"}</p><p className="text-muted">{a.effectiveness.completedWithDates} completed assignments with dates</p></div>
          <div><p className="text-xs font-bold uppercase text-muted">Productivity / cost impact</p><p className="text-2xl font-bold text-muted">Not measured</p><p className="text-muted">Requires recorded baseline and post-training measurements. No savings are estimated.</p></div>
        </div>
      </Card>
      <Alert tone="info">Need the detail behind a number? Open the <Link className="font-semibold underline" href="/matrix">skills matrix</Link> or the <Link className="font-semibold underline" href="/coverage">coverage table</Link>.</Alert>
    </div>
  );
}
