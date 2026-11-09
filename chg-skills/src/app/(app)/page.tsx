import Link from "next/link";
import { redirect } from "next/navigation";
import { Grid3x3, ClipboardCheck, GraduationCap, ShieldCheck } from "lucide-react";
import { pageUser } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { loadPlant } from "@/lib/data/plant";
import { coverageFor } from "@/lib/data/analytics";
import { Card, CardHeader, PageHeader, Stat, Table, Td, Th, Badge, SampleBadge } from "@/components/ui";
import { StatusLegend } from "@/components/StatusCell";
import { fmtDate } from "@/lib/utils";

export const metadata = { title: "Home" };

export default async function Home() {
  const u = await pageUser();
  if (!u.permissions.includes("matrix:view")) {
    if (u.employeeId) redirect(`/employees/${u.employeeId}`);
    redirect("/denied");
  }
  const plant = await loadPlant();
  const sums = plant.employees.map((e) => ({ e, s: plant.employeeSummary(e.id) }));
  const [pendingPromotions, pendingVerifications, overdue] = await Promise.all([
    prisma.promotionRequest.count({ where: { status: { in: ["SUBMITTED", "MANAGER_APPROVED", "HR_VERIFIED"] } } }),
    prisma.skillAssessment.count({ where: { state: "PENDING_VERIFICATION" } }),
    prisma.trainingAssignment.count({ where: { status: { in: ["ASSIGNED", "IN_PROGRESS"] }, dueDate: { lt: new Date() } } }),
  ]);
  const expiring = sums.reduce((n, x) => n + x.s.expiringSoon, 0);
  const { rows } = await coverageFor(plant);
  const gaps = rows.filter((r) => r.status === "GAP");
  const soon = plant.employees
    .flatMap((e) => [...(plant.skillMap.get(e.id)?.values() ?? [])].filter((r) => r.status === "COMPETENT" && r.expiresAt && r.expiresAt > plant.now && r.expiresAt.getTime() < plant.now.getTime() + 30 * 86_400_000).map((r) => ({ e, r })))
    .sort((a, b) => a.r.expiresAt!.getTime() - b.r.expiresAt!.getTime())
    .slice(0, 8);
  const compById = new Map(plant.allComps.map((c) => [c.id, c]));
  const recById = new Map<string, string>();
  for (const [, m] of plant.skillMap) for (const [cid, r] of m) recById.set(r.id, cid);

  return (
    <div className="grid gap-6">
      <PageHeader
        title={`Good ${greeting()}, ${u.name.split(" ")[0] === "Demo" ? u.name.replace(/^Demo /, "") : u.name.split(" ")[0]}`}
        subtitle="Production skills at a glance. Every number links to the detail behind it."
        actions={
          <>
            <Link href="/matrix" className="inline-flex h-10 items-center gap-2 rounded-md bg-navy-800 px-4 text-sm font-semibold text-white hover:bg-navy-700"><Grid3x3 className="h-4 w-4" /> Open skills matrix</Link>
            {u.permissions.includes("skill:assess") && <Link href="/assess" className="inline-flex h-10 items-center gap-2 rounded-md border border-line bg-white px-4 text-sm font-semibold hover:bg-canvas"><ClipboardCheck className="h-4 w-4" /> Quick assessment</Link>}
          </>
        }
      />
      <section className="grid grid-cols-2 gap-3 md:grid-cols-4 xl:grid-cols-7" aria-label="Key figures">
        <Stat label="Total employees" value={plant.employees.length} href="/employees" />
        <Stat label="Fully qualified" value={sums.filter((x) => x.s.fullyQualified).length} tone="blue" hint="All job skills Competent" href="/matrix" />
        <Stat label="In training" value={sums.filter((x) => x.s.inTraining > 0).length} tone="yellow" href="/training" />
        <Stat label="Need training" value={sums.filter((x) => x.s.needsTraining > 0).length} tone="red" hint="Has a required skill gap" href="/matrix" />
        <Stat label="Cross-trained" value={sums.filter((x) => x.s.crossTrained).length} tone="green" hint="Qualified in 2+ production areas" href="/analytics" />
        <Stat label="Pending promotion review" value={pendingPromotions} href="/promotions" />
        <Stat label="Expiring in 30 days" value={expiring} tone={expiring ? "red" : "gray"} hint="Qualifications to renew" href="/training" />
      </section>

      <div className="grid gap-6 xl:grid-cols-2">
        <Card>
          <CardHeader title="Needs attention" subtitle="Work waiting on someone" />
          <ul className="divide-y divide-line">
            <AttentionRow href="/training?tab=verify" icon={<ShieldCheck className="h-5 w-5" />} label="Qualifications waiting for verification" n={pendingVerifications} />
            <AttentionRow href="/training?tab=overdue" icon={<GraduationCap className="h-5 w-5" />} label="Training assignments overdue" n={overdue} />
            <AttentionRow href="/coverage" icon={<Grid3x3 className="h-5 w-5" />} label="Departments below shift staffing target" n={gaps.length} />
            <AttentionRow href="/promotions" icon={<ClipboardCheck className="h-5 w-5" />} label="Promotion requests in review" n={pendingPromotions} />
          </ul>
        </Card>
        <Card>
          <CardHeader title="Qualifications expiring soon" subtitle="Next 30 days" />
          {soon.length ? (
            <Table>
              <thead><tr><Th>Employee</Th><Th>Competency</Th><Th>Expires</Th></tr></thead>
              <tbody>
                {soon.map(({ e, r }) => {
                  const c = compById.get(recById.get(r.id) ?? "");
                  return (
                    <tr key={r.id}>
                      <Td><Link className="font-semibold text-navy-700 hover:underline" href={`/employees/${e.id}`}>{e.firstName} {e.lastName}</Link> {e.isSample && <SampleBadge />}</Td>
                      <Td>{c?.name} <span className="text-muted">· {c?.areaName}</span></Td>
                      <Td><Badge tone="red">{fmtDate(r.expiresAt)}</Badge></Td>
                    </tr>
                  );
                })}
              </tbody>
            </Table>
          ) : <p className="px-5 py-8 text-center text-sm text-muted">Nothing expires in the next 30 days.</p>}
        </Card>
      </div>
      <Card className="p-4"><StatusLegend /></Card>
    </div>
  );
}

function AttentionRow({ href, icon, label, n }: { href: string; icon: React.ReactNode; label: string; n: number }) {
  return (
    <li>
      <Link href={href} className="flex items-center gap-3 px-5 py-3.5 hover:bg-canvas">
        <span className="text-navy-700">{icon}</span>
        <span className="flex-1 font-medium">{label}</span>
        <span className={`min-w-10 rounded-full px-2.5 py-0.5 text-center text-sm font-bold ${n ? "bg-red-50 text-st-red" : "bg-gray-100 text-muted"}`}>{n}</span>
      </Link>
    </li>
  );
}

function greeting() {
  const h = Number(new Date().toLocaleString("en-US", { hour: "numeric", hour12: false, timeZone: "America/Chicago" }));
  return h < 12 ? "morning" : h < 17 ? "afternoon" : "evening";
}
