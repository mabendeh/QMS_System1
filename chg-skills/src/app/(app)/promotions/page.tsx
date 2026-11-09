import Link from "next/link";
import { pageUser } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { loadPlant } from "@/lib/data/plant";
import { STEP_FOR_STATUS, type PromotionStatus } from "@/lib/domain/promotion";
import { Alert, Badge, Card, CardHeader, Empty, PageHeader, SampleBadge, Table, Td, Th } from "@/components/ui";
import { cn, fmtDate } from "@/lib/utils";
import { STATUS_TEXT, STATUS_TONE } from "@/lib/promotionUi";

export const metadata = { title: "Promotion Review" };

export default async function PromotionsPage({ searchParams }: { searchParams: Promise<{ view?: string }> }) {
  const u = await pageUser("promotion:view");
  const { view = "open" } = await searchParams;
  const isSupervisorOnly = u.roles.every((r) => r === "SUPERVISOR");
  const where = {
    ...(view === "open" ? { status: { in: ["SUBMITTED", "MANAGER_APPROVED", "HR_VERIFIED"] as PromotionStatus[] } } : { status: { in: ["APPROVED", "REJECTED", "WITHDRAWN"] as PromotionStatus[] } }),
    ...(isSupervisorOnly ? { OR: [{ submittedById: u.id }, { employee: { supervisorUserId: u.id } }] } : {}),
  };
  const [requests, users] = await Promise.all([
    prisma.promotionRequest.findMany({ where, include: { employee: { include: { shift: true } }, fromLevel: true, toLevel: true, ruleSet: true }, orderBy: { submittedAt: "desc" } }),
    prisma.user.findMany({ select: { id: true, name: true } }),
  ]);
  const names = new Map(users.map((x) => [x.id, x.name]));
  const myTurn = requests.filter((r) => {
    const s = STEP_FOR_STATUS[r.status as PromotionStatus];
    return s && u.permissions.includes(s.permission as never) && r.submittedById !== u.id;
  });

  let ready: { id: string; name: string; isSample: boolean; next: string; shift: string }[] = [];
  if (u.permissions.includes("promotion:submit")) {
    const plant = await loadPlant();
    const openIds = new Set((await prisma.promotionRequest.findMany({ where: { status: { in: ["SUBMITTED", "MANAGER_APPROVED", "HR_VERIFIED"] } }, select: { employeeId: true } })).map((r) => r.employeeId));
    ready = plant.employees
      .filter((e) => !openIds.has(e.id) && (u.roles.includes("PRODUCTION_MANAGER") || e.supervisorUserId === u.id))
      .map((e) => ({ e, n: plant.nextLevel(e.id) }))
      .filter((x) => x.n?.result?.readyForReview)
      .map((x) => ({ id: x.e.id, name: `${x.e.firstName} ${x.e.lastName}`, isSample: x.e.isSample, next: x.n!.level.name, shift: x.e.shift.code }));
  }

  return (
    <div className="grid gap-5">
      <PageHeader
        title="Promotion review"
        subtitle="Skill-based recommendations. Supervisor submits → Production Manager reviews → HR verifies qualification and pay policy → authorized approver decides."
        actions={<a href="/api/export/promotions?format=xlsx" className="inline-flex h-10 items-center rounded-md border border-line bg-white px-4 text-sm font-semibold hover:bg-canvas">Export report</a>}
      />
      <Alert tone="info">This system records recommendations and decisions. It never changes pay. Approved pay changes are processed by HR under the approved compensation policy.</Alert>

      {myTurn.length > 0 && (
        <Card>
          <CardHeader title={`Waiting for you (${myTurn.length})`} />
          <ul className="divide-y divide-line">
            {myTurn.map((r) => (
              <li key={r.id}><Link href={`/promotions/${r.id}`} className="flex flex-wrap items-center justify-between gap-2 px-5 py-3 hover:bg-canvas">
                <span><b>{r.employee.firstName} {r.employee.lastName}</b> {r.employee.isSample && <SampleBadge />} · {r.fromLevel?.name ?? "—"} → {r.toLevel.name}</span>
                <Badge tone={STATUS_TONE[r.status]}>{STEP_FOR_STATUS[r.status as PromotionStatus]?.label}</Badge>
              </Link></li>
            ))}
          </ul>
        </Card>
      )}

      {ready.length > 0 && (
        <Card>
          <CardHeader title="Meets documented requirements – not yet recommended" subtitle="System-checked requirements are met. A supervisor still decides whether to recommend." />
          <ul className="divide-y divide-line">
            {ready.map((r) => (
              <li key={r.id} className="flex flex-wrap items-center justify-between gap-2 px-5 py-3">
                <span><Link href={`/employees/${r.id}`} className="font-semibold text-navy-700 hover:underline">{r.name}</Link> {r.isSample && <SampleBadge />} · Shift {r.shift} · next: {r.next}</span>
                <Link href={`/promotions/new?employee=${r.id}`} className="text-sm font-semibold text-navy-700 underline">Recommend</Link>
              </li>
            ))}
          </ul>
        </Card>
      )}

      <Card>
        <div className="flex gap-1 border-b border-line p-2">
          {[["open", "In review"], ["closed", "Decided"]].map(([k, l]) => (
            <Link key={k} href={`/promotions?view=${k}`} className={cn("rounded-md px-3 py-2 text-sm font-semibold", view === k ? "bg-navy-800 text-white" : "text-navy-700 hover:bg-navy-50")}>{l}</Link>
          ))}
        </div>
        {requests.length === 0 ? <Empty title="No requests" /> : (
          <Table>
            <thead><tr><Th>Employee</Th><Th>From → To</Th><Th>Rules</Th><Th>Submitted</Th><Th>By</Th><Th>Status</Th></tr></thead>
            <tbody>
              {requests.map((r) => (
                <tr key={r.id} className="hover:bg-canvas">
                  <Td><Link href={`/promotions/${r.id}`} className="font-semibold text-navy-700 hover:underline">{r.employee.firstName} {r.employee.lastName}</Link> {r.employee.isSample && <SampleBadge />}<div className="text-xs text-muted">{r.employee.employeeNumber} · Shift {r.employee.shift.code}</div></Td>
                  <Td>{r.fromLevel?.name ?? "—"} → <b>{r.toLevel.name}</b></Td>
                  <Td>v{r.ruleSet.version}</Td>
                  <Td>{fmtDate(r.submittedAt)}</Td>
                  <Td>{names.get(r.submittedById)}</Td>
                  <Td><Badge tone={STATUS_TONE[r.status]}>{STATUS_TEXT[r.status]}</Badge></Td>
                </tr>
              ))}
            </tbody>
          </Table>
        )}
      </Card>
    </div>
  );
}
