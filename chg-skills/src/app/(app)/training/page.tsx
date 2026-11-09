import Link from "next/link";
import { pageUser } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { loadPlant } from "@/lib/data/plant";
import { isExpiringSoon } from "@/lib/domain/status";
import { Card, PageHeader, Stat } from "@/components/ui";
import { TrainingClient, type TrainingData } from "./TrainingClient";
import { cn, fmtDate } from "@/lib/utils";

export const metadata = { title: "Training" };
const TABS = [
  ["active", "In progress"], ["verify", "Pending approval"], ["overdue", "Overdue"], ["requal", "Requalification due"],
  ["gaps", "Gaps to assign"], ["done", "Completed"], ["assign", "Assign training"],
] as const;

export default async function TrainingPage({ searchParams }: { searchParams: Promise<{ tab?: string }> }) {
  const u = await pageUser("training:view");
  const { tab = "active" } = await searchParams;
  const plant = await loadPlant();
  const now = new Date();
  const [assignments, pendingAssess, users, trainerAuths, evals] = await Promise.all([
    prisma.trainingAssignment.findMany({
      where: { employee: { status: "ACTIVE" } },
      include: { employee: { include: { shift: true } }, competency: { include: { area: true } } },
      orderBy: [{ dueDate: "asc" }],
    }),
    prisma.skillAssessment.findMany({
      where: { state: "PENDING_VERIFICATION" },
      include: { employee: true, competency: { include: { area: true } } },
      orderBy: { createdAt: "asc" },
    }),
    prisma.user.findMany({ select: { id: true, name: true } }),
    prisma.trainerAuthorization.findMany({ include: { user: { select: { id: true, name: true } } } }),
    prisma.practicalEvaluation.findMany({ select: { employeeId: true, competencyId: true, result: true }, orderBy: { evaluatedAt: "asc" } }),
  ]);
  const names = new Map(users.map((x) => [x.id, x.name]));
  const open = assignments.filter((a) => ["ASSIGNED", "IN_PROGRESS", "PRACTICAL_PASSED"].includes(a.status));
  const overdue = open.filter((a) => a.dueDate && a.dueDate < now);
  const done = assignments.filter((a) => a.status === "COMPLETED").sort((a, b) => (b.completedDate?.getTime() ?? 0) - (a.completedDate?.getTime() ?? 0));
  const openKey = new Set(open.map((a) => `${a.employeeId}:${a.competencyId}`));

  const requal: TrainingData["requal"] = [];
  const gaps: TrainingData["gaps"] = [];
  for (const e of plant.employees) {
    for (const c of plant.allComps) {
      const r = plant.rec(e.id, c.id);
      const st = plant.display(e.id, c.id);
      if (st === "EXPIRED" || isExpiringSoon(r, 30, now)) {
        requal.push({ employeeId: e.id, employee: `${e.firstName} ${e.lastName}`, isSample: e.isSample, shift: e.shift.code, competencyId: c.id, competency: c.name, area: c.areaName, expiresAt: r?.expiresAt?.toISOString() ?? null, expired: st === "EXPIRED", assigned: openKey.has(`${e.id}:${c.id}`) });
      } else if (plant.required(e.id, c.id) && (st === "NEEDS_TRAINING") && !openKey.has(`${e.id}:${c.id}`)) {
        gaps.push({ employeeId: e.id, employee: `${e.firstName} ${e.lastName}`, isSample: e.isSample, shift: e.shift.code, competencyId: c.id, competency: c.name, area: c.areaName, critical: c.isCritical });
      }
    }
  }
  requal.sort((a, b) => (a.expiresAt ?? "").localeCompare(b.expiresAt ?? ""));
  gaps.sort((a, b) => Number(b.critical) - Number(a.critical) || a.area.localeCompare(b.area));

  const first = new Map<string, string>();
  for (const ev of evals) { const k = `${ev.employeeId}:${ev.competencyId}`; if (!first.has(k)) first.set(k, ev.result); }
  const fa = [...first.values()];
  const durations = done.filter((d) => d.startDate && d.completedDate).map((d) => (d.completedDate!.getTime() - d.startDate!.getTime()) / 86_400_000);
  const ninety = new Date(now.getTime() - 90 * 86_400_000);

  const row = (a: (typeof assignments)[number]) => ({
    id: a.id, version: a.version, employeeId: a.employeeId, employee: `${a.employee.firstName} ${a.employee.lastName}`, isSample: a.employee.isSample, shift: a.employee.shift.code,
    competencyId: a.competencyId, competency: a.competency.name, area: a.competency.area.name, status: a.status, twiStep: a.twiStep,
    trainer: a.trainerUserId ? names.get(a.trainerUserId) ?? "—" : "Unassigned", due: a.dueDate?.toISOString() ?? null,
    overdue: !!a.dueDate && a.dueDate < now && a.status !== "COMPLETED", completed: a.completedDate?.toISOString() ?? null,
  });

  const data: TrainingData = {
    tab,
    canManage: u.permissions.includes("training:manage"),
    active: open.map(row),
    overdue: overdue.map(row),
    done: done.slice(0, 100).map(row),
    verify: pendingAssess.map((p) => ({
      id: p.id, employeeId: p.employeeId, employee: `${p.employee.firstName} ${p.employee.lastName}`, isSample: p.employee.isSample,
      competencyId: p.competencyId, competency: p.competency.name, area: p.competency.area.name, by: names.get(p.submittedById) ?? "—", at: p.createdAt.toISOString(),
    })),
    requal: requal.slice(0, 200),
    gaps: gaps.slice(0, 200),
    gapsTotal: gaps.length,
    employees: plant.employees.map((e) => ({ id: e.id, name: `${e.firstName} ${e.lastName} (${e.employeeNumber})`, supervisorId: e.supervisorUserId })),
    areas: plant.areas.map((a) => ({ id: a.id, name: a.name, competencies: a.competencies.map((c) => ({ id: c.id, name: c.name })) })),
    trainers: trainerAuths.map((t) => ({ userId: t.userId, name: t.user.name, areaId: t.areaId })),
  };

  return (
    <div>
      <PageHeader
        title="Training management"
        subtitle="Identify gap → assign trainer → train (TWI: Prepare, Present, Try Out, Follow Up) → practical demonstration → verify → matrix updates automatically."
        actions={<>
          <a href="/api/export/gaps?format=xlsx" className="inline-flex h-10 items-center rounded-md border border-line bg-white px-4 text-sm font-semibold hover:bg-canvas">Export gaps (Excel)</a>
          <a href="/api/export/gaps?format=pdf" className="inline-flex h-10 items-center rounded-md border border-line bg-white px-4 text-sm font-semibold hover:bg-canvas">PDF</a>
        </>}
      />
      <section className="mb-5 grid grid-cols-2 gap-3 md:grid-cols-3 xl:grid-cols-6">
        <Stat label="Overdue" value={overdue.length} tone={overdue.length ? "red" : "gray"} href="/training?tab=overdue" />
        <Stat label="In progress" value={open.length} tone="yellow" href="/training?tab=active" />
        <Stat label="Pending approval" value={pendingAssess.length} tone="navy" href="/training?tab=verify" />
        <Stat label="Completed (90 days)" value={done.filter((d) => d.completedDate && d.completedDate > ninety).length} tone="blue" href="/training?tab=done" />
        <Stat label="Requalification due" value={requal.length} tone={requal.length ? "red" : "gray"} hint="Expired or due in 30 days" href="/training?tab=requal" />
        <Stat label="First-try pass rate" value={fa.length ? `${Math.round((fa.filter((r) => r === "PASS").length / fa.length) * 100)}%` : "—"} hint={fa.length ? `${fa.length} practical evaluations${durations.length ? ` · avg ${Math.round(durations.reduce((a, b) => a + b, 0) / durations.length)} days to complete` : ""}` : "No evaluations yet"} />
      </section>
      <nav className="mb-4 flex gap-1 overflow-x-auto" aria-label="Training views">
        {TABS.filter(([k]) => k !== "assign" || data.canManage).map(([k, label]) => (
          <Link key={k} href={`/training?tab=${k}`} aria-current={tab === k ? "page" : undefined}
            className={cn("whitespace-nowrap rounded-md px-3 py-2 text-sm font-semibold", tab === k ? "bg-navy-800 text-white" : "bg-white text-navy-700 ring-1 ring-line hover:bg-navy-50")}>
            {label}
          </Link>
        ))}
      </nav>
      <Card><TrainingClient data={data} /></Card>
      <p className="mt-3 text-xs text-muted">Training effectiveness uses only recorded practical evaluations and completion dates (as of {fmtDate(now)}). Productivity impact is not estimated.</p>
    </div>
  );
}
