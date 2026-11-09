import Link from "next/link";
import { notFound, redirect } from "next/navigation";
import { CheckCircle2, Circle, Download, FileText, Lock } from "lucide-react";
import { pageUser } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { loadPlant } from "@/lib/data/plant";
import { canViewEmployee } from "@/lib/scope";
import { STATUS_META } from "@/lib/domain/status";
import { Badge, Card, CardHeader, SampleBadge, Table, Td, Th } from "@/components/ui";

import { ProfileSkills, type ProfileArea } from "./ProfileSkills";
import { cn, fmtDate } from "@/lib/utils";

export const metadata = { title: "Employee profile" };

export default async function EmployeeProfile({ params }: { params: Promise<{ id: string }> }) {
  const u = await pageUser();
  const { id } = await params;
  if (!canViewEmployee(u, id)) redirect("/denied");
  const plant = await loadPlant({ includeInactiveEmployees: true, employeeIds: [id] });
  const e = plant.employees[0];
  if (!e) notFound();
  const s = plant.employeeSummary(e.id);
  const next = plant.nextLevel(e.id);
  const isSelf = u.employeeId === e.id && !u.permissions.includes("employee:view_all");
  const canComp = u.permissions.includes("comp:view");

  const [evidence, trainings, promotions, payGrade, lastAssessors] = await Promise.all([
    prisma.evidence.groupBy({ by: ["competencyId"], where: { employeeId: e.id }, _count: true }),
    prisma.trainingAssignment.findMany({ where: { employeeId: e.id }, include: { competency: { include: { area: true } } }, orderBy: [{ status: "asc" }, { dueDate: "asc" }] }),
    prisma.promotionRequest.findMany({ where: { employeeId: e.id }, include: { toLevel: true, fromLevel: true }, orderBy: { submittedAt: "desc" } }),
    canComp && e.careerLevelId ? prisma.payGrade.findFirst({ where: { careerLevelId: e.careerLevelId, isActive: true } }) : null,
    prisma.employeeSkill.findMany({ where: { employeeId: e.id }, select: { competencyId: true, lastAssessedAt: true, lastAssessedById: true, qualifiedAt: true, expiresAt: true } }),
  ]);
  const userNames = new Map((await prisma.user.findMany({ where: { id: { in: lastAssessors.map((l) => l.lastAssessedById).filter(Boolean) as string[] } }, select: { id: true, name: true } })).map((x) => [x.id, x.name]));
  const evCount = new Map(evidence.map((x) => [x.competencyId, x._count]));
  const meta = new Map(lastAssessors.map((l) => [l.competencyId, l]));
  const openTraining = new Map(trainings.filter((t) => t.status !== "COMPLETED" && t.status !== "CANCELLED").map((t) => [t.competencyId, t]));

  const areas: ProfileArea[] = plant.areas
    .map((a) => ({
      id: a.id, name: a.name, kind: a.kind, primary: e.primaryAreaId === a.id, qualified: s.qualifiedAreaIds.has(a.id),
      rows: a.competencies
        .map((c) => {
          const r = plant.rec(e.id, c.id);
          const m = meta.get(c.id);
          const st = plant.display(e.id, c.id);
          const t = openTraining.get(c.id);
          return {
            id: c.id, name: c.name, critical: c.isCritical, status: st, pending: !!r?.pendingStatus, required: plant.required(e.id, c.id),
            assessedAt: m?.lastAssessedAt?.toISOString() ?? null, evaluator: m?.lastAssessedById ? userNames.get(m.lastAssessedById) ?? null : null,
            evidence: evCount.get(c.id) ?? 0, qualifiedAt: m?.qualifiedAt?.toISOString() ?? null, expiresAt: m?.expiresAt?.toISOString() ?? null,
            need: st === "COMPETENT" ? null : t ? `Training ${t.status === "PRACTICAL_PASSED" ? "– practical passed" : "assigned"}${t.dueDate ? `, due ${fmtDate(t.dueDate)}` : ""}` : plant.required(e.id, c.id) ? (st === "EXPIRED" ? "Requalification needed" : "Training needed") : null,
          };
        })
        .filter((r) => r.required || r.status !== "NOT_ASSESSED"),
    }))
    .filter((a) => a.rows.length > 0);

  return (
    <div className="grid gap-6">
      <div className="flex flex-wrap items-start justify-between gap-4">
        <div>
          <p className="text-sm font-semibold text-muted">{isSelf ? "My qualifications" : <Link href="/employees" className="hover:underline">Employees</Link>}</p>
          <h1 className="mt-1 flex flex-wrap items-center gap-3 text-2xl font-bold">{e.firstName} {e.lastName} {e.isSample && <SampleBadge />} {e.status === "INACTIVE" && <Badge>Inactive</Badge>}</h1>
        </div>
        <div className="no-print flex flex-wrap gap-2">
          <a href={`/api/export/employee?id=${e.id}&format=pdf`} className="inline-flex h-10 items-center gap-2 rounded-md border border-line bg-white px-4 text-sm font-semibold hover:bg-canvas"><FileText className="h-4 w-4" /> PDF report</a>
          <a href={`/api/export/employee?id=${e.id}&format=xlsx`} className="inline-flex h-10 items-center gap-2 rounded-md border border-line bg-white px-4 text-sm font-semibold hover:bg-canvas"><Download className="h-4 w-4" /> Excel</a>
          {u.permissions.includes("promotion:submit") && next?.ruleSetId && (
            <Link href={`/promotions/new?employee=${e.id}`} className="inline-flex h-10 items-center rounded-md bg-navy-800 px-4 text-sm font-semibold text-white hover:bg-navy-700">Recommend for promotion</Link>
          )}
        </div>
      </div>

      <Card>
        <dl className="grid grid-cols-2 gap-x-6 gap-y-4 p-5 text-sm md:grid-cols-4 xl:grid-cols-5">
          <Info k="Employee ID" v={<span className="font-mono">{e.employeeNumber}</span>} />
          <Info k="Department" v={e.primaryArea?.name ?? "—"} />
          <Info k="Shift" v={e.shift.name} />
          <Info k="Job title" v={e.jobRole?.title ?? "—"} />
          <Info k="Hire date" v={fmtDate(e.hireDate)} />
          <Info k="Supervisor" v={e.supervisor?.name ?? "—"} />
          <Info k="Current skill grade" v={e.careerLevel ? `Level ${e.careerLevel.rank} – ${e.careerLevel.name}` : "—"} />
          <Info k="Qualification status" v={s.fullyQualified ? <Badge tone="blue">Fully qualified for job</Badge> : <Badge tone="red">{s.requiredCount - s.competentRequired} job skill(s) to complete</Badge>} />
          <Info k="Production areas qualified" v={`${s.productionAreasQualified}${s.crossTrained ? " · cross-trained" : ""}`} />
          {canComp && (
            <Info k={<span className="inline-flex items-center gap-1"><Lock className="h-3 w-3" /> Pay grade (HR only)</span>} v={payGrade ? `${payGrade.code} · ${payGrade.minRate && payGrade.maxRate ? `$${payGrade.minRate} – $${payGrade.maxRate}` : "range not configured"}` : "Not configured"} />
          )}
        </dl>
      </Card>

      <div className="grid gap-6 xl:grid-cols-[1fr_380px]">
        <ProfileSkills employeeId={e.id} areas={areas} />

        <div className="grid content-start gap-6">
          <Card>
            <CardHeader title="Career progression" subtitle="Level titles and rules are configured by HR" />
            <div className="p-5">
              <ol className="grid gap-2">
                {plant.levels.map((l) => {
                  const cur = e.careerLevel?.rank === l.rank;
                  const done = (e.careerLevel?.rank ?? -1) > l.rank;
                  return (
                    <li key={l.id} className={cn("flex items-center gap-3 rounded-md px-3 py-2 text-sm", cur && "bg-navy-800 text-white", next?.level.id === l.id && "ring-2 ring-navy-200")}>
                      {done || cur ? <CheckCircle2 className="h-5 w-5 shrink-0" /> : <Circle className="h-5 w-5 shrink-0 text-muted" />}
                      <span className="font-semibold">L{l.rank} {l.name}</span>
                      {cur && <span className="ml-auto text-xs">Current</span>}
                      {next?.level.id === l.id && <span className="ml-auto text-xs font-semibold text-navy-700">Next</span>}
                    </li>
                  );
                })}
              </ol>
              {next && next.result ? (
                <div className="mt-5 border-t border-line pt-4">
                  <p className="text-sm text-muted">To reach <b className="text-ink">{next.level.name}</b> (rules v{next.version}):</p>
                  <ul className="mt-3 grid gap-2 text-sm">
                    {next.result.items.map((it) => (
                      <li key={it.id} className="flex gap-2">
                        {it.met ? <CheckCircle2 className="mt-0.5 h-4 w-4 shrink-0 text-st-blue" aria-label="Met" /> : <Circle className={cn("mt-0.5 h-4 w-4 shrink-0", it.manual ? "text-muted" : "text-st-red")} aria-label={it.manual ? "Confirmed in review" : "Not met"} />}
                        <span><b>{it.label}</b><br /><span className="text-muted">{it.detail}</span></span>
                      </li>
                    ))}
                  </ul>
                  <div className="mt-4">
                    <div className="flex justify-between text-sm"><span className="font-semibold">Progress</span><span>{next.result.metCount} of {next.result.total} requirements</span></div>
                    <div className="mt-1 h-2 rounded bg-canvas"><div className="h-2 rounded bg-st-blue" style={{ width: `${(next.result.metCount / next.result.total) * 100}%` }} /></div>
                    <p className="mt-2 text-xs text-muted">
                      {next.result.readyForReview ? "All system-checked requirements are met. Remaining items are confirmed during the promotion review." : "Time employed is not counted – only verified competencies."}
                    </p>
                  </div>
                </div>
              ) : <p className="mt-4 text-sm text-muted">{next ? "No promotion rules are configured for the next level yet." : "Highest level reached."}</p>}
            </div>
          </Card>

          <Card>
            <CardHeader title="Training" />
            {trainings.length ? (
              <ul className="divide-y divide-line text-sm">
                {trainings.slice(0, 10).map((t) => (
                  <li key={t.id} className="px-5 py-3">
                    <p className="font-semibold">{t.competency.name} <span className="font-normal text-muted">· {t.competency.area.name}</span></p>
                    <p className="text-muted">{t.status.replace(/_/g, " ").toLowerCase()} · TWI: {t.twiStep.replace("_", " ").toLowerCase()}{t.dueDate ? ` · due ${fmtDate(t.dueDate)}` : ""}</p>
                  </li>
                ))}
              </ul>
            ) : <p className="px-5 py-6 text-sm text-muted">No training assignments.</p>}
          </Card>

          {!isSelf && promotions.length > 0 && (
            <Card>
              <CardHeader title="Promotion history" />
              <ul className="divide-y divide-line text-sm">
                {promotions.map((p) => (
                  <li key={p.id} className="px-5 py-3">
                    <Link href={`/promotions/${p.id}`} className="font-semibold text-navy-700 hover:underline">{p.fromLevel?.name ?? "—"} → {p.toLevel.name}</Link>
                    <p className="text-muted">{p.status.replace(/_/g, " ").toLowerCase()} · submitted {fmtDate(p.submittedAt)}</p>
                  </li>
                ))}
              </ul>
            </Card>
          )}
        </div>
      </div>
      <p className="text-xs text-muted">Status colours: {Object.values(STATUS_META).map((m) => `${m.short} = ${m.label}`).join(" · ")}</p>
    </div>
  );
}

function Info({ k, v }: { k: React.ReactNode; v: React.ReactNode }) {
  return (
    <div>
      <dt className="text-xs font-bold uppercase tracking-wide text-muted">{k}</dt>
      <dd className="mt-1 font-medium">{v}</dd>
    </div>
  );
}

