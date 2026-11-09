import Link from "next/link";
import { notFound, redirect } from "next/navigation";
import { CheckCircle2, Circle, XCircle } from "lucide-react";
import { pageUser } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { loadPlant } from "@/lib/data/plant";
import { STEP_FOR_STATUS, type PromotionStatus } from "@/lib/domain/promotion";
import { STATUS_TEXT, STATUS_TONE } from "@/lib/promotionUi";
import { Alert, Badge, Card, CardHeader, SampleBadge } from "@/components/ui";
import { DecisionPanel } from "./DecisionPanel";
import { fmtDate, fmtDateTime } from "@/lib/utils";

export const metadata = { title: "Promotion request" };

type SnapItem = { id: string; label: string; met: boolean; manual: boolean; detail: string };

export default async function PromotionDetail({ params }: { params: Promise<{ id: string }> }) {
  const u = await pageUser("promotion:view");
  const { id } = await params;
  const r = await prisma.promotionRequest.findUnique({
    where: { id },
    include: { employee: { include: { shift: true, jobRole: true } }, fromLevel: true, toLevel: true, ruleSet: true, approvals: { orderBy: { decidedAt: "asc" } } },
  });
  if (!r) notFound();
  const supervisorOnly = u.roles.every((x) => x === "SUPERVISOR");
  if (supervisorOnly && r.submittedById !== u.id && r.employee.supervisorUserId !== u.id) redirect("/denied");
  const users = new Map((await prisma.user.findMany({ select: { id: true, name: true } })).map((x) => [x.id, x.name]));
  const plant = await loadPlant({ includeInactiveEmployees: true, employeeIds: [r.employeeId] });
  const live = plant.evaluateFor(r.employeeId, r.toLevelId);
  const snap = r.evaluationSnapshot as { items?: SnapItem[]; ruleVersion?: number; evaluatedAt?: string };
  const step = STEP_FOR_STATUS[r.status as PromotionStatus];
  const canDecide = !!step && u.permissions.includes(step.permission as never) && r.submittedById !== u.id && !r.approvals.some((a) => a.decidedById === u.id);
  const steps = [
    { key: "SUBMIT", label: "Supervisor recommendation", who: users.get(r.submittedById), at: r.submittedAt, decision: "APPROVED" as const, comments: r.justification },
    ...(["MANAGER_REVIEW", "HR_VERIFICATION", "FINAL_APPROVAL"] as const).map((s) => {
      const a = r.approvals.find((x) => x.step === s);
      return { key: s, label: s === "MANAGER_REVIEW" ? "Production Manager review" : s === "HR_VERIFICATION" ? "HR qualification & pay-policy check" : "Final approval", who: a ? users.get(a.decidedById) : undefined, at: a?.decidedAt, decision: a?.decision, comments: a?.comments };
    }),
  ];

  return (
    <div className="grid gap-5">
      <div>
        <Link href="/promotions" className="text-sm font-semibold text-muted hover:underline">Promotion review</Link>
        <h1 className="mt-1 flex flex-wrap items-center gap-3 text-2xl font-bold">
          {r.employee.firstName} {r.employee.lastName} {r.employee.isSample && <SampleBadge />}
          <Badge tone={STATUS_TONE[r.status]}>{STATUS_TEXT[r.status]}</Badge>
        </h1>
        <p className="mt-1 text-sm text-muted">{r.employee.employeeNumber} · Shift {r.employee.shift.code} · {r.employee.jobRole?.title} · {r.fromLevel?.name ?? "—"} → <b className="text-ink">{r.toLevel.name}</b> · rules v{r.ruleSet.version}</p>
      </div>

      <div className="grid gap-5 xl:grid-cols-[1fr_400px]">
        <div className="grid content-start gap-5">
          <Card>
            <CardHeader title="Requirements" subtitle={snap.evaluatedAt ? `At submission (${fmtDateTime(snap.evaluatedAt)}) and now` : "Checked live against current qualifications"} />
            <table className="w-full text-sm">
              <thead><tr className="text-left text-xs font-bold uppercase text-muted"><th className="px-5 py-2">Requirement</th><th className="px-2 py-2">At submission</th><th className="px-5 py-2">Now</th></tr></thead>
              <tbody>
                {(snap.items?.length ? snap.items : (live?.result.items ?? []).map((x) => ({ ...x, unknown: true }))).map((it: SnapItem & { unknown?: boolean }) => {
                  const now = live?.result.items.find((x) => x.id === it.id);
                  return (
                    <tr key={it.id} className="border-t border-line">
                      <td className="px-5 py-2.5"><b>{it.label}</b><div className="text-muted">{now?.detail ?? it.detail}</div></td>
                      <td className="px-2 py-2.5">{it.unknown ? <span className="text-muted">Not recorded</span> : <Mark met={it.met} manual={it.manual} />}</td>
                      <td className="px-5 py-2.5">{now ? <Mark met={now.met} manual={now.manual} /> : <span className="text-muted">Rules changed</span>}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
            {live && !live.result.readyForReview && ["SUBMITTED", "MANAGER_APPROVED", "HR_VERIFIED"].includes(r.status) && (
              <div className="p-4"><Alert tone="warn">A requirement that was met at submission is no longer met (for example an expired qualification). Reviewers should take this into account.</Alert></div>
            )}
          </Card>
          <Card>
            <CardHeader title="Recommendation" subtitle={`By ${users.get(r.submittedById)} on ${fmtDate(r.submittedAt)}`} />
            <p className="whitespace-pre-wrap p-5 text-sm">{r.justification}</p>
          </Card>
        </div>

        <div className="grid content-start gap-5">
          <Card>
            <CardHeader title="Decision history" subtitle="Permanent record – cannot be edited" />
            <ol className="grid gap-4 p-5">
              {steps.map((s) => (
                <li key={s.key} className="flex gap-3">
                  {s.decision === "APPROVED" ? <CheckCircle2 className="h-5 w-5 shrink-0 text-st-blue" /> : s.decision === "REJECTED" ? <XCircle className="h-5 w-5 shrink-0 text-st-red" /> : <Circle className="h-5 w-5 shrink-0 text-muted" />}
                  <div className="text-sm">
                    <p className="font-semibold">{s.label}</p>
                    {s.who ? (
                      <p className="text-muted">{s.key === "SUBMIT" ? "Submitted" : s.decision === "APPROVED" ? "Approved" : "Rejected"} by {s.who} · {fmtDateTime(s.at)}</p>
                    ) : <p className="text-muted">{r.status === "REJECTED" || r.status === "WITHDRAWN" ? "Not reached" : "Pending"}</p>}
                    {s.key !== "SUBMIT" && s.comments && <p className="mt-1 rounded bg-canvas p-2">“{s.comments}”</p>}
                  </div>
                </li>
              ))}
            </ol>
            {(r.decisionDate || r.effectiveDate) && (
              <dl className="grid grid-cols-2 gap-3 border-t border-line p-5 text-sm">
                <div><dt className="text-xs font-bold uppercase text-muted">Decision date</dt><dd>{fmtDate(r.decisionDate)}</dd></div>
                <div><dt className="text-xs font-bold uppercase text-muted">Effective date</dt><dd>{fmtDate(r.effectiveDate)}</dd></div>
              </dl>
            )}
          </Card>
          {(canDecide || (r.submittedById === u.id && step)) && (
            <DecisionPanel id={r.id} version={r.version} stepLabel={step?.label ?? ""} canDecide={canDecide} isFinal={r.status === "HR_VERIFIED"} canWithdraw={r.submittedById === u.id && !!step} />
          )}
          {step && !canDecide && r.submittedById !== u.id && (
            <Alert tone="info">Waiting for: {step.label}.</Alert>
          )}
        </div>
      </div>
    </div>
  );
}

function Mark({ met, manual }: { met: boolean; manual: boolean }) {
  if (manual) return <Badge>In review</Badge>;
  return met ? <Badge tone="blue">Met</Badge> : <Badge tone="red">Not met</Badge>;
}
