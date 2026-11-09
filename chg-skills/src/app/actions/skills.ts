"use server";
import { revalidatePath } from "next/cache";
import { z } from "zod";
import { Prisma, type SkillStatus } from "@prisma/client";
import { prisma } from "@/lib/db";
import { AuthError, requireUser, type CurrentUser } from "@/lib/auth";
import { audit } from "@/lib/audit";
import { canAssessEmployee, canEvaluatePractical, canViewEmployee } from "@/lib/scope";
import { storeUpload, UploadError } from "@/lib/uploads";
import { requalificationMonths } from "@/lib/expiry";
import { computeExpiry, decideAssessment, displayStatus, STORED_LABEL, verificationBlockers } from "@/lib/domain/status";

export type ActionResult =
  | { ok: true; message: string; outcome?: "APPLIED" | "PENDING" | "VERIFIED" | "REJECTED" }
  | { ok: false; error: string; reasons?: string[]; conflict?: boolean };

class Conflict extends Error {}

function fail(e: unknown): ActionResult {
  if (e instanceof AuthError) return { ok: false, error: e.message };
  if (e instanceof UploadError) return { ok: false, error: e.message };
  if (e instanceof Conflict) return { ok: false, conflict: true, error: e.message };
  if (e instanceof z.ZodError) return { ok: false, error: e.errors.map((x) => x.message).join(" ") };
  console.error(e);
  return { ok: false, error: "Something went wrong saving this change. Nothing was saved. Please try again." };
}

const date = z
  .string()
  .optional()
  .transform((v) => (v ? new Date(`${v}T12:00:00Z`) : null))
  .refine((d) => !d || (!isNaN(d.getTime()) && d.getTime() <= Date.now() + 86_400_000), "Dates cannot be in the future.");

// ------------------------------------------------------------------ read: cell detail

export async function getCellDetail(employeeId: string, competencyId: string) {
  const u = await requireUser();
  if (!canViewEmployee(u, employeeId)) throw new AuthError("FORBIDDEN", "You cannot view this employee.");
  const [emp, comp, rec, practical, evidence, history] = await Promise.all([
    prisma.employee.findUniqueOrThrow({ where: { id: employeeId }, include: { jobRole: { include: { requirements: { where: { competencyId } } } } } }),
    prisma.competency.findUniqueOrThrow({ where: { id: competencyId }, include: { area: true } }),
    prisma.employeeSkill.findUnique({ where: { employeeId_competencyId: { employeeId, competencyId } } }),
    prisma.practicalEvaluation.findFirst({ where: { employeeId, competencyId }, orderBy: { evaluatedAt: "desc" } }),
    prisma.evidence.findMany({ where: { employeeId, competencyId }, orderBy: { createdAt: "desc" } }),
    prisma.skillAssessment.findMany({ where: { employeeId, competencyId }, orderBy: { createdAt: "desc" }, take: 8 }),
  ]);
  const userIds = [...new Set([...history.map((h) => h.submittedById), practical?.evaluatorUserId, rec?.lastAssessedById].filter(Boolean) as string[])];
  const names = new Map((await prisma.user.findMany({ where: { id: { in: userIds } }, select: { id: true, name: true } })).map((x) => [x.id, x.name]));
  const pending = rec?.pendingAssessmentId ? history.find((h) => h.id === rec.pendingAssessmentId) ?? null : null;
  const required = (emp.jobRole?.requirements.length ?? 0) > 0;
  const canAssess = await canAssessEmployee(u, emp, comp.areaId);
  const canEval = await canEvaluatePractical(u, comp.areaId);
  const canVerify = u.permissions.includes("qualification:verify");
  return {
    employee: { id: emp.id, name: `${emp.firstName} ${emp.lastName}`, number: emp.employeeNumber, isSample: emp.isSample },
    competency: { id: comp.id, code: comp.code, name: comp.name, description: comp.description, isCritical: comp.isCritical, requiresEvidence: comp.requiresEvidence, area: comp.area.name },
    record: rec && {
      status: rec.status, pendingStatus: rec.pendingStatus, version: rec.version, trainingDate: rec.trainingDate, qualifiedAt: rec.qualifiedAt,
      expiresAt: rec.expiresAt, lastAssessedAt: rec.lastAssessedAt, lastAssessedBy: rec.lastAssessedById ? names.get(rec.lastAssessedById) ?? null : null, notes: rec.notes,
    },
    display: displayStatus(rec, required),
    required,
    practical: practical && { result: practical.result, independentOperation: practical.independentOperation, evaluatedAt: practical.evaluatedAt, evaluator: names.get(practical.evaluatorUserId) ?? "Unknown", notes: practical.notes },
    evidence: evidence.map((e) => ({ id: e.id, name: e.originalName, at: e.createdAt })),
    history: history.map((h) => ({ id: h.id, at: h.createdAt, requested: h.requestedStatus, state: h.state, by: names.get(h.submittedById) ?? "Unknown", notes: h.notes, evaluatorName: h.evaluatorName })),
    pending: pending && { id: pending.id, requested: pending.requestedStatus, by: names.get(pending.submittedById) ?? "Unknown", at: pending.createdAt, notes: pending.notes },
    permissions: { canAssess, canEval, canVerify },
    actorName: u.name,
  };
}

// ------------------------------------------------------------------ write: assessment

const assessSchema = z.object({
  employeeId: z.string().min(1),
  competencyId: z.string().min(1),
  status: z.enum(["NOT_APPLICABLE", "NEEDS_TRAINING", "IN_TRAINING", "TRAINED", "COMPETENT"]),
  expectedVersion: z.coerce.number().int().min(0),
  notes: z.string().max(2000).optional().transform((v) => v?.trim() || null),
  trainingDate: date,
  evaluatorName: z.string().max(120).optional().transform((v) => v?.trim() || null),
});

export async function submitAssessment(form: FormData): Promise<ActionResult> {
  try {
    const u = await requireUser("skill:assess");
    const input = assessSchema.parse(Object.fromEntries([...form.entries()].filter(([, v]) => typeof v === "string")));
    const file = form.get("evidence");
    const emp = await prisma.employee.findUniqueOrThrow({ where: { id: input.employeeId } });
    const comp = await prisma.competency.findUniqueOrThrow({ where: { id: input.competencyId }, include: { area: true } });
    if (!(await canAssessEmployee(u, emp, comp.areaId))) {
      return { ok: false, error: "You can only assess employees assigned to you (or areas you are authorized to train)." };
    }
    const current = await prisma.employeeSkill.findUnique({ where: { employeeId_competencyId: { employeeId: emp.id, competencyId: comp.id } } });
    if ((current?.version ?? 0) !== input.expectedVersion) {
      return conflictResult(current?.lastAssessedById ?? null, current?.updatedAt ?? null);
    }
    const practical = await prisma.practicalEvaluation.findFirst({ where: { employeeId: emp.id, competencyId: comp.id }, orderBy: { evaluatedAt: "desc" } });
    const evidenceCount = (await prisma.evidence.count({ where: { employeeId: emp.id, competencyId: comp.id } })) + (file instanceof File && file.size > 0 ? 1 : 0);
    const decision = decideAssessment({
      requested: input.status, current: current?.status ?? null, trainingDate: input.trainingDate ?? current?.trainingDate ?? null, notes: input.notes,
      competency: comp, practical, evidenceCount, actor: { id: u.id, canVerify: u.permissions.includes("qualification:verify") },
    });
    if (decision.outcome === "BLOCKED") {
      return { ok: false, error: `${comp.name} cannot be marked ${STORED_LABEL[input.status]} yet.`, reasons: decision.reasons };
    }
    const stored = file instanceof File && file.size > 0 ? await storeUpload(file) : null;
    const months = await requalificationMonths(prisma, comp.id, comp.areaId);
    const now = new Date();

    const result = await prisma.$transaction(async (tx) => {
      // Supersede any older pending request for this cell
      if (current?.pendingAssessmentId) {
        await tx.skillAssessment.updateMany({ where: { id: current.pendingAssessmentId, state: "PENDING_VERIFICATION" }, data: { state: "SUPERSEDED", decidedAt: now } });
      }
      const pending = decision.outcome === "PENDING";
      const assessment = await tx.skillAssessment.create({
        data: {
          employeeId: emp.id, competencyId: comp.id, previousStatus: current?.status ?? null, requestedStatus: input.status,
          state: pending ? "PENDING_VERIFICATION" : decision.verified ? "VERIFIED" : "APPLIED",
          notes: input.notes, trainingDate: input.trainingDate, evaluatorName: input.evaluatorName ?? u.name, submittedById: u.id,
          practicalEvaluationId: input.status === "COMPETENT" ? practical?.id ?? null : null, decidedAt: pending ? null : now,
        },
      });
      if (stored) await tx.evidence.create({ data: { ...stored, employeeId: emp.id, competencyId: comp.id, assessmentId: assessment.id, uploadedById: u.id } });
      if (!pending && decision.verified) {
        await tx.approvalRecord.create({ data: { assessmentId: assessment.id, step: "QUALIFICATION_VERIFICATION", decision: "APPROVED", decidedById: u.id, comments: "Verified at assessment" } });
      }
      const newStatus: SkillStatus = pending ? current?.status ?? "TRAINED" : input.status;
      const becameCompetent = !pending && input.status === "COMPETENT";
      const data = {
        status: newStatus,
        pendingStatus: pending ? input.status : null,
        pendingAssessmentId: pending ? assessment.id : null,
        trainingDate: input.trainingDate ?? current?.trainingDate ?? null,
        qualifiedAt: becameCompetent ? now : newStatus === "COMPETENT" ? current?.qualifiedAt ?? null : null,
        expiresAt: becameCompetent ? computeExpiry(now, months) : newStatus === "COMPETENT" ? current?.expiresAt ?? null : null,
        lastAssessedAt: now,
        lastAssessedById: u.id,
        notes: input.notes ?? current?.notes ?? null,
      };
      if (current) {
        const upd = await tx.employeeSkill.updateMany({ where: { id: current.id, version: input.expectedVersion }, data: { ...data, version: { increment: 1 } } });
        if (upd.count === 0) throw new Conflict("Someone else changed this skill a moment ago. Your change was not saved – review the latest status and try again.");
      } else {
        await tx.employeeSkill.create({ data: { ...data, employeeId: emp.id, competencyId: comp.id } });
      }
      if (becameCompetent) await completeTraining(tx, emp.id, comp.id, now);
      await audit(tx, u, pending ? "skill.assessment_pending" : "skill.assessment", "EmployeeSkill", `${emp.id}:${comp.id}`,
        `${u.name} ${pending ? "requested" : "set"} ${emp.firstName} ${emp.lastName} – ${comp.name} → ${STORED_LABEL[input.status]}${pending ? " (pending verification)" : ""}`,
        { status: current?.status ?? null }, { status: newStatus, pendingStatus: data.pendingStatus, notes: input.notes, evidence: stored?.originalName });
      return pending
        ? ({ ok: true, outcome: "PENDING", message: (decision as { reason: string }).reason } as const)
        : ({ ok: true, outcome: decision.verified ? "VERIFIED" : "APPLIED", message: `${comp.name} updated to ${STORED_LABEL[input.status]}.` } as const);
    });
    revalidatePath("/", "layout");
    return result;
  } catch (e) {
    if (e instanceof Prisma.PrismaClientKnownRequestError && e.code === "P2002") {
      return { ok: false, conflict: true, error: "Someone else saved this skill at the same moment. Your change was not saved – review and try again." };
    }
    return fail(e);
  }
}

async function conflictResult(byId: string | null, at: Date | null): Promise<ActionResult> {
  const who = byId ? (await prisma.user.findUnique({ where: { id: byId }, select: { name: true } }))?.name : null;
  return {
    ok: false, conflict: true,
    error: `This skill was updated${who ? ` by ${who}` : ""}${at ? ` at ${at.toLocaleTimeString("en-US", { timeZone: "America/Chicago" })}` : ""} after you opened it. Your change was not saved. The latest status is now shown – review it and submit again if needed.`,
  };
}

async function completeTraining(tx: Prisma.TransactionClient, employeeId: string, competencyId: string, now: Date) {
  await tx.trainingAssignment.updateMany({
    where: { employeeId, competencyId, status: { in: ["ASSIGNED", "IN_PROGRESS", "PRACTICAL_PASSED"] } },
    data: { status: "COMPLETED", completedDate: now, twiStep: "FOLLOW_UP", version: { increment: 1 } },
  });
}

// ------------------------------------------------------------------ write: practical evaluation

const practicalSchema = z.object({
  employeeId: z.string().min(1),
  competencyId: z.string().min(1),
  result: z.enum(["PASS", "FAIL"]),
  independentOperation: z.string().optional().transform((v) => v === "on" || v === "true"),
  evaluatedAt: date,
  notes: z.string().max(2000).optional().transform((v) => v?.trim() || null),
  trainingAssignmentId: z.string().optional().transform((v) => v || null),
});

export async function recordPracticalEvaluation(form: FormData): Promise<ActionResult> {
  try {
    const u = await requireUser("training:evaluate");
    const input = practicalSchema.parse(Object.fromEntries([...form.entries()].filter(([, v]) => typeof v === "string")));
    const comp = await prisma.competency.findUniqueOrThrow({ where: { id: input.competencyId } });
    const emp = await prisma.employee.findUniqueOrThrow({ where: { id: input.employeeId } });
    if (!(await canEvaluatePractical(u, comp.areaId))) return { ok: false, error: "You are not an authorized trainer for this area." };
    if (input.result === "PASS" && !input.independentOperation) {
      return { ok: false, error: "A pass must confirm the employee operated independently. Otherwise record it as not passed.", reasons: ["Tick 'Operated independently' or choose 'Not passed'."] };
    }
    const file = form.get("evidence");
    const stored = file instanceof File && file.size > 0 ? await storeUpload(file) : null;
    await prisma.$transaction(async (tx) => {
      const pe = await tx.practicalEvaluation.create({
        data: {
          employeeId: emp.id, competencyId: comp.id, evaluatorUserId: u.id, evaluatedAt: input.evaluatedAt ?? new Date(), result: input.result,
          independentOperation: input.independentOperation, notes: input.notes, trainingAssignmentId: input.trainingAssignmentId,
        },
      });
      if (stored) await tx.evidence.create({ data: { ...stored, employeeId: emp.id, competencyId: comp.id, practicalEvaluationId: pe.id, trainingAssignmentId: input.trainingAssignmentId, uploadedById: u.id } });
      if (input.result === "PASS") {
        await tx.trainingAssignment.updateMany({
          where: { employeeId: emp.id, competencyId: comp.id, status: { in: ["ASSIGNED", "IN_PROGRESS"] } },
          data: { status: "PRACTICAL_PASSED", twiStep: "FOLLOW_UP", version: { increment: 1 } },
        });
      }
      await audit(tx, u, "training.practical_evaluation", "PracticalEvaluation", pe.id,
        `${u.name} recorded a ${input.result === "PASS" ? "passed" : "failed"} practical evaluation for ${emp.firstName} ${emp.lastName} – ${comp.name}`,
        undefined, { result: input.result, independentOperation: input.independentOperation, evidence: stored?.originalName });
    });
    revalidatePath("/", "layout");
    return { ok: true, message: input.result === "PASS" ? "Practical evaluation recorded. The skill can now be verified as Competent by an authorized evaluator." : "Practical evaluation recorded as not passed." };
  } catch (e) {
    return fail(e);
  }
}

// ------------------------------------------------------------------ write: verification decision

const verifySchema = z.object({
  assessmentId: z.string().min(1),
  decision: z.enum(["APPROVED", "REJECTED"]),
  comments: z.string().max(2000).optional().transform((v) => v?.trim() || null),
});

export async function decideVerification(form: FormData): Promise<ActionResult> {
  try {
    const u = await requireUser("qualification:verify");
    const input = verifySchema.parse(Object.fromEntries(form.entries()));
    if (input.decision === "REJECTED" && !input.comments) return { ok: false, error: "Add a comment explaining why the qualification is not approved." };
    return await verify(u, input.assessmentId, input.decision, input.comments);
  } catch (e) {
    return fail(e);
  }
}

async function verify(u: CurrentUser, assessmentId: string, decision: "APPROVED" | "REJECTED", comments: string | null): Promise<ActionResult> {
  const a = await prisma.skillAssessment.findUniqueOrThrow({ where: { id: assessmentId }, include: { competency: true, employee: true } });
  if (a.state !== "PENDING_VERIFICATION") return { ok: false, conflict: true, error: "This request has already been decided by someone else." };
  const practical = a.practicalEvaluationId ? await prisma.practicalEvaluation.findUnique({ where: { id: a.practicalEvaluationId } }) : null;
  const blockers = verificationBlockers({ verifierId: u.id, verifierCanVerify: true, submittedById: a.submittedById, practicalEvaluatorId: practical?.evaluatorUserId ?? null });
  if (decision === "APPROVED") {
    const evidenceCount = await prisma.evidence.count({ where: { employeeId: a.employeeId, competencyId: a.competencyId } });
    const d = decideAssessment({
      requested: "COMPETENT", current: null, competency: a.competency, practical, evidenceCount, actor: { id: u.id, canVerify: true },
    });
    if (d.outcome === "BLOCKED") blockers.push(...d.reasons);
  }
  if (blockers.length) return { ok: false, error: "This qualification cannot be approved yet.", reasons: blockers };
  const months = await requalificationMonths(prisma, a.competencyId, a.competency.areaId);
  const now = new Date();
  await prisma.$transaction(async (tx) => {
    const claimed = await tx.skillAssessment.updateMany({ where: { id: a.id, state: "PENDING_VERIFICATION" }, data: { state: decision === "APPROVED" ? "VERIFIED" : "REJECTED", decidedAt: now } });
    if (claimed.count === 0) throw new Conflict("This request has already been decided by someone else.");
    await tx.approvalRecord.create({ data: { assessmentId: a.id, step: "QUALIFICATION_VERIFICATION", decision, comments, decidedById: u.id } });
    const rec = await tx.employeeSkill.findUniqueOrThrow({ where: { employeeId_competencyId: { employeeId: a.employeeId, competencyId: a.competencyId } } });
    await tx.employeeSkill.update({
      where: { id: rec.id },
      data: decision === "APPROVED"
        ? { status: "COMPETENT", qualifiedAt: now, expiresAt: computeExpiry(now, months), pendingStatus: null, pendingAssessmentId: null, version: { increment: 1 }, lastAssessedAt: now, lastAssessedById: u.id }
        : { pendingStatus: null, pendingAssessmentId: null, version: { increment: 1 } },
    });
    if (decision === "APPROVED") await completeTraining(tx, a.employeeId, a.competencyId, now);
    await audit(tx, u, decision === "APPROVED" ? "skill.verified" : "skill.verification_rejected", "EmployeeSkill", `${a.employeeId}:${a.competencyId}`,
      `${u.name} ${decision === "APPROVED" ? "verified" : "rejected"} Competent for ${a.employee.firstName} ${a.employee.lastName} – ${a.competency.name}`,
      { status: rec.status, pendingStatus: rec.pendingStatus }, { decision, comments });
  });
  revalidatePath("/", "layout");
  return decision === "APPROVED"
    ? { ok: true, outcome: "VERIFIED", message: `${a.competency.name} verified. ${a.employee.firstName} is now Competent.` }
    : { ok: true, outcome: "REJECTED", message: "Request rejected. The previous status is kept and the reason is recorded." };
}
