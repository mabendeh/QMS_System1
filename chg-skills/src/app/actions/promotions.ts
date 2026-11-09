"use server";
import { revalidatePath } from "next/cache";
import { z } from "zod";
import type { Prisma } from "@prisma/client";
import { prisma } from "@/lib/db";
import { AuthError, requireUser } from "@/lib/auth";
import { audit } from "@/lib/audit";
import { can, type Permission } from "@/lib/rbac";
import { loadPlant } from "@/lib/data/plant";
import { nextPromotionStatus, STEP_FOR_STATUS, type PromotionStatus } from "@/lib/domain/promotion";
import type { ActionResult } from "./skills";

function fail(e: unknown): ActionResult {
  if (e instanceof AuthError) return { ok: false, error: e.message };
  if (e instanceof z.ZodError) return { ok: false, error: e.errors.map((x) => x.message).join(" ") };
  console.error(e);
  return { ok: false, error: "Something went wrong. Nothing was saved." };
}

const submitSchema = z.object({
  employeeId: z.string().min(1),
  toLevelId: z.string().min(1),
  justification: z.string().trim().min(20, "Explain the recommendation in at least 20 characters (what the employee has demonstrated)."),
});

export async function submitPromotion(form: FormData): Promise<ActionResult & { id?: string }> {
  try {
    const u = await requireUser("promotion:submit");
    const input = submitSchema.parse(Object.fromEntries(form.entries()));
    const emp = await prisma.employee.findUniqueOrThrow({ where: { id: input.employeeId }, include: { careerLevel: true } });
    if (!u.roles.includes("PRODUCTION_MANAGER") && emp.supervisorUserId !== u.id) {
      return { ok: false, error: "Supervisors can recommend only employees assigned to them." };
    }
    const target = await prisma.careerLevel.findUniqueOrThrow({ where: { id: input.toLevelId } });
    if (target.rank <= (emp.careerLevel?.rank ?? -1)) return { ok: false, error: "The target level must be above the employee's current level." };
    const open = await prisma.promotionRequest.findFirst({ where: { employeeId: emp.id, status: { in: ["SUBMITTED", "MANAGER_APPROVED", "HR_VERIFIED"] } } });
    if (open) return { ok: false, error: "This employee already has a promotion request in review." };
    const plant = await loadPlant({ employeeIds: [emp.id] });
    const ev = plant.evaluateFor(emp.id, target.id);
    if (!ev) return { ok: false, error: `No promotion rules are configured for ${target.name}. HR must set them up first.` };
    if (!ev.result.readyForReview) {
      return { ok: false, error: `Not all documented requirements for ${target.name} are met yet.`, reasons: ev.result.missing };
    }
    const snapshot = { evaluatedAt: new Date().toISOString(), ruleVersion: ev.version, items: ev.result.items, metCount: ev.result.metCount, total: ev.result.total };
    const req = await prisma.$transaction(async (tx) => {
      const r = await tx.promotionRequest.create({
        data: {
          employeeId: emp.id, fromLevelId: emp.careerLevelId, toLevelId: target.id, ruleSetId: ev.ruleSetId,
          justification: input.justification, evaluationSnapshot: snapshot as unknown as Prisma.InputJsonValue, submittedById: u.id,
        },
      });
      await audit(tx, u, "promotion.submitted", "PromotionRequest", r.id, `${u.name} recommended ${emp.firstName} ${emp.lastName} for ${target.name}`, undefined, { toLevel: target.name, ruleVersion: ev.version });
      return r;
    });
    revalidatePath("/promotions");
    return { ok: true, message: "Recommendation submitted for Production Manager review.", id: req.id };
  } catch (e) {
    return fail(e);
  }
}

const decideSchema = z.object({
  id: z.string().min(1),
  expectedVersion: z.coerce.number().int(),
  decision: z.enum(["APPROVED", "REJECTED"]),
  comments: z.string().max(4000).optional().transform((v) => v?.trim() || null),
  effectiveDate: z.string().optional().transform((v) => (v ? new Date(`${v}T12:00:00Z`) : null)),
});

export async function decidePromotion(form: FormData): Promise<ActionResult> {
  try {
    const u = await requireUser("promotion:view");
    const input = decideSchema.parse(Object.fromEntries(form.entries()));
    const r = await prisma.promotionRequest.findUniqueOrThrow({ where: { id: input.id }, include: { employee: true, toLevel: true, approvals: true } });
    const step = STEP_FOR_STATUS[r.status as PromotionStatus];
    if (!step) return { ok: false, error: "This request is already closed." };
    if (!can(u.roles, step.permission as Permission)) return { ok: false, error: `Only an authorized ${step.label.toLowerCase()} reviewer can decide this step.` };
    if (r.submittedById === u.id) return { ok: false, error: "You submitted this recommendation, so you cannot also review it." };
    if (r.approvals.some((a) => a.decidedById === u.id)) return { ok: false, error: "You already decided an earlier step. A different person must decide this one." };
    if (input.decision === "REJECTED" && !input.comments) return { ok: false, error: "A reason is required when rejecting." };
    const next = nextPromotionStatus(r.status as PromotionStatus, input.decision)!;
    const final = next === "APPROVED";
    if (final && !input.effectiveDate) return { ok: false, error: "Enter the effective date for the approved career level." };
    const now = new Date();
    const ok = await prisma.$transaction(async (tx) => {
      const upd = await tx.promotionRequest.updateMany({
        where: { id: r.id, version: input.expectedVersion, status: r.status },
        data: {
          status: next, version: { increment: 1 },
          ...(next === "APPROVED" || next === "REJECTED" ? { decisionDate: now } : {}),
          ...(final ? { effectiveDate: input.effectiveDate } : {}),
        },
      });
      if (upd.count === 0) return false;
      await tx.approvalRecord.create({ data: { promotionRequestId: r.id, step: step.step, decision: input.decision, comments: input.comments, decidedById: u.id } });
      if (final) {
        // Records the approved career level. This system never changes pay.
        await tx.employee.update({ where: { id: r.employeeId }, data: { careerLevelId: r.toLevelId } });
      }
      await audit(tx, u, `promotion.${input.decision === "APPROVED" ? "approved" : "rejected"}`, "PromotionRequest", r.id,
        `${u.name} ${input.decision === "APPROVED" ? "approved" : "rejected"} ${step.label.toLowerCase()} for ${r.employee.firstName} ${r.employee.lastName} → ${r.toLevel.name}`,
        { status: r.status }, { status: next, comments: input.comments, effectiveDate: input.effectiveDate });
      return true;
    });
    if (!ok) return { ok: false, conflict: true, error: "Someone else acted on this request first. The page has been refreshed." };
    revalidatePath("/promotions");
    revalidatePath(`/promotions/${r.id}`);
    return {
      ok: true,
      message: final
        ? `Approved. Career level recorded as ${r.toLevel.name} effective ${input.effectiveDate!.toLocaleDateString("en-US")}. Any pay change must be processed by HR through the approved payroll process.`
        : input.decision === "REJECTED" ? "Rejected. The reason and decision history are kept." : `Approved – sent to ${STEP_FOR_STATUS[next]?.label ?? "next step"}.`,
    };
  } catch (e) {
    return fail(e);
  }
}

export async function withdrawPromotion(form: FormData): Promise<ActionResult> {
  try {
    const u = await requireUser("promotion:submit");
    const id = String(form.get("id") ?? "");
    const reason = String(form.get("comments") ?? "").trim();
    const r = await prisma.promotionRequest.findUniqueOrThrow({ where: { id }, include: { employee: true } });
    if (r.submittedById !== u.id) return { ok: false, error: "Only the person who submitted the request can withdraw it." };
    if (!["SUBMITTED", "MANAGER_APPROVED", "HR_VERIFIED"].includes(r.status)) return { ok: false, error: "This request is already closed." };
    if (reason.length < 5) return { ok: false, error: "Give a short reason for withdrawing." };
    await prisma.$transaction(async (tx) => {
      await tx.promotionRequest.update({ where: { id }, data: { status: "WITHDRAWN", decisionDate: new Date(), version: { increment: 1 } } });
      await audit(tx, u, "promotion.withdrawn", "PromotionRequest", id, `${u.name} withdrew the promotion request for ${r.employee.firstName} ${r.employee.lastName}`, { status: r.status }, { status: "WITHDRAWN", reason });
    });
    revalidatePath("/promotions");
    return { ok: true, message: "Request withdrawn." };
  } catch (e) {
    return fail(e);
  }
}
