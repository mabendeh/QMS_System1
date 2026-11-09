"use server";
import { revalidatePath } from "next/cache";
import { z } from "zod";
import { prisma } from "@/lib/db";
import { AuthError, requireUser } from "@/lib/auth";
import { audit } from "@/lib/audit";
import type { ActionResult } from "./skills";

const TWI = ["PREPARE", "PRESENT", "TRY_OUT", "FOLLOW_UP"] as const;

function fail(e: unknown): ActionResult {
  if (e instanceof AuthError) return { ok: false, error: e.message };
  if (e instanceof z.ZodError) return { ok: false, error: e.errors.map((x) => x.message).join(" ") };
  console.error(e);
  return { ok: false, error: "Something went wrong. Nothing was saved." };
}

const assignSchema = z.object({
  employeeId: z.string().min(1, "Choose an employee."),
  competencyId: z.string().min(1, "Choose a competency."),
  trainerUserId: z.string().optional().transform((v) => v || null),
  dueDate: z.string().optional().transform((v) => (v ? new Date(`${v}T12:00:00Z`) : null)),
  notes: z.string().max(2000).optional().transform((v) => v?.trim() || null),
});

export async function assignTraining(form: FormData): Promise<ActionResult> {
  try {
    const u = await requireUser("training:manage");
    const input = assignSchema.parse(Object.fromEntries(form.entries()));
    const emp = await prisma.employee.findUniqueOrThrow({ where: { id: input.employeeId } });
    const comp = await prisma.competency.findUniqueOrThrow({ where: { id: input.competencyId }, include: { area: true } });
    const onlyOwn = u.roles.length === 1 && u.roles[0] === "SUPERVISOR";
    if (onlyOwn && emp.supervisorUserId !== u.id) return { ok: false, error: "Supervisors can assign training only to their own employees." };
    if (input.trainerUserId) {
      const auth = await prisma.trainerAuthorization.findUnique({ where: { userId_areaId: { userId: input.trainerUserId, areaId: comp.areaId } } });
      if (!auth) return { ok: false, error: `That trainer is not authorized for ${comp.area.name}.` };
    }
    const open = await prisma.trainingAssignment.findFirst({ where: { employeeId: emp.id, competencyId: comp.id, status: { in: ["ASSIGNED", "IN_PROGRESS", "PRACTICAL_PASSED"] } } });
    if (open) return { ok: false, error: "This employee already has open training for this competency." };
    await prisma.$transaction(async (tx) => {
      const t = await tx.trainingAssignment.create({
        data: { employeeId: emp.id, competencyId: comp.id, trainerUserId: input.trainerUserId, assignedById: u.id, dueDate: input.dueDate, notes: input.notes },
      });
      // Assigning training moves an untouched/red skill to In Training
      const rec = await tx.employeeSkill.findUnique({ where: { employeeId_competencyId: { employeeId: emp.id, competencyId: comp.id } } });
      if (!rec) {
        await tx.employeeSkill.create({ data: { employeeId: emp.id, competencyId: comp.id, status: "IN_TRAINING", lastAssessedAt: new Date(), lastAssessedById: u.id } });
      } else if (rec.status === "NEEDS_TRAINING" || rec.status === "NOT_APPLICABLE") {
        await tx.employeeSkill.update({ where: { id: rec.id }, data: { status: "IN_TRAINING", version: { increment: 1 }, lastAssessedAt: new Date(), lastAssessedById: u.id } });
      }
      await audit(tx, u, "training.assigned", "TrainingAssignment", t.id, `${u.name} assigned ${comp.name} training to ${emp.firstName} ${emp.lastName}`, undefined, input);
    });
    revalidatePath("/", "layout");
    return { ok: true, message: `Training assigned: ${comp.name} for ${emp.firstName} ${emp.lastName}.` };
  } catch (e) {
    return fail(e);
  }
}

const advanceSchema = z.object({
  id: z.string().min(1),
  expectedVersion: z.coerce.number().int(),
  action: z.enum(["start", "next", "cancel"]),
});

export async function advanceTraining(form: FormData): Promise<ActionResult> {
  try {
    const u = await requireUser("training:manage");
    const input = advanceSchema.parse(Object.fromEntries(form.entries()));
    const t = await prisma.trainingAssignment.findUniqueOrThrow({ where: { id: input.id }, include: { employee: true, competency: true } });
    if (t.status === "COMPLETED" || t.status === "CANCELLED") return { ok: false, error: "This training is already closed." };
    let data: Record<string, unknown>;
    let summary: string;
    if (input.action === "cancel") {
      data = { status: "CANCELLED" };
      summary = "cancelled";
    } else if (input.action === "start") {
      data = { status: "IN_PROGRESS", startDate: t.startDate ?? new Date(), twiStep: t.twiStep === "PREPARE" ? "PRESENT" : t.twiStep };
      summary = "started";
    } else {
      const i = TWI.indexOf(t.twiStep);
      if (i >= TWI.length - 1) return { ok: false, error: "Training is at Follow Up. Record the practical evaluation to finish." };
      data = { twiStep: TWI[i + 1], status: t.status === "ASSIGNED" ? "IN_PROGRESS" : t.status, startDate: t.startDate ?? new Date() };
      summary = `moved to ${TWI[i + 1].replace("_", " ").toLowerCase()}`;
    }
    const res = await prisma.$transaction(async (tx) => {
      const upd = await tx.trainingAssignment.updateMany({ where: { id: t.id, version: input.expectedVersion }, data: { ...data, version: { increment: 1 } } });
      if (upd.count === 0) return false;
      await audit(tx, u, "training.updated", "TrainingAssignment", t.id, `${u.name} ${summary} ${t.competency.name} training for ${t.employee.firstName} ${t.employee.lastName}`, { status: t.status, twiStep: t.twiStep }, data);
      return true;
    });
    if (!res) return { ok: false, conflict: true, error: "Someone else updated this training. The list has been refreshed." };
    revalidatePath("/training");
    return { ok: true, message: `Training ${summary}.` };
  } catch (e) {
    return fail(e);
  }
}
