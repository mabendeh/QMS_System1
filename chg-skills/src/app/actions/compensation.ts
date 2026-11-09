"use server";
import { revalidatePath } from "next/cache";
import { z } from "zod";
import { Prisma } from "@prisma/client";
import { prisma } from "@/lib/db";
import { AuthError, requireUser } from "@/lib/auth";
import { audit } from "@/lib/audit";
import type { ActionResult } from "./skills";

const money = z.string().optional().transform((v, ctx) => {
  if (!v || !v.trim()) return null;
  const n = Number(v.replace(/[$,]/g, ""));
  if (!isFinite(n) || n < 0 || n > 100000) { ctx.addIssue({ code: "custom", message: "Amounts must be positive numbers." }); return z.NEVER; }
  return new Prisma.Decimal(n.toFixed(2));
});
const optDate = z.string().optional().transform((v) => (v ? new Date(`${v}T12:00:00Z`) : null));
const text = (max: number) => z.string().max(max).optional().transform((v) => v?.trim() || null);

const gradeSchema = z.object({
  id: z.string().optional().transform((v) => v || null),
  code: z.string().trim().min(1, "Code is required.").max(20),
  name: z.string().trim().min(1, "Name is required.").max(100),
  careerLevelId: z.string().optional().transform((v) => v || null),
  minRate: money, maxRate: money,
  effectiveDate: optDate, endDate: optDate,
  approvalAuthority: text(120), notes: text(1000),
  isActive: z.string().optional().transform((v) => v === "on" || v === "true"),
});

function fail(e: unknown): ActionResult {
  if (e instanceof AuthError) return { ok: false, error: e.message };
  if (e instanceof z.ZodError) return { ok: false, error: e.errors.map((x) => x.message).join(" ") };
  if (e instanceof Prisma.PrismaClientKnownRequestError && e.code === "P2002") return { ok: false, error: "That code or name is already used." };
  console.error(e);
  return { ok: false, error: "Something went wrong. Nothing was saved." };
}

export async function savePayGrade(form: FormData): Promise<ActionResult> {
  try {
    const u = await requireUser("comp:manage");
    const d = gradeSchema.parse(Object.fromEntries(form.entries()));
    if (d.minRate && d.maxRate && d.minRate.gt(d.maxRate)) return { ok: false, error: "Minimum cannot be greater than maximum." };
    if ((d.minRate || d.maxRate) && !d.effectiveDate) return { ok: false, error: "Enter an effective date for an approved pay range." };
    if (d.effectiveDate && d.endDate && d.endDate < d.effectiveDate) return { ok: false, error: "End date must be after the effective date." };
    const { id, ...data } = d;
    await prisma.$transaction(async (tx) => {
      const before = id ? await tx.payGrade.findUnique({ where: { id } }) : null;
      const g = id ? await tx.payGrade.update({ where: { id }, data }) : await tx.payGrade.create({ data });
      await audit(tx, u, id ? "comp.paygrade_updated" : "comp.paygrade_created", "PayGrade", g.id, `${u.name} ${id ? "updated" : "created"} pay grade ${g.code}`, before, g);
    });
    revalidatePath("/hr/compensation");
    return { ok: true, message: "Pay grade saved." };
  } catch (e) {
    return fail(e);
  }
}

const premiumSchema = z.object({
  id: z.string().optional().transform((v) => v || null),
  name: z.string().trim().min(1, "Name is required.").max(100),
  areaId: z.string().optional().transform((v) => v || null),
  amount: money,
  unit: z.string().trim().max(30).default("per hour"),
  effectiveDate: optDate, endDate: optDate,
  approvalAuthority: text(120), notes: text(1000),
  isActive: z.string().optional().transform((v) => v === "on" || v === "true"),
});

export async function saveSkillPremium(form: FormData): Promise<ActionResult> {
  try {
    const u = await requireUser("comp:manage");
    const d = premiumSchema.parse(Object.fromEntries(form.entries()));
    if (d.amount && !d.effectiveDate) return { ok: false, error: "Enter an effective date for an approved premium." };
    const { id, ...data } = d;
    await prisma.$transaction(async (tx) => {
      const before = id ? await tx.skillPremium.findUnique({ where: { id } }) : null;
      const p = id ? await tx.skillPremium.update({ where: { id }, data }) : await tx.skillPremium.create({ data });
      await audit(tx, u, id ? "comp.premium_updated" : "comp.premium_created", "SkillPremium", p.id, `${u.name} ${id ? "updated" : "created"} skill premium ${p.name}`, before, p);
    });
    revalidatePath("/hr/compensation");
    return { ok: true, message: "Skill premium saved." };
  } catch (e) {
    return fail(e);
  }
}
