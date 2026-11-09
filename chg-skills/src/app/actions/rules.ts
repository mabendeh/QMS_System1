"use server";
import { revalidatePath } from "next/cache";
import { z } from "zod";
import { prisma } from "@/lib/db";
import { AuthError, requireUser } from "@/lib/auth";
import { audit } from "@/lib/audit";
import type { ActionResult } from "./skills";

function fail(e: unknown): ActionResult {
  if (e instanceof AuthError) return { ok: false, error: e.message };
  if (e instanceof z.ZodError) return { ok: false, error: e.errors.map((x) => x.message).join(" ") };
  console.error(e);
  return { ok: false, error: "Something went wrong. Nothing was saved." };
}

const levelSchema = z.object({
  id: z.string().min(1),
  name: z.string().trim().min(2).max(60),
  description: z.string().trim().max(500).optional(),
  isActive: z.string().optional().transform((v) => v === "on"),
});

export async function saveCareerLevel(form: FormData): Promise<ActionResult> {
  try {
    const u = await requireUser("rules:manage");
    const d = levelSchema.parse(Object.fromEntries(form.entries()));
    await prisma.$transaction(async (tx) => {
      const before = await tx.careerLevel.findUniqueOrThrow({ where: { id: d.id } });
      const after = await tx.careerLevel.update({ where: { id: d.id }, data: { name: d.name, description: d.description || null, isActive: d.isActive } });
      await audit(tx, u, "rules.level_updated", "CareerLevel", d.id, `${u.name} updated career level ${after.name}`, before, after);
    });
    revalidatePath("/admin/rules");
    return { ok: true, message: "Career level saved." };
  } catch (e) {
    return fail(e);
  }
}

const reqSchema = z.object({
  type: z.enum(["PRIMARY_AREA_QUALIFIED", "AREA_QUALIFIED", "MIN_PRODUCTION_AREAS_QUALIFIED", "COMPETENCY_COMPETENT", "CURRENT_LEVEL_AT_LEAST", "MANUAL_VERIFICATION"]),
  label: z.string().trim().min(3, "Each requirement needs a label.").max(200),
  areaId: z.string().nullish().transform((v) => v || null),
  competencyId: z.string().nullish().transform((v) => v || null),
  minCount: z.coerce.number().int().min(0).max(20).nullish(),
});
const ruleSchema = z.object({
  careerLevelId: z.string().min(1),
  notes: z.string().trim().max(1000).optional(),
  requirements: z.array(reqSchema).min(1, "Add at least one requirement."),
});

/** Saving never edits an existing version: it creates version N+1 and deactivates the old one. */
export async function saveRuleVersion(json: string): Promise<ActionResult> {
  try {
    const u = await requireUser("rules:manage");
    const d = ruleSchema.parse(JSON.parse(json));
    for (const r of d.requirements) {
      if (r.type === "AREA_QUALIFIED" && !r.areaId) return { ok: false, error: `“${r.label}” needs a department.` };
      if (r.type === "COMPETENCY_COMPETENT" && !r.competencyId) return { ok: false, error: `“${r.label}” needs a competency.` };
      if ((r.type === "MIN_PRODUCTION_AREAS_QUALIFIED" || r.type === "CURRENT_LEVEL_AT_LEAST") && r.minCount == null) return { ok: false, error: `“${r.label}” needs a number.` };
    }
    const v = await prisma.$transaction(async (tx) => {
      const last = await tx.promotionRuleSet.findFirst({ where: { careerLevelId: d.careerLevelId }, orderBy: { version: "desc" }, include: { requirements: true } });
      await tx.promotionRuleSet.updateMany({ where: { careerLevelId: d.careerLevelId, isActive: true }, data: { isActive: false } });
      const rs = await tx.promotionRuleSet.create({
        data: {
          careerLevelId: d.careerLevelId, version: (last?.version ?? 0) + 1, notes: d.notes || null, createdById: u.id,
          requirements: { create: d.requirements.map((r, i) => ({ ...r, minCount: r.minCount ?? null, sortOrder: i })) },
        },
      });
      await audit(tx, u, "rules.version_created", "PromotionRuleSet", rs.id, `${u.name} published promotion rules v${rs.version}`, last ? { version: last.version, requirements: last.requirements } : null, { version: rs.version, requirements: d.requirements });
      return rs.version;
    });
    revalidatePath("/admin/rules");
    return { ok: true, message: `Rules v${v} published. Earlier versions are kept for requests already submitted.` };
  } catch (e) {
    return fail(e);
  }
}
