"use server";
import { revalidatePath } from "next/cache";
import { prisma } from "@/lib/db";
import { AuthError, requireUser } from "@/lib/auth";
import { audit } from "@/lib/audit";
import type { ActionResult } from "./skills";

export async function saveTargets(form: FormData): Promise<ActionResult> {
  try {
    const u = await requireUser("coverage:manage");
    const entries = [...form.entries()].filter(([k]) => k.startsWith("t:"));
    const changes: { areaId: string; shiftId: string; target: number | null }[] = [];
    for (const [k, v] of entries) {
      const [, areaId, shiftId] = k.split(":");
      const s = String(v).trim();
      if (s === "") { changes.push({ areaId, shiftId, target: null }); continue; }
      const n = Number(s);
      if (!Number.isInteger(n) || n < 0 || n > 200) return { ok: false, error: "Targets must be whole numbers between 0 and 200." };
      changes.push({ areaId, shiftId, target: n });
    }
    await prisma.$transaction(async (tx) => {
      const before = await tx.shiftCoverageTarget.findMany();
      for (const c of changes) {
        if (c.target == null) await tx.shiftCoverageTarget.deleteMany({ where: { areaId: c.areaId, shiftId: c.shiftId } });
        else await tx.shiftCoverageTarget.upsert({ where: { areaId_shiftId: { areaId: c.areaId, shiftId: c.shiftId } }, create: { areaId: c.areaId, shiftId: c.shiftId, targetQualified: c.target }, update: { targetQualified: c.target } });
      }
      await audit(tx, u, "coverage.targets_updated", "ShiftCoverageTarget", null, `${u.name} updated shift staffing targets`, before.map((b) => ({ a: b.areaId, s: b.shiftId, t: b.targetQualified })), changes);
    });
    revalidatePath("/coverage");
    return { ok: true, message: "Staffing targets saved." };
  } catch (e) {
    if (e instanceof AuthError) return { ok: false, error: e.message };
    console.error(e);
    return { ok: false, error: "Could not save targets." };
  }
}
