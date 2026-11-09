import "server-only";
import type { Tx } from "./db";

/** Competency-level rule wins over area-level rule; null = does not expire. */
export async function requalificationMonths(tx: Tx, competencyId: string, areaId: string): Promise<number | null> {
  const rules = await tx.qualificationExpirationRule.findMany({
    where: { OR: [{ competencyId }, { areaId, competencyId: null }] },
  });
  return rules.find((r) => r.competencyId === competencyId)?.months ?? rules.find((r) => r.areaId === areaId)?.months ?? null;
}
