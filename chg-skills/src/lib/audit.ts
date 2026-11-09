import type { Prisma } from "@prisma/client";
import type { Tx } from "./db";

export interface Actor {
  id: string;
  name: string;
}

/** Writes one append-only audit row inside the caller's transaction. */
export async function audit(
  tx: Tx,
  actor: Actor | null,
  action: string,
  entityType: string,
  entityId: string | null,
  summary: string,
  before?: unknown,
  after?: unknown,
) {
  await tx.auditLog.create({
    data: {
      actorUserId: actor?.id ?? null,
      actorName: actor?.name ?? "system",
      action,
      entityType,
      entityId,
      summary,
      before: before === undefined ? undefined : (JSON.parse(JSON.stringify(before)) as Prisma.InputJsonValue),
      after: after === undefined ? undefined : (JSON.parse(JSON.stringify(after)) as Prisma.InputJsonValue),
    },
  });
}
