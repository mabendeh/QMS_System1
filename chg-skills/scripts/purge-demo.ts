/** Deletes all fictional sample employees and their records. Usage: CONFIRM=DELETE npm run demo:purge */
import { PrismaClient } from "@prisma/client";

const prisma = new PrismaClient();
async function main() {
  if (process.env.CONFIRM !== "DELETE") throw new Error("Refusing to run. Set CONFIRM=DELETE to delete all sample data.");
  const n = await prisma.$transaction(async (tx) => {
    const ids = (await tx.employee.findMany({ where: { isSample: true }, select: { id: true } })).map((e) => e.id);
    const reqIds = (await tx.promotionRequest.findMany({ where: { employeeId: { in: ids } }, select: { id: true } })).map((r) => r.id);
    const assessIds = (await tx.skillAssessment.findMany({ where: { employeeId: { in: ids } }, select: { id: true } })).map((r) => r.id);
    await tx.$executeRawUnsafe(`SET LOCAL chg.allow_history_purge = 'on'`);
    await tx.approvalRecord.deleteMany({ where: { OR: [{ promotionRequestId: { in: reqIds } }, { assessmentId: { in: assessIds } }] } });
    await tx.$executeRawUnsafe(`SET LOCAL chg.allow_history_purge = 'off'`);
    await tx.promotionRequest.deleteMany({ where: { id: { in: reqIds } } });
    await tx.user.updateMany({ where: { email: { endsWith: "@demo.local" } }, data: { isActive: false } });
    await tx.user.updateMany({ where: { employeeId: { in: ids } }, data: { employeeId: null } });
    await tx.employee.deleteMany({ where: { id: { in: ids } } });
    await tx.auditLog.create({ data: { action: "system.sample_deleted", entityType: "System", summary: `Command line: deleted ${ids.length} sample employees and disabled @demo.local accounts` } });
    return ids.length;
  }, { timeout: 60_000 });
  console.log(`Deleted ${n} sample employees. All @demo.local accounts are disabled.`);
}
main().catch((e) => { console.error(e.message); process.exit(1); }).finally(() => prisma.$disconnect());
