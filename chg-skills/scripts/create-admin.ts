/**
 * Creates (or resets) an administrator account without any hardcoded password.
 * Usage: ADMIN_EMAIL=you@chg.example ADMIN_NAME="Your Name" ADMIN_ROLES=SYSTEM_ADMIN,HR_ADMIN npm run admin:create
 * The password is read from ADMIN_PASSWORD, or generated and printed once.
 */
import { PrismaClient, type Role } from "@prisma/client";
import bcrypt from "bcryptjs";
import { randomBytes } from "node:crypto";

const prisma = new PrismaClient();
async function main() {
  const email = process.env.ADMIN_EMAIL?.trim().toLowerCase();
  const name = process.env.ADMIN_NAME?.trim() || "Administrator";
  const roles = (process.env.ADMIN_ROLES || "SYSTEM_ADMIN").split(",").map((r) => r.trim()) as Role[];
  if (!email) throw new Error("Set ADMIN_EMAIL");
  let pw = process.env.ADMIN_PASSWORD;
  const generated = !pw;
  if (!pw) pw = randomBytes(12).toString("base64url");
  if (pw.length < 12) throw new Error("ADMIN_PASSWORD must be at least 12 characters");
  const hash = await bcrypt.hash(pw, 12);
  const u = await prisma.user.upsert({ where: { email }, create: { email, name, passwordHash: hash }, update: { passwordHash: hash, isActive: true } });
  for (const role of roles) await prisma.userRole.upsert({ where: { userId_role: { userId: u.id, role } }, create: { userId: u.id, role }, update: {} });
  await prisma.auditLog.create({ data: { action: "users.admin_bootstrap", entityType: "User", entityId: u.id, summary: `Admin account ${email} created/reset from command line (${roles.join(", ")})` } });
  console.log(`Admin ready: ${email} (${roles.join(", ")})`);
  if (generated) console.log(`Generated password (shown once – change it after first sign-in): ${pw}`);
}
main().catch((e) => { console.error(e.message); process.exit(1); }).finally(() => prisma.$disconnect());
