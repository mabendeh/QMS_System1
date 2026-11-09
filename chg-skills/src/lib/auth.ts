import "server-only";
import { createHash, randomBytes } from "node:crypto";
import bcrypt from "bcryptjs";
import { cookies, headers } from "next/headers";
import { redirect } from "next/navigation";
import type { Role } from "@prisma/client";
import { prisma } from "./db";
import { can, permissionsFor, type Permission } from "./rbac";

const COOKIE = "chg_session";
const SESSION_HOURS = 12; // roughly one shift
const sha256 = (s: string) => createHash("sha256").update(s).digest("hex");

export interface CurrentUser {
  id: string;
  name: string;
  email: string;
  roles: Role[];
  employeeId: string | null;
  permissions: Permission[];
}

export class AuthError extends Error {
  constructor(public code: "UNAUTHENTICATED" | "FORBIDDEN", message: string) {
    super(message);
  }
}

export async function hashPassword(pw: string) {
  return bcrypt.hash(pw, 12);
}

// ---- simple in-memory brute-force protection (per server instance) ----
const failures = new Map<string, { count: number; until: number }>();
function throttled(key: string) {
  const f = failures.get(key);
  return !!f && f.count >= 5 && f.until > Date.now();
}
function noteFailure(key: string) {
  const f = failures.get(key);
  const count = f && f.until > Date.now() ? f.count + 1 : 1;
  failures.set(key, { count, until: Date.now() + 15 * 60_000 });
}

export async function login(email: string, password: string): Promise<{ ok: true } | { ok: false; error: string }> {
  const key = email.trim().toLowerCase();
  if (throttled(key)) return { ok: false, error: "Too many failed attempts. Wait 15 minutes and try again." };
  const user = await prisma.user.findUnique({ where: { email: key } });
  const ok = !!user && user.isActive && !!user.passwordHash && (await bcrypt.compare(password, user.passwordHash));
  if (!ok || !user) {
    noteFailure(key);
    return { ok: false, error: "Email or password is incorrect." };
  }
  failures.delete(key);
  const token = randomBytes(32).toString("base64url");
  const ua = (await headers()).get("user-agent")?.slice(0, 200) ?? null;
  await prisma.session.create({
    data: { id: sha256(token), userId: user.id, expiresAt: new Date(Date.now() + SESSION_HOURS * 3_600_000), userAgent: ua },
  });
  (await cookies()).set(COOKIE, token, {
    httpOnly: true,
    sameSite: "lax",
    secure: process.env.COOKIE_SECURE === "true",
    path: "/",
    maxAge: SESSION_HOURS * 3600,
  });
  await prisma.auditLog.create({
    data: { actorUserId: user.id, actorName: user.name, action: "auth.login", entityType: "User", entityId: user.id, summary: `${user.name} signed in` },
  });
  return { ok: true };
}

export async function logout() {
  const jar = await cookies();
  const token = jar.get(COOKIE)?.value;
  if (token) await prisma.session.deleteMany({ where: { id: sha256(token) } });
  jar.delete(COOKIE);
}

export async function getCurrentUser(): Promise<CurrentUser | null> {
  const token = (await cookies()).get(COOKIE)?.value;
  if (!token) return null;
  const session = await prisma.session.findUnique({
    where: { id: sha256(token) },
    include: { user: { include: { roles: true } } },
  });
  if (!session || session.expiresAt < new Date() || !session.user.isActive) return null;
  if (Date.now() - session.lastSeenAt.getTime() > 5 * 60_000) {
    await prisma.session.update({ where: { id: session.id }, data: { lastSeenAt: new Date() } }).catch(() => {});
  }
  const roles = session.user.roles.map((r) => r.role);
  return {
    id: session.user.id,
    name: session.user.name,
    email: session.user.email,
    roles,
    employeeId: session.user.employeeId,
    permissions: [...permissionsFor(roles)],
  };
}

/** For server actions / route handlers: throws instead of redirecting. */
export async function requireUser(perm?: Permission): Promise<CurrentUser> {
  const u = await getCurrentUser();
  if (!u) throw new AuthError("UNAUTHENTICATED", "Please sign in again.");
  if (perm && !can(u.roles, perm)) throw new AuthError("FORBIDDEN", "You do not have permission to do this.");
  return u;
}

/** For pages: redirects to login, or shows the access-denied page. */
export async function pageUser(perm?: Permission | Permission[]): Promise<CurrentUser> {
  const u = await getCurrentUser();
  if (!u) redirect("/login");
  const perms = perm ? (Array.isArray(perm) ? perm : [perm]) : [];
  if (perms.length && !perms.some((p) => can(u.roles, p))) redirect("/denied");
  return u;
}

export function hasPerm(u: CurrentUser | null, p: Permission) {
  return !!u && u.permissions.includes(p);
}
