import "server-only";
import { prisma } from "./db";
import type { CurrentUser } from "./auth";

/**
 * Who may assess a given employee in a given area:
 *  - Production Manager: anyone
 *  - Supervisor: employees assigned to them
 *  - Trainer: areas they are authorized to train
 */
export async function canAssessEmployee(u: CurrentUser, emp: { supervisorUserId: string | null }, areaId: string) {
  if (!u.permissions.includes("skill:assess")) return false;
  if (u.roles.includes("PRODUCTION_MANAGER")) return true;
  if (u.roles.includes("SUPERVISOR") && emp.supervisorUserId === u.id) return true;
  if (u.roles.includes("TRAINER") && (await isAuthorizedTrainer(u.id, areaId))) return true;
  return false;
}

export async function isAuthorizedTrainer(userId: string, areaId: string) {
  const a = await prisma.trainerAuthorization.findUnique({ where: { userId_areaId: { userId, areaId } } });
  return !!a && (!a.validUntil || a.validUntil > new Date());
}

export async function canEvaluatePractical(u: CurrentUser, areaId: string) {
  if (!u.permissions.includes("training:evaluate")) return false;
  if (u.roles.includes("PRODUCTION_MANAGER")) return true;
  return u.roles.includes("TRAINER") && isAuthorizedTrainer(u.id, areaId);
}

/** Employees can only see their own record; staff with employee:view_all see everyone. */
export function canViewEmployee(u: CurrentUser, employeeId: string) {
  return u.permissions.includes("employee:view_all") || (u.permissions.includes("self:view") && u.employeeId === employeeId);
}
