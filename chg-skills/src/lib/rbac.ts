import type { Role } from "@prisma/client";

/**
 * Least-privilege permission map. Every server action and route handler checks
 * one of these before reading or changing data. The UI only hides controls;
 * it is never the security boundary.
 */
export type Permission =
  | "matrix:view"
  | "employee:view_all"
  | "employee:manage"
  | "employee:import"
  | "skill:assess"
  | "qualification:verify"
  | "training:view"
  | "training:manage"
  | "training:evaluate"
  | "promotion:view"
  | "promotion:submit"
  | "promotion:manager_review"
  | "promotion:hr_verify"
  | "promotion:final_approve"
  | "rules:manage"
  | "comp:view"
  | "comp:manage"
  | "config:manage"
  | "coverage:view"
  | "coverage:manage"
  | "analytics:view"
  | "audit:view"
  | "users:manage"
  | "export:data"
  | "sample:delete"
  | "self:view";

export const ROLE_LABELS: Record<Role, string> = {
  HR_ADMIN: "HR Administrator",
  PRODUCTION_MANAGER: "Production Manager",
  SUPERVISOR: "Supervisor",
  TRAINER: "Trainer / Training Coordinator",
  CI_MANAGER: "Continuous Improvement Manager",
  QUALITY: "Quality & Food Safety",
  LEADERSHIP: "Plant Leadership",
  EMPLOYEE: "Employee",
  SYSTEM_ADMIN: "System Administrator",
};

const STAFF_VIEW: Permission[] = ["matrix:view", "employee:view_all", "training:view", "coverage:view", "export:data"];

export const ROLE_PERMISSIONS: Record<Role, Permission[]> = {
  HR_ADMIN: [
    ...STAFF_VIEW, "employee:manage", "employee:import", "training:manage", "promotion:view",
    "promotion:hr_verify", "rules:manage", "comp:view", "comp:manage", "config:manage",
    "analytics:view", "audit:view",
  ],
  PRODUCTION_MANAGER: [
    ...STAFF_VIEW, "skill:assess", "qualification:verify", "training:manage", "training:evaluate",
    "promotion:view", "promotion:submit", "promotion:manager_review", "coverage:manage", "analytics:view",
  ],
  SUPERVISOR: [...STAFF_VIEW, "skill:assess", "training:manage", "promotion:view", "promotion:submit"],
  TRAINER: [...STAFF_VIEW, "skill:assess", "training:manage", "training:evaluate"],
  CI_MANAGER: [...STAFF_VIEW, "analytics:view"],
  QUALITY: [...STAFF_VIEW, "qualification:verify", "analytics:view", "audit:view"],
  LEADERSHIP: [...STAFF_VIEW, "promotion:view", "promotion:final_approve", "comp:view", "analytics:view"],
  EMPLOYEE: ["self:view"],
  // System admins run the system but do NOT see pay data or approve people decisions.
  SYSTEM_ADMIN: [...STAFF_VIEW, "config:manage", "users:manage", "audit:view", "sample:delete"],
};

export function permissionsFor(roles: Role[]): Set<Permission> {
  const out = new Set<Permission>();
  for (const r of roles) for (const p of ROLE_PERMISSIONS[r] ?? []) out.add(p);
  return out;
}

export function can(roles: Role[], perm: Permission): boolean {
  return permissionsFor(roles).has(perm);
}
