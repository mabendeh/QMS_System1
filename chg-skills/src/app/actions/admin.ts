"use server";
import { revalidatePath } from "next/cache";
import { z } from "zod";
import ExcelJS from "exceljs";
import { Prisma, type Role } from "@prisma/client";
import { prisma } from "@/lib/db";
import { AuthError, hashPassword, requireUser } from "@/lib/auth";
import { audit } from "@/lib/audit";
import type { ActionResult } from "./skills";

function fail(e: unknown): ActionResult {
  if (e instanceof AuthError) return { ok: false, error: e.message };
  if (e instanceof z.ZodError) return { ok: false, error: e.errors.map((x) => x.message).join(" ") };
  if (e instanceof Prisma.PrismaClientKnownRequestError && e.code === "P2002") return { ok: false, error: "That name, code or email is already in use." };
  console.error(e);
  return { ok: false, error: "Something went wrong. Nothing was saved." };
}
const bool = z.string().optional().transform((v) => v === "on" || v === "true");
const codeFrom = (name: string) => name.replace(/[^A-Za-z0-9]+/g, "").slice(0, 6).toUpperCase() || "AREA";

// ------------------------------------------------------------------ departments / areas

const areaSchema = z.object({
  id: z.string().optional().transform((v) => v || null),
  name: z.string().trim().min(2, "Name is required.").max(60),
  code: z.string().trim().max(10).optional(),
  description: z.string().trim().max(500).optional(),
  kind: z.enum(["PRODUCTION", "PLANT_WIDE"]),
  isActive: bool,
});

export async function saveArea(form: FormData): Promise<ActionResult> {
  try {
    const u = await requireUser("config:manage");
    const d = areaSchema.parse(Object.fromEntries(form.entries()));
    await prisma.$transaction(async (tx) => {
      if (d.id) {
        const before = await tx.productionArea.findUniqueOrThrow({ where: { id: d.id } });
        const a = await tx.productionArea.update({ where: { id: d.id }, data: { name: d.name, description: d.description || null, kind: d.kind, isActive: d.isActive } });
        await audit(tx, u, "config.area_updated", "ProductionArea", a.id, `${u.name} updated department ${a.name}${before.name !== a.name ? ` (renamed from ${before.name})` : ""}${before.isActive !== a.isActive ? (a.isActive ? " – enabled" : " – disabled") : ""}`, before, a);
      } else {
        const max = await tx.productionArea.aggregate({ _max: { sortOrder: true } });
        let code = (d.code || codeFrom(d.name)).toUpperCase();
        if (await tx.productionArea.findUnique({ where: { code } })) code = `${code}${Math.floor(Math.random() * 90 + 10)}`;
        const a = await tx.productionArea.create({ data: { name: d.name, code, description: d.description || null, kind: d.kind, isActive: d.isActive, sortOrder: (max._max.sortOrder ?? 0) + 10 } });
        await audit(tx, u, "config.area_created", "ProductionArea", a.id, `${u.name} added department ${a.name}`, undefined, a);
      }
    });
    revalidatePath("/", "layout");
    return { ok: true, message: "Department saved." };
  } catch (e) {
    return fail(e);
  }
}

export async function moveItem(form: FormData): Promise<ActionResult> {
  try {
    const u = await requireUser("config:manage");
    const kind = String(form.get("kind"));
    const id = String(form.get("id"));
    const dir = form.get("dir") === "up" ? -1 : 1;
    await prisma.$transaction(async (tx) => {
      const list = kind === "area"
        ? await tx.productionArea.findMany({ orderBy: [{ sortOrder: "asc" }, { name: "asc" }], select: { id: true } })
        : await tx.competency.findMany({ where: { areaId: (await tx.competency.findUniqueOrThrow({ where: { id } })).areaId }, orderBy: [{ sortOrder: "asc" }, { name: "asc" }], select: { id: true } });
      const i = list.findIndex((x) => x.id === id);
      const j = i + dir;
      if (i < 0 || j < 0 || j >= list.length) return;
      [list[i], list[j]] = [list[j], list[i]];
      for (const [k, x] of list.entries()) {
        if (kind === "area") await tx.productionArea.update({ where: { id: x.id }, data: { sortOrder: (k + 1) * 10 } });
        else await tx.competency.update({ where: { id: x.id }, data: { sortOrder: (k + 1) * 10 } });
      }
      await audit(tx, u, "config.reordered", kind === "area" ? "ProductionArea" : "Competency", id, `${u.name} reordered ${kind === "area" ? "departments" : "competencies"}`);
    });
    revalidatePath("/", "layout");
    return { ok: true, message: "Order updated." };
  } catch (e) {
    return fail(e);
  }
}

const compSchema = z.object({
  id: z.string().optional().transform((v) => v || null),
  areaId: z.string().min(1),
  name: z.string().trim().min(2, "Name is required.").max(120),
  description: z.string().trim().max(1000).optional(),
  isCritical: bool,
  isActive: bool,
  requalMonths: z.string().optional().transform((v) => (v && v.trim() ? Number(v) : null)).refine((v) => v == null || (Number.isInteger(v) && v > 0 && v <= 120), "Requalification must be 1–120 months."),
});

export async function saveCompetency(form: FormData): Promise<ActionResult> {
  try {
    const u = await requireUser("config:manage");
    const d = compSchema.parse(Object.fromEntries(form.entries()));
    await prisma.$transaction(async (tx) => {
      const data = { name: d.name, description: d.description || null, isCritical: d.isCritical, requiresEvidence: d.isCritical, isActive: d.isActive };
      let id = d.id;
      if (id) {
        const before = await tx.competency.findUniqueOrThrow({ where: { id } });
        const c = await tx.competency.update({ where: { id }, data });
        await audit(tx, u, "config.competency_updated", "Competency", c.id, `${u.name} updated competency ${c.name}`, before, c);
      } else {
        const area = await tx.productionArea.findUniqueOrThrow({ where: { id: d.areaId } });
        const count = await tx.competency.count({ where: { areaId: area.id } });
        const max = await tx.competency.aggregate({ where: { areaId: area.id }, _max: { sortOrder: true } });
        let code = `${area.code}-${String(count + 1).padStart(2, "0")}`;
        while (await tx.competency.findUnique({ where: { code } })) code = `${code}x`;
        const c = await tx.competency.create({ data: { ...data, areaId: area.id, code, sortOrder: (max._max.sortOrder ?? 0) + 10 } });
        id = c.id;
        await audit(tx, u, "config.competency_created", "Competency", c.id, `${u.name} added competency ${c.name} to ${area.name}`, undefined, c);
      }
      const existing = await tx.qualificationExpirationRule.findFirst({ where: { competencyId: id } });
      if (d.requalMonths == null && existing) await tx.qualificationExpirationRule.delete({ where: { id: existing.id } });
      else if (d.requalMonths != null && existing) await tx.qualificationExpirationRule.update({ where: { id: existing.id }, data: { months: d.requalMonths } });
      else if (d.requalMonths != null) await tx.qualificationExpirationRule.create({ data: { competencyId: id, months: d.requalMonths } });
    });
    revalidatePath("/", "layout");
    return { ok: true, message: "Competency saved. Changes to requalification apply to qualifications verified from now on." };
  } catch (e) {
    return fail(e);
  }
}

export async function saveAreaExpiry(form: FormData): Promise<ActionResult> {
  try {
    const u = await requireUser("config:manage");
    const areaId = String(form.get("areaId"));
    const raw = String(form.get("months") ?? "").trim();
    const months = raw ? Number(raw) : null;
    if (months != null && (!Number.isInteger(months) || months < 1 || months > 120)) return { ok: false, error: "Requalification must be 1–120 months." };
    await prisma.$transaction(async (tx) => {
      const existing = await tx.qualificationExpirationRule.findFirst({ where: { areaId, competencyId: null } });
      if (months == null && existing) await tx.qualificationExpirationRule.delete({ where: { id: existing.id } });
      else if (months != null && existing) await tx.qualificationExpirationRule.update({ where: { id: existing.id }, data: { months } });
      else if (months != null) await tx.qualificationExpirationRule.create({ data: { areaId, months } });
      await audit(tx, u, "config.expiry_updated", "ProductionArea", areaId, `${u.name} set department requalification to ${months ?? "none"}`);
    });
    revalidatePath("/admin/areas");
    return { ok: true, message: "Requalification interval saved." };
  } catch (e) {
    return fail(e);
  }
}

// ------------------------------------------------------------------ job titles

export async function saveJobRole(form: FormData): Promise<ActionResult> {
  try {
    const u = await requireUser("config:manage");
    const id = String(form.get("id") ?? "") || null;
    const title = z.string().trim().min(2).max(80).parse(form.get("title"));
    const isActive = form.get("isActive") === "on";
    const areaIds = form.getAll("areaIds").map(String);
    await prisma.$transaction(async (tx) => {
      const role = id ? await tx.jobRole.update({ where: { id }, data: { title, isActive } }) : await tx.jobRole.create({ data: { title, isActive } });
      const comps = await tx.competency.findMany({ where: { areaId: { in: areaIds }, isActive: true }, select: { id: true } });
      await tx.competencyRequirement.deleteMany({ where: { jobRoleId: role.id } });
      await tx.competencyRequirement.createMany({ data: comps.map((c) => ({ jobRoleId: role.id, competencyId: c.id })) });
      await audit(tx, u, "config.jobrole_saved", "JobRole", role.id, `${u.name} saved job title ${title} (${areaIds.length} required departments)`, undefined, { title, areaIds });
    });
    revalidatePath("/", "layout");
    return { ok: true, message: "Job title saved." };
  } catch (e) {
    return fail(e);
  }
}

// ------------------------------------------------------------------ users

const ROLES: Role[] = ["HR_ADMIN", "PRODUCTION_MANAGER", "SUPERVISOR", "TRAINER", "CI_MANAGER", "QUALITY", "LEADERSHIP", "EMPLOYEE", "SYSTEM_ADMIN"];
const userSchema = z.object({
  id: z.string().optional().transform((v) => v || null),
  email: z.string().trim().toLowerCase().email("Enter a valid email."),
  name: z.string().trim().min(2).max(100),
  isActive: bool,
  employeeId: z.string().optional().transform((v) => v || null),
  password: z.string().optional().transform((v) => v || null).refine((v) => !v || v.length >= 12, "Passwords must be at least 12 characters."),
});

export async function saveUser(form: FormData): Promise<ActionResult> {
  try {
    const u = await requireUser("users:manage");
    const d = userSchema.parse(Object.fromEntries([...form.entries()].filter(([k]) => !["roles", "trainerAreas"].includes(k))));
    const roles = form.getAll("roles").map(String).filter((r): r is Role => ROLES.includes(r as Role));
    const trainerAreas = form.getAll("trainerAreas").map(String);
    if (roles.length === 0) return { ok: false, error: "Choose at least one role." };
    if (roles.includes("EMPLOYEE") && !d.employeeId) return { ok: false, error: "Employee accounts must be linked to an employee record." };
    if (!d.id && !d.password) return { ok: false, error: "Set an initial password (12+ characters), or connect SSO." };
    if (d.id === u.id && (!d.isActive || !roles.includes("SYSTEM_ADMIN"))) return { ok: false, error: "You cannot remove your own administrator access." };
    await prisma.$transaction(async (tx) => {
      const before = d.id ? await tx.user.findUnique({ where: { id: d.id }, include: { roles: true } }) : null;
      const data = { email: d.email, name: d.name, isActive: d.isActive, employeeId: d.employeeId, ...(d.password ? { passwordHash: await hashPassword(d.password) } : {}) };
      const user = d.id ? await tx.user.update({ where: { id: d.id }, data }) : await tx.user.create({ data });
      await tx.userRole.deleteMany({ where: { userId: user.id } });
      await tx.userRole.createMany({ data: roles.map((role) => ({ userId: user.id, role })) });
      await tx.trainerAuthorization.deleteMany({ where: { userId: user.id, areaId: { notIn: roles.includes("TRAINER") ? trainerAreas : [] } } });
      if (roles.includes("TRAINER")) for (const areaId of trainerAreas) await tx.trainerAuthorization.upsert({ where: { userId_areaId: { userId: user.id, areaId } }, create: { userId: user.id, areaId }, update: {} });
      if (!d.isActive || d.password) await tx.session.deleteMany({ where: { userId: user.id } });
      await audit(tx, u, d.id ? "users.updated" : "users.created", "User", user.id, `${u.name} ${d.id ? "updated" : "created"} user ${user.email}`,
        before ? { email: before.email, isActive: before.isActive, roles: before.roles.map((r) => r.role) } : undefined,
        { email: user.email, isActive: user.isActive, roles, trainerAreas, passwordChanged: !!d.password });
    });
    revalidatePath("/admin/users");
    return { ok: true, message: "User saved." };
  } catch (e) {
    return fail(e);
  }
}

// ------------------------------------------------------------------ employee import

export interface ImportRow {
  row: number;
  employeeNumber: string;
  firstName: string;
  lastName: string;
  shift: string;
  department: string;
  jobTitle: string;
  hireDate: string;
  supervisorEmail: string;
  errors: string[];
  action: "create" | "update" | "skip";
}

const HEADERS: Record<string, keyof ImportRow> = {
  "employee id": "employeeNumber", "employee number": "employeeNumber", id: "employeeNumber",
  "first name": "firstName", "last name": "lastName", shift: "shift", department: "department",
  "job title": "jobTitle", "hire date": "hireDate", "supervisor email": "supervisorEmail",
};

export async function previewImport(form: FormData): Promise<{ ok: true; rows: ImportRow[]; summary: { create: number; update: number; errors: number } } | { ok: false; error: string }> {
  try {
    await requireUser("employee:import");
    const file = form.get("file");
    if (!(file instanceof File) || file.size === 0) return { ok: false, error: "Choose a CSV or XLSX file." };
    if (file.size > 5 * 1024 * 1024) return { ok: false, error: "Import files must be 5 MB or smaller." };
    const buf = Buffer.from(await file.arrayBuffer());
    const wb = new ExcelJS.Workbook();
    const isXlsx = buf.subarray(0, 2).toString() === "PK";
    let grid: string[][] = [];
    if (isXlsx) {
      await wb.xlsx.load(buf as unknown as ArrayBuffer);
      const ws = wb.worksheets[0];
      ws.eachRow({ includeEmpty: false }, (r) => {
        const vals = (r.values as unknown[]).slice(1).map((v) => (v instanceof Date ? v.toISOString().slice(0, 10) : v && typeof v === "object" && "text" in (v as object) ? String((v as { text: string }).text) : v == null ? "" : String(v)));
        grid.push(vals);
      });
    } else {
      grid = parseCsv(buf.toString("utf8"));
    }
    if (grid.length < 2) return { ok: false, error: "The file has no data rows." };
    const header = grid[0].map((h) => HEADERS[h.trim().toLowerCase()]);
    if (!header.includes("employeeNumber") || !header.includes("firstName") || !header.includes("lastName") || !header.includes("shift")) {
      return { ok: false, error: "Required columns: Employee ID, First Name, Last Name, Shift. Optional: Department, Job Title, Hire Date, Supervisor Email." };
    }
    const [shifts, areas, jobs, sups, existing] = await Promise.all([
      prisma.shift.findMany(), prisma.productionArea.findMany(), prisma.jobRole.findMany(),
      prisma.user.findMany({ where: { roles: { some: { role: "SUPERVISOR" } } }, select: { email: true } }),
      prisma.employee.findMany({ select: { employeeNumber: true } }),
    ]);
    const existingIds = new Set(existing.map((e) => e.employeeNumber.toLowerCase()));
    const seen = new Map<string, number>();
    const rows: ImportRow[] = grid.slice(1).filter((r) => r.some((c) => String(c).trim())).slice(0, 2000).map((r, i) => {
      const rec: Record<string, string> = {};
      header.forEach((k, j) => { if (k) rec[k] = String(r[j] ?? "").trim(); });
      const row: ImportRow = {
        row: i + 2, employeeNumber: rec.employeeNumber ?? "", firstName: rec.firstName ?? "", lastName: rec.lastName ?? "", shift: (rec.shift ?? "").replace(/^shift\s*/i, "").toUpperCase(),
        department: rec.department ?? "", jobTitle: rec.jobTitle ?? "", hireDate: rec.hireDate ?? "", supervisorEmail: (rec.supervisorEmail ?? "").toLowerCase(), errors: [], action: "create",
      };
      if (!row.employeeNumber) row.errors.push("Employee ID is missing");
      else if (!/^[A-Za-z0-9_-]{1,30}$/.test(row.employeeNumber)) row.errors.push("Employee ID may only contain letters, numbers, - and _");
      if (!row.firstName || !row.lastName) row.errors.push("First and last name are required");
      if (!shifts.some((s) => s.code === row.shift)) row.errors.push(`Unknown shift “${row.shift}” (use A, B or C)`);
      if (row.department && !areas.some((a) => a.name.toLowerCase() === row.department.toLowerCase())) row.errors.push(`Unknown department “${row.department}”`);
      if (row.jobTitle && !jobs.some((j) => j.title.toLowerCase() === row.jobTitle.toLowerCase())) row.errors.push(`Unknown job title “${row.jobTitle}”`);
      if (row.hireDate && isNaN(new Date(row.hireDate).getTime())) row.errors.push("Hire date is not a valid date");
      if (row.supervisorEmail && !sups.some((s) => s.email === row.supervisorEmail)) row.errors.push(`No supervisor account for ${row.supervisorEmail}`);
      const key = row.employeeNumber.toLowerCase();
      if (key && seen.has(key)) row.errors.push(`Duplicate Employee ID (also on row ${seen.get(key)})`);
      if (key) seen.set(key, row.row);
      if (key && existingIds.has(key)) row.action = "update";
      if (row.errors.length) row.action = "skip";
      return row;
    });
    return { ok: true, rows, summary: { create: rows.filter((r) => r.action === "create").length, update: rows.filter((r) => r.action === "update").length, errors: rows.filter((r) => r.action === "skip").length } };
  } catch (e) {
    if (e instanceof AuthError) return { ok: false, error: e.message };
    console.error(e);
    return { ok: false, error: "The file could not be read. Save it as CSV or XLSX and try again." };
  }
}

function parseCsv(text: string): string[][] {
  const out: string[][] = [];
  let row: string[] = [], cell = "", q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) {
      if (ch === '"' && text[i + 1] === '"') { cell += '"'; i++; } else if (ch === '"') q = false; else cell += ch;
    } else if (ch === '"') q = true;
    else if (ch === ",") { row.push(cell); cell = ""; }
    else if (ch === "\n" || ch === "\r") { if (ch === "\r" && text[i + 1] === "\n") i++; row.push(cell); out.push(row); row = []; cell = ""; }
    else cell += ch;
  }
  if (cell || row.length) { row.push(cell); out.push(row); }
  return out.map((r) => r.map((c) => c.replace(/^﻿/, "")));
}

/** Creates new employees and (only if explicitly allowed) updates details of existing ones.
 *  Never touches qualification records. */
export async function commitImport(json: string, updateExisting: boolean): Promise<ActionResult> {
  try {
    const u = await requireUser("employee:import");
    const rows = (JSON.parse(json) as ImportRow[]).filter((r) => r.action !== "skip" && r.errors.length === 0);
    const [shifts, areas, jobs, sups] = await Promise.all([prisma.shift.findMany(), prisma.productionArea.findMany(), prisma.jobRole.findMany(), prisma.user.findMany({ select: { id: true, email: true } })]);
    let created = 0, updated = 0, skipped = 0;
    await prisma.$transaction(async (tx) => {
      for (const r of rows) {
        const data = {
          firstName: r.firstName, lastName: r.lastName, shiftId: shifts.find((s) => s.code === r.shift)!.id,
          primaryAreaId: areas.find((a) => a.name.toLowerCase() === r.department.toLowerCase())?.id ?? null,
          jobRoleId: jobs.find((j) => j.title.toLowerCase() === r.jobTitle.toLowerCase())?.id ?? null,
          supervisorUserId: sups.find((s) => s.email === r.supervisorEmail)?.id ?? null,
          hireDate: r.hireDate ? new Date(r.hireDate) : null,
        };
        const existing = await tx.employee.findUnique({ where: { employeeNumber: r.employeeNumber } });
        if (existing) {
          if (!updateExisting) { skipped++; continue; }
          await tx.employee.update({ where: { id: existing.id }, data });
          updated++;
        } else {
          await tx.employee.create({ data: { ...data, employeeNumber: r.employeeNumber } });
          created++;
        }
      }
      await audit(tx, u, "employees.imported", "Employee", null, `${u.name} imported employees: ${created} created, ${updated} updated, ${skipped} existing skipped`, undefined, { created, updated, skipped });
    }, { timeout: 60_000 });
    revalidatePath("/", "layout");
    return { ok: true, message: `Import complete: ${created} created, ${updated} updated${skipped ? `, ${skipped} existing employees left unchanged` : ""}. No qualifications were changed.` };
  } catch (e) {
    return fail(e);
  }
}

// ------------------------------------------------------------------ sample data

export async function deleteSampleData(form: FormData): Promise<ActionResult> {
  try {
    const u = await requireUser("sample:delete");
    if (String(form.get("confirm") ?? "").trim() !== "DELETE SAMPLE DATA") return { ok: false, error: "Type DELETE SAMPLE DATA exactly to confirm." };
    const result = await prisma.$transaction(async (tx) => {
      const ids = (await tx.employee.findMany({ where: { isSample: true }, select: { id: true } })).map((e) => e.id);
      if (ids.length === 0) return 0;
      const reqIds = (await tx.promotionRequest.findMany({ where: { employeeId: { in: ids } }, select: { id: true } })).map((r) => r.id);
      const assessIds = (await tx.skillAssessment.findMany({ where: { employeeId: { in: ids } }, select: { id: true } })).map((r) => r.id);
      // Approval history is append-only; this guarded purge is the only path allowed to remove sample rows.
      await tx.$executeRawUnsafe(`SET LOCAL chg.allow_history_purge = 'on'`);
      await tx.approvalRecord.deleteMany({ where: { OR: [{ promotionRequestId: { in: reqIds } }, { assessmentId: { in: assessIds } }] } });
      await tx.$executeRawUnsafe(`SET LOCAL chg.allow_history_purge = 'off'`);
      await tx.promotionRequest.deleteMany({ where: { id: { in: reqIds } } });
      await tx.user.updateMany({ where: { employeeId: { in: ids } }, data: { isActive: false, employeeId: null } });
      await tx.employee.deleteMany({ where: { id: { in: ids } } });
      await audit(tx, u, "system.sample_deleted", "System", null, `${u.name} deleted ${ids.length} sample employees and their sample records`);
      return ids.length;
    }, { timeout: 60_000 });
    revalidatePath("/", "layout");
    return { ok: true, message: result ? `Deleted ${result} sample employees and all their sample qualifications, training and promotion records. Demo login accounts linked to them were disabled.` : "There is no sample data left." };
  } catch (e) {
    return fail(e);
  }
}
