import "server-only";
import type { SkillStatus } from "@prisma/client";
import { prisma } from "@/lib/db";
import {
  areaRollup, displayStatus, isCurrentlyQualified, isExpiringSoon, type DisplayStatus,
} from "@/lib/domain/status";
import { evaluateRequirements, qualifiedAreas, type EvaluationResult, type RequirementDef } from "@/lib/domain/promotion";

export interface SkillRec {
  id: string;
  status: SkillStatus;
  pendingStatus: SkillStatus | null;
  expiresAt: Date | null;
  qualifiedAt: Date | null;
  version: number;
}

/**
 * Loads the current state of the plant (active areas, competencies, employees and
 * qualifications) once, and exposes helpers used by the matrix, dashboards,
 * coverage and promotion screens. Everything is computed from stored records –
 * nothing is estimated.
 */
export async function loadPlant(opts: { includeInactiveEmployees?: boolean; employeeIds?: string[] } = {}) {
  const now = new Date();
  const [areas, employees, reqs, levels, ruleSets] = await Promise.all([
    prisma.productionArea.findMany({
      where: { isActive: true },
      orderBy: [{ kind: "asc" }, { sortOrder: "asc" }],
      include: { competencies: { where: { isActive: true }, orderBy: { sortOrder: "asc" } } },
    }),
    prisma.employee.findMany({
      where: {
        ...(opts.includeInactiveEmployees ? {} : { status: "ACTIVE" }),
        ...(opts.employeeIds ? { id: { in: opts.employeeIds } } : {}),
      },
      orderBy: [{ lastName: "asc" }, { firstName: "asc" }],
      include: { shift: true, jobRole: true, primaryArea: true, supervisor: { select: { id: true, name: true } }, careerLevel: true, department: true },
    }),
    prisma.competencyRequirement.findMany({ select: { jobRoleId: true, competencyId: true } }),
    prisma.careerLevel.findMany({ where: { isActive: true }, orderBy: { rank: "asc" } }),
    prisma.promotionRuleSet.findMany({ where: { isActive: true }, include: { requirements: { orderBy: { sortOrder: "asc" } } }, orderBy: { version: "desc" } }),
  ]);
  const skills = await prisma.employeeSkill.findMany({
    where: { employeeId: { in: employees.map((e) => e.id) } },
    select: { id: true, employeeId: true, competencyId: true, status: true, pendingStatus: true, expiresAt: true, qualifiedAt: true, version: true },
  });

  const skillMap = new Map<string, Map<string, SkillRec>>();
  for (const s of skills) {
    if (!skillMap.has(s.employeeId)) skillMap.set(s.employeeId, new Map());
    skillMap.get(s.employeeId)!.set(s.competencyId, s);
  }
  const requiredByJob = new Map<string, Set<string>>();
  for (const r of reqs) {
    if (!requiredByJob.has(r.jobRoleId)) requiredByJob.set(r.jobRoleId, new Set());
    requiredByJob.get(r.jobRoleId)!.add(r.competencyId);
  }
  const competenciesByArea = new Map(areas.map((a) => [a.id, a.competencies.map((c) => c.id)]));
  const productionAreaIds = new Set(areas.filter((a) => a.kind === "PRODUCTION").map((a) => a.id));
  const areaName = new Map(areas.map((a) => [a.id, a.name]));
  const compName = new Map(areas.flatMap((a) => a.competencies.map((c) => [c.id, `${c.name} (${a.name})`] as const)));
  const allComps = areas.flatMap((a) => a.competencies.map((c) => ({ ...c, areaName: a.name, areaKind: a.kind })));

  const empById = new Map(employees.map((e) => [e.id, e]));
  const rec = (empId: string, compId: string) => skillMap.get(empId)?.get(compId) ?? null;
  const required = (empId: string, compId: string) => {
    const e = empById.get(empId);
    return !!e?.jobRoleId && !!requiredByJob.get(e.jobRoleId)?.has(compId);
  };
  const display = (empId: string, compId: string): DisplayStatus => displayStatus(rec(empId, compId), required(empId, compId), now);
  const competentIds = (empId: string) => {
    const out = new Set<string>();
    for (const [cid, r] of skillMap.get(empId) ?? []) if (isCurrentlyQualified(r, now)) out.add(cid);
    return out;
  };
  const qualifiedAreaIds = (empId: string) => qualifiedAreas(competenciesByArea, competentIds(empId));
  const areaCell = (empId: string, areaId: string) => {
    const comps = competenciesByArea.get(areaId) ?? [];
    const e = empById.get(empId);
    const r = areaRollup(comps.map((c) => display(empId, c)), e?.primaryAreaId === areaId);
    const pending = comps.some((c) => !!rec(empId, c)?.pendingStatus);
    return { ...r, pending };
  };

  const activeRuleFor = (levelId: string) => ruleSets.find((r) => r.careerLevelId === levelId) ?? null;
  const toDefs = (rs: (typeof ruleSets)[number]): RequirementDef[] =>
    rs.requirements.map((r) => ({ id: r.id, type: r.type, label: r.label, areaId: r.areaId, competencyId: r.competencyId, minCount: r.minCount }));

  function evaluateFor(empId: string, targetLevelId: string): { ruleSetId: string; version: number; result: EvaluationResult } | null {
    const e = empById.get(empId);
    const rs = activeRuleFor(targetLevelId);
    if (!e || !rs) return null;
    const result = evaluateRequirements(
      toDefs(rs),
      {
        primaryAreaId: e.primaryAreaId,
        currentLevelRank: e.careerLevel?.rank ?? null,
        qualifiedAreaIds: qualifiedAreaIds(empId),
        productionAreaIds,
        competentCompetencyIds: competentIds(empId),
      },
      { areaName: (id) => areaName.get(id) ?? "Unknown area", competencyName: (id) => compName.get(id) ?? "Unknown competency" },
    );
    return { ruleSetId: rs.id, version: rs.version, result };
  }

  function nextLevel(empId: string) {
    const e = empById.get(empId);
    const rank = e?.careerLevel?.rank ?? -1;
    const next = levels.find((l) => l.rank > rank) ?? null;
    if (!next) return null;
    const ev = evaluateFor(empId, next.id);
    return ev ? { level: next, ...ev } : { level: next, ruleSetId: null, version: null, result: null };
  }

  function employeeSummary(empId: string) {
    const e = empById.get(empId)!;
    const req = e.jobRoleId ? [...(requiredByJob.get(e.jobRoleId) ?? [])].filter((c) => compName.has(c)) : [];
    const statuses = req.map((c) => display(empId, c));
    const qa = qualifiedAreaIds(empId);
    const prodQualified = [...qa].filter((a) => productionAreaIds.has(a));
    return {
      requiredCount: req.length,
      competentRequired: statuses.filter((s) => s === "COMPETENT").length,
      needsTraining: statuses.filter((s) => s === "NEEDS_TRAINING" || s === "EXPIRED").length,
      inTraining: statuses.filter((s) => s === "IN_TRAINING" || s === "TRAINED").length,
      fullyQualified: req.length > 0 && statuses.every((s) => s === "COMPETENT"),
      qualifiedAreaIds: qa,
      productionAreasQualified: prodQualified.length,
      crossTrained: prodQualified.length >= 2,
      pending: [...(skillMap.get(empId)?.values() ?? [])].some((r) => !!r.pendingStatus),
      expiringSoon: [...(skillMap.get(empId)?.values() ?? [])].filter((r) => isExpiringSoon(r, 30, now)).length,
    };
  }

  return {
    now, areas, employees, levels, ruleSets, skillMap, allComps, productionAreaIds, competenciesByArea,
    areaName, compName, rec, required, display, competentIds, qualifiedAreaIds, areaCell, evaluateFor, nextLevel, employeeSummary,
  };
}

export type Plant = Awaited<ReturnType<typeof loadPlant>>;
