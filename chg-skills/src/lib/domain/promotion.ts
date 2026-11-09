/**
 * Career progression evaluation. Pure function: given configured requirements
 * and an employee's current qualifications, explain what is met and what is missing.
 * This never authorizes pay or employment changes – it only informs a human review.
 * Time employed is deliberately NOT an input.
 */
export type RequirementType =
  | "PRIMARY_AREA_QUALIFIED"
  | "AREA_QUALIFIED"
  | "MIN_PRODUCTION_AREAS_QUALIFIED"
  | "COMPETENCY_COMPETENT"
  | "CURRENT_LEVEL_AT_LEAST"
  | "MANUAL_VERIFICATION";

export interface RequirementDef {
  id: string;
  type: RequirementType;
  label: string;
  areaId?: string | null;
  competencyId?: string | null;
  minCount?: number | null;
}

export interface EmployeeQualificationSnapshot {
  primaryAreaId: string | null;
  currentLevelRank: number | null;
  qualifiedAreaIds: Set<string>;
  /** Only PRODUCTION-kind areas count toward cross-training totals. */
  productionAreaIds: Set<string>;
  competentCompetencyIds: Set<string>;
}

export interface Lookups {
  areaName: (id: string) => string;
  competencyName: (id: string) => string;
}

export interface RequirementResult {
  id: string;
  label: string;
  type: RequirementType;
  met: boolean;
  manual: boolean;
  detail: string;
}

export interface EvaluationResult {
  items: RequirementResult[];
  metCount: number;
  total: number;
  /** All system-checkable requirements met; manual ones are confirmed in the review workflow. */
  readyForReview: boolean;
  missing: string[];
}

export function evaluateRequirements(reqs: RequirementDef[], emp: EmployeeQualificationSnapshot, lk: Lookups): EvaluationResult {
  const items: RequirementResult[] = reqs.map((r) => {
    switch (r.type) {
      case "PRIMARY_AREA_QUALIFIED": {
        const met = !!emp.primaryAreaId && emp.qualifiedAreaIds.has(emp.primaryAreaId);
        const name = emp.primaryAreaId ? lk.areaName(emp.primaryAreaId) : "assigned area";
        return { id: r.id, label: r.label, type: r.type, manual: false, met, detail: met ? `Qualified in ${name}` : `Complete all ${name} competencies` };
      }
      case "AREA_QUALIFIED": {
        const met = !!r.areaId && emp.qualifiedAreaIds.has(r.areaId);
        const name = r.areaId ? lk.areaName(r.areaId) : "area";
        return { id: r.id, label: r.label, type: r.type, manual: false, met, detail: met ? `Qualified in ${name}` : `Complete ${name} qualification` };
      }
      case "MIN_PRODUCTION_AREAS_QUALIFIED": {
        const need = r.minCount ?? 1;
        const have = [...emp.qualifiedAreaIds].filter((a) => emp.productionAreaIds.has(a)).length;
        return { id: r.id, label: r.label, type: r.type, manual: false, met: have >= need, detail: `${have} of ${need} production areas qualified` };
      }
      case "COMPETENCY_COMPETENT": {
        const met = !!r.competencyId && emp.competentCompetencyIds.has(r.competencyId);
        const name = r.competencyId ? lk.competencyName(r.competencyId) : "competency";
        return { id: r.id, label: r.label, type: r.type, manual: false, met, detail: met ? `${name}: Competent` : `${name}: not yet Competent` };
      }
      case "CURRENT_LEVEL_AT_LEAST": {
        const need = r.minCount ?? 0;
        const have = emp.currentLevelRank ?? -1;
        return { id: r.id, label: r.label, type: r.type, manual: false, met: have >= need, detail: have >= need ? `Currently Level ${have}` : `Must currently be Level ${need} or higher` };
      }
      case "MANUAL_VERIFICATION":
      default:
        return { id: r.id, label: r.label, type: r.type, manual: true, met: false, detail: "Confirmed during the review workflow" };
    }
  });
  const auto = items.filter((i) => !i.manual);
  return {
    items,
    metCount: items.filter((i) => i.met).length,
    total: items.length,
    readyForReview: auto.every((i) => i.met),
    missing: items.filter((i) => !i.met && !i.manual).map((i) => i.detail),
  };
}

/** Area is qualified when every active competency in it is currently Competent. */
export function qualifiedAreas(competenciesByArea: Map<string, string[]>, competentIds: Set<string>): Set<string> {
  const out = new Set<string>();
  for (const [areaId, comps] of competenciesByArea) {
    if (comps.length > 0 && comps.every((c) => competentIds.has(c))) out.add(areaId);
  }
  return out;
}

/** Promotion workflow state machine. Returns the next status, or an error. */
export type PromotionStatus = "SUBMITTED" | "MANAGER_APPROVED" | "HR_VERIFIED" | "APPROVED" | "REJECTED" | "WITHDRAWN";
export type ApprovalStep = "MANAGER_REVIEW" | "HR_VERIFICATION" | "FINAL_APPROVAL";

export const STEP_FOR_STATUS: Partial<Record<PromotionStatus, { step: ApprovalStep; permission: string; label: string }>> = {
  SUBMITTED: { step: "MANAGER_REVIEW", permission: "promotion:manager_review", label: "Production Manager review" },
  MANAGER_APPROVED: { step: "HR_VERIFICATION", permission: "promotion:hr_verify", label: "HR qualification & pay-policy check" },
  HR_VERIFIED: { step: "FINAL_APPROVAL", permission: "promotion:final_approve", label: "Final approval" },
};

export function nextPromotionStatus(current: PromotionStatus, decision: "APPROVED" | "REJECTED"): PromotionStatus | null {
  if (!STEP_FOR_STATUS[current]) return null;
  if (decision === "REJECTED") return "REJECTED";
  return current === "SUBMITTED" ? "MANAGER_APPROVED" : current === "MANAGER_APPROVED" ? "HR_VERIFIED" : "APPROVED";
}
