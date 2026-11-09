/**
 * Qualification status rules. Pure functions – no database access – so they
 * can be unit-tested and reused by the matrix, exports and dashboards.
 */
export type StoredStatus = "NOT_APPLICABLE" | "NEEDS_TRAINING" | "IN_TRAINING" | "TRAINED" | "COMPETENT";
export type DisplayStatus = "NOT_ASSESSED" | "NEEDS_TRAINING" | "IN_TRAINING" | "TRAINED" | "COMPETENT" | "EXPIRED";

export const STATUS_META: Record<
  DisplayStatus,
  { label: string; short: string; color: string; text: string; argb: string; fontArgb: string; description: string }
> = {
  NEEDS_TRAINING: {
    label: "Needs Training", short: "NT", color: "#C62828", text: "#FFFFFF", argb: "FFC62828", fontArgb: "FFFFFFFF",
    description: "Has not demonstrated the required competency.",
  },
  IN_TRAINING: {
    label: "In Training", short: "IT", color: "#F2C230", text: "#1A2433", argb: "FFF2C230", fontArgb: "FF1A2433",
    description: "Learning; must work under supervision.",
  },
  TRAINED: {
    label: "Trained", short: "TR", color: "#2E8B57", text: "#FFFFFF", argb: "FF2E8B57", fontArgb: "FFFFFFFF",
    description: "Training complete; awaiting independent competency verification.",
  },
  COMPETENT: {
    label: "Competent", short: "C", color: "#1E63C4", text: "#FFFFFF", argb: "FF1E63C4", fontArgb: "FFFFFFFF",
    description: "Verified to operate independently; passed qualification.",
  },
  EXPIRED: {
    label: "Expired", short: "EXP", color: "#8E1B1B", text: "#FFFFFF", argb: "FF8E1B1B", fontArgb: "FFFFFFFF",
    description: "Qualification lapsed; requalification needed. Not counted as qualified.",
  },
  NOT_ASSESSED: {
    label: "Not Applicable / Not Assessed", short: "—", color: "#E3E7EC", text: "#5D6B7E", argb: "FFE3E7EC", fontArgb: "FF5D6B7E",
    description: "Not required for this job, or not assessed yet.",
  },
};

export const STATUS_ORDER: DisplayStatus[] = ["NEEDS_TRAINING", "IN_TRAINING", "TRAINED", "COMPETENT", "EXPIRED", "NOT_ASSESSED"];

export const SELECTABLE_STATUSES: StoredStatus[] = ["NEEDS_TRAINING", "IN_TRAINING", "TRAINED", "COMPETENT", "NOT_APPLICABLE"];

export const STORED_LABEL: Record<StoredStatus, string> = {
  NOT_APPLICABLE: "Not Applicable",
  NEEDS_TRAINING: "Needs Training",
  IN_TRAINING: "In Training",
  TRAINED: "Trained",
  COMPETENT: "Competent",
};

export interface SkillRecordLike {
  status: StoredStatus;
  expiresAt?: Date | string | null;
}

export function isExpired(rec: SkillRecordLike | null | undefined, now: Date = new Date()): boolean {
  if (!rec || rec.status !== "COMPETENT" || !rec.expiresAt) return false;
  return new Date(rec.expiresAt).getTime() < now.getTime();
}

/** Colour shown in the matrix. Unassessed skills required by the job show red. */
export function displayStatus(rec: SkillRecordLike | null | undefined, requiredForJob: boolean, now: Date = new Date()): DisplayStatus {
  if (!rec) return requiredForJob ? "NEEDS_TRAINING" : "NOT_ASSESSED";
  if (rec.status === "NOT_APPLICABLE") return "NOT_ASSESSED";
  if (isExpired(rec, now)) return "EXPIRED";
  return rec.status;
}

/** Counts toward qualified coverage only when Competent and not expired. */
export function isCurrentlyQualified(rec: SkillRecordLike | null | undefined, now: Date = new Date()): boolean {
  return !!rec && rec.status === "COMPETENT" && !isExpired(rec, now);
}

export function isExpiringSoon(rec: SkillRecordLike | null | undefined, days = 30, now: Date = new Date()): boolean {
  if (!rec || rec.status !== "COMPETENT" || !rec.expiresAt) return false;
  const t = new Date(rec.expiresAt).getTime();
  return t >= now.getTime() && t <= now.getTime() + days * 86_400_000;
}

export function computeExpiry(qualifiedAt: Date, months: number | null | undefined): Date | null {
  if (!months || months <= 0) return null;
  const d = new Date(qualifiedAt);
  d.setMonth(d.getMonth() + months);
  return d;
}

/** Roll a set of competency statuses up to one area-level colour. */
export function areaRollup(statuses: DisplayStatus[], isPrimaryArea: boolean): { status: DisplayStatus; competent: number; total: number } {
  const relevant = statuses.filter((s) => s !== "NOT_ASSESSED");
  const competent = statuses.filter((s) => s === "COMPETENT").length;
  const total = statuses.length;
  if (total > 0 && competent === total) return { status: "COMPETENT", competent, total };
  if (relevant.length === 0) return { status: isPrimaryArea ? "NEEDS_TRAINING" : "NOT_ASSESSED", competent, total };
  if (statuses.every((s) => s === "TRAINED" || s === "COMPETENT")) return { status: "TRAINED", competent, total };
  if (relevant.some((s) => s === "IN_TRAINING" || s === "TRAINED" || s === "COMPETENT")) return { status: "IN_TRAINING", competent, total };
  if (relevant.some((s) => s === "EXPIRED")) return { status: "EXPIRED", competent, total };
  return { status: "NEEDS_TRAINING", competent, total };
}

// ------------------------------------------------------------------ assessment decision

export interface AssessmentDecisionInput {
  requested: StoredStatus;
  current: StoredStatus | null;
  trainingDate?: Date | null;
  notes?: string | null;
  competency: { isCritical: boolean; requiresEvidence: boolean };
  /** Latest practical evaluation for this employee + competency, if any. */
  practical: { result: "PASS" | "FAIL"; independentOperation: boolean; evaluatorUserId: string } | null;
  evidenceCount: number;
  actor: { id: string; canVerify: boolean };
}

export type AssessmentDecision =
  | { outcome: "APPLY"; verified: boolean }
  | { outcome: "PENDING"; reason: string }
  | { outcome: "BLOCKED"; reasons: string[] };

/**
 * Decides what happens when someone submits a rating.
 * - Competent requires a passed practical evaluation showing independent operation,
 *   evidence for critical skills, and verification by an authorized evaluator who
 *   did not perform the practical evaluation themselves.
 * - Anything else applies immediately (with the rules below).
 */
export function decideAssessment(i: AssessmentDecisionInput): AssessmentDecision {
  const reasons: string[] = [];
  if (i.requested === "TRAINED" && !i.trainingDate) {
    reasons.push("Enter the training completion date before marking Trained.");
  }
  if (i.current === "COMPETENT" && i.requested !== "COMPETENT" && !(i.notes && i.notes.trim().length >= 5)) {
    reasons.push("Add a note explaining why a Competent qualification is being lowered.");
  }
  if (i.requested === "COMPETENT") {
    if (!i.practical) {
      reasons.push("A practical evaluation is required. A trainer must record a passed practical demonstration first.");
    } else if (i.practical.result !== "PASS") {
      reasons.push("The latest practical evaluation was not passed. A new passed evaluation is required.");
    } else if (!i.practical.independentOperation) {
      reasons.push("The practical evaluation must confirm independent operation.");
    }
    if ((i.competency.isCritical || i.competency.requiresEvidence) && i.evidenceCount === 0) {
      reasons.push("This is a critical skill: upload supporting evidence (signed checklist, photo or record).");
    }
  }
  if (reasons.length) return { outcome: "BLOCKED", reasons };
  if (i.requested === "COMPETENT") {
    if (!i.actor.canVerify) {
      return { outcome: "PENDING", reason: "Sent for verification by an authorized evaluator (Production Manager or Quality)." };
    }
    if (i.practical && i.practical.evaluatorUserId === i.actor.id) {
      return { outcome: "PENDING", reason: "You performed the practical evaluation, so a different authorized evaluator must verify." };
    }
    return { outcome: "APPLY", verified: true };
  }
  return { outcome: "APPLY", verified: false };
}

/** Whether a verifier may approve a pending Competent request. */
export function verificationBlockers(p: {
  verifierId: string;
  verifierCanVerify: boolean;
  submittedById: string;
  practicalEvaluatorId: string | null;
}): string[] {
  const out: string[] = [];
  if (!p.verifierCanVerify) out.push("You are not authorized to verify qualifications.");
  if (p.practicalEvaluatorId && p.practicalEvaluatorId === p.verifierId) {
    out.push("You performed the practical evaluation; a different evaluator must verify.");
  }
  return out;
}
