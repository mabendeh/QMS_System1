import { describe, expect, it } from "vitest";
import { areaRollup, computeExpiry, decideAssessment, displayStatus, isCurrentlyQualified, isExpiringSoon, verificationBlockers } from "@/lib/domain/status";

const now = new Date("2026-06-01T12:00:00Z");
const base = {
  current: "TRAINED" as const,
  competency: { isCritical: false, requiresEvidence: false },
  practical: { result: "PASS" as const, independentOperation: true, evaluatorUserId: "trainer" },
  evidenceCount: 0,
  actor: { id: "mgr", canVerify: true },
};

describe("display status", () => {
  it("shows unassessed required skills as Needs Training (red) and others as grey", () => {
    expect(displayStatus(null, true, now)).toBe("NEEDS_TRAINING");
    expect(displayStatus(null, false, now)).toBe("NOT_ASSESSED");
    expect(displayStatus({ status: "NOT_APPLICABLE" }, true, now)).toBe("NOT_ASSESSED");
  });
  it("maps stored statuses to their colours", () => {
    expect(displayStatus({ status: "IN_TRAINING" }, true, now)).toBe("IN_TRAINING");
    expect(displayStatus({ status: "COMPETENT" }, true, now)).toBe("COMPETENT");
  });
  it("Scenario 5: an expired qualification shows Expired and no longer counts as qualified", () => {
    const rec = { status: "COMPETENT" as const, expiresAt: new Date("2026-05-31T00:00:00Z") };
    expect(displayStatus(rec, true, now)).toBe("EXPIRED");
    expect(isCurrentlyQualified(rec, now)).toBe(false);
    expect(isCurrentlyQualified({ status: "COMPETENT", expiresAt: new Date("2026-07-01") }, now)).toBe(true);
    expect(isCurrentlyQualified({ status: "COMPETENT", expiresAt: null }, now)).toBe(true);
  });
  it("flags qualifications expiring within 30 days", () => {
    expect(isExpiringSoon({ status: "COMPETENT", expiresAt: new Date("2026-06-20") }, 30, now)).toBe(true);
    expect(isExpiringSoon({ status: "COMPETENT", expiresAt: new Date("2026-08-20") }, 30, now)).toBe(false);
  });
  it("computes requalification dates", () => {
    expect(computeExpiry(new Date("2026-01-15T00:00:00Z"), 12)?.toISOString().slice(0, 10)).toBe("2027-01-15");
    expect(computeExpiry(new Date(), null)).toBeNull();
  });
});

describe("area roll-up", () => {
  it("is Competent only when every competency is Competent", () => {
    expect(areaRollup(["COMPETENT", "COMPETENT"], false).status).toBe("COMPETENT");
    expect(areaRollup(["COMPETENT", "TRAINED"], false).status).toBe("TRAINED");
    expect(areaRollup(["COMPETENT", "NEEDS_TRAINING"], false).status).toBe("IN_TRAINING");
    expect(areaRollup(["NEEDS_TRAINING", "NEEDS_TRAINING"], false).status).toBe("NEEDS_TRAINING");
    expect(areaRollup(["NOT_ASSESSED", "NOT_ASSESSED"], false).status).toBe("NOT_ASSESSED");
    expect(areaRollup(["NOT_ASSESSED"], true).status).toBe("NEEDS_TRAINING");
  });
});

describe("assessment decisions", () => {
  it("Scenario 1: Needs Training → In Training applies immediately", () => {
    expect(decideAssessment({ ...base, current: "NEEDS_TRAINING", requested: "IN_TRAINING", practical: null })).toEqual({ outcome: "APPLY", verified: false });
  });
  it("Scenario 8: Competent without a practical evaluation is blocked with an explanation", () => {
    const d = decideAssessment({ ...base, requested: "COMPETENT", practical: null });
    expect(d.outcome).toBe("BLOCKED");
    if (d.outcome === "BLOCKED") expect(d.reasons.join(" ")).toMatch(/practical evaluation is required/i);
  });
  it("blocks Competent when the practical was failed or not independent", () => {
    expect(decideAssessment({ ...base, requested: "COMPETENT", practical: { ...base.practical, result: "FAIL" } }).outcome).toBe("BLOCKED");
    expect(decideAssessment({ ...base, requested: "COMPETENT", practical: { ...base.practical, independentOperation: false } }).outcome).toBe("BLOCKED");
  });
  it("requires evidence for critical skills", () => {
    const d = decideAssessment({ ...base, requested: "COMPETENT", competency: { isCritical: true, requiresEvidence: true } });
    expect(d.outcome).toBe("BLOCKED");
    expect(decideAssessment({ ...base, requested: "COMPETENT", competency: { isCritical: true, requiresEvidence: true }, evidenceCount: 1 }).outcome).toBe("APPLY");
  });
  it("Scenario 2: a supervisor's Competent request goes to pending verification, not straight to Competent", () => {
    expect(decideAssessment({ ...base, requested: "COMPETENT", actor: { id: "sup", canVerify: false } }).outcome).toBe("PENDING");
  });
  it("the evaluator who ran the practical cannot also verify it", () => {
    expect(decideAssessment({ ...base, requested: "COMPETENT", actor: { id: "trainer", canVerify: true } }).outcome).toBe("PENDING");
    expect(verificationBlockers({ verifierId: "trainer", verifierCanVerify: true, submittedById: "sup", practicalEvaluatorId: "trainer" })).toHaveLength(1);
    expect(verificationBlockers({ verifierId: "mgr", verifierCanVerify: true, submittedById: "sup", practicalEvaluatorId: "trainer" })).toHaveLength(0);
  });
  it("needs a training date for Trained and a note when lowering Competent", () => {
    expect(decideAssessment({ ...base, requested: "TRAINED", trainingDate: null }).outcome).toBe("BLOCKED");
    expect(decideAssessment({ ...base, requested: "TRAINED", trainingDate: now }).outcome).toBe("APPLY");
    expect(decideAssessment({ ...base, current: "COMPETENT", requested: "IN_TRAINING" }).outcome).toBe("BLOCKED");
    expect(decideAssessment({ ...base, current: "COMPETENT", requested: "IN_TRAINING", notes: "Failed audit observation" }).outcome).toBe("APPLY");
  });
});
