-- CreateEnum
CREATE TYPE "Role" AS ENUM ('HR_ADMIN', 'PRODUCTION_MANAGER', 'SUPERVISOR', 'TRAINER', 'CI_MANAGER', 'QUALITY', 'LEADERSHIP', 'EMPLOYEE', 'SYSTEM_ADMIN');

-- CreateEnum
CREATE TYPE "SkillStatus" AS ENUM ('NOT_APPLICABLE', 'NEEDS_TRAINING', 'IN_TRAINING', 'TRAINED', 'COMPETENT');

-- CreateEnum
CREATE TYPE "AreaKind" AS ENUM ('PRODUCTION', 'PLANT_WIDE');

-- CreateEnum
CREATE TYPE "EmployeeStatus" AS ENUM ('ACTIVE', 'INACTIVE');

-- CreateEnum
CREATE TYPE "AssessmentState" AS ENUM ('APPLIED', 'PENDING_VERIFICATION', 'VERIFIED', 'REJECTED', 'SUPERSEDED');

-- CreateEnum
CREATE TYPE "EvaluationResult" AS ENUM ('PASS', 'FAIL');

-- CreateEnum
CREATE TYPE "TrainingStatus" AS ENUM ('ASSIGNED', 'IN_PROGRESS', 'PRACTICAL_PASSED', 'COMPLETED', 'CANCELLED');

-- CreateEnum
CREATE TYPE "TwiStep" AS ENUM ('PREPARE', 'PRESENT', 'TRY_OUT', 'FOLLOW_UP');

-- CreateEnum
CREATE TYPE "RequirementType" AS ENUM ('PRIMARY_AREA_QUALIFIED', 'AREA_QUALIFIED', 'MIN_PRODUCTION_AREAS_QUALIFIED', 'COMPETENCY_COMPETENT', 'CURRENT_LEVEL_AT_LEAST', 'MANUAL_VERIFICATION');

-- CreateEnum
CREATE TYPE "PromotionStatus" AS ENUM ('SUBMITTED', 'MANAGER_APPROVED', 'HR_VERIFIED', 'APPROVED', 'REJECTED', 'WITHDRAWN');

-- CreateEnum
CREATE TYPE "ApprovalStep" AS ENUM ('QUALIFICATION_VERIFICATION', 'MANAGER_REVIEW', 'HR_VERIFICATION', 'FINAL_APPROVAL');

-- CreateEnum
CREATE TYPE "Decision" AS ENUM ('APPROVED', 'REJECTED');

-- CreateTable
CREATE TABLE "users" (
    "id" TEXT NOT NULL,
    "email" TEXT NOT NULL,
    "name" TEXT NOT NULL,
    "passwordHash" TEXT,
    "isActive" BOOLEAN NOT NULL DEFAULT true,
    "employeeId" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "users_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "user_roles" (
    "id" TEXT NOT NULL,
    "userId" TEXT NOT NULL,
    "role" "Role" NOT NULL,

    CONSTRAINT "user_roles_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "sessions" (
    "id" TEXT NOT NULL,
    "userId" TEXT NOT NULL,
    "expiresAt" TIMESTAMP(3) NOT NULL,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "lastSeenAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "userAgent" TEXT,

    CONSTRAINT "sessions_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "shifts" (
    "id" TEXT NOT NULL,
    "code" TEXT NOT NULL,
    "name" TEXT NOT NULL,
    "isActive" BOOLEAN NOT NULL DEFAULT true,

    CONSTRAINT "shifts_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "departments" (
    "id" TEXT NOT NULL,
    "code" TEXT NOT NULL,
    "name" TEXT NOT NULL,
    "isActive" BOOLEAN NOT NULL DEFAULT true,

    CONSTRAINT "departments_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "production_areas" (
    "id" TEXT NOT NULL,
    "code" TEXT NOT NULL,
    "name" TEXT NOT NULL,
    "description" TEXT,
    "kind" "AreaKind" NOT NULL DEFAULT 'PRODUCTION',
    "sortOrder" INTEGER NOT NULL DEFAULT 0,
    "isActive" BOOLEAN NOT NULL DEFAULT true,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "production_areas_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "job_roles" (
    "id" TEXT NOT NULL,
    "title" TEXT NOT NULL,
    "description" TEXT,
    "isActive" BOOLEAN NOT NULL DEFAULT true,

    CONSTRAINT "job_roles_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "competencies" (
    "id" TEXT NOT NULL,
    "areaId" TEXT NOT NULL,
    "code" TEXT NOT NULL,
    "name" TEXT NOT NULL,
    "description" TEXT,
    "sortOrder" INTEGER NOT NULL DEFAULT 0,
    "isActive" BOOLEAN NOT NULL DEFAULT true,
    "isCritical" BOOLEAN NOT NULL DEFAULT false,
    "requiresEvidence" BOOLEAN NOT NULL DEFAULT false,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "competencies_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "competency_requirements" (
    "id" TEXT NOT NULL,
    "jobRoleId" TEXT NOT NULL,
    "competencyId" TEXT NOT NULL,

    CONSTRAINT "competency_requirements_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "career_levels" (
    "id" TEXT NOT NULL,
    "rank" INTEGER NOT NULL,
    "name" TEXT NOT NULL,
    "description" TEXT,
    "isActive" BOOLEAN NOT NULL DEFAULT true,

    CONSTRAINT "career_levels_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "employees" (
    "id" TEXT NOT NULL,
    "employeeNumber" TEXT NOT NULL,
    "firstName" TEXT NOT NULL,
    "lastName" TEXT NOT NULL,
    "shiftId" TEXT NOT NULL,
    "departmentId" TEXT,
    "primaryAreaId" TEXT,
    "jobRoleId" TEXT,
    "supervisorUserId" TEXT,
    "careerLevelId" TEXT,
    "hireDate" TIMESTAMP(3),
    "status" "EmployeeStatus" NOT NULL DEFAULT 'ACTIVE',
    "isSample" BOOLEAN NOT NULL DEFAULT false,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "employees_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "qualification_records" (
    "id" TEXT NOT NULL,
    "employeeId" TEXT NOT NULL,
    "competencyId" TEXT NOT NULL,
    "status" "SkillStatus" NOT NULL,
    "pendingStatus" "SkillStatus",
    "pendingAssessmentId" TEXT,
    "trainingDate" TIMESTAMP(3),
    "qualifiedAt" TIMESTAMP(3),
    "expiresAt" TIMESTAMP(3),
    "lastAssessedAt" TIMESTAMP(3),
    "lastAssessedById" TEXT,
    "notes" TEXT,
    "version" INTEGER NOT NULL DEFAULT 1,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "qualification_records_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "employee_assessments" (
    "id" TEXT NOT NULL,
    "employeeId" TEXT NOT NULL,
    "competencyId" TEXT NOT NULL,
    "previousStatus" "SkillStatus",
    "requestedStatus" "SkillStatus" NOT NULL,
    "state" "AssessmentState" NOT NULL,
    "notes" TEXT,
    "trainingDate" TIMESTAMP(3),
    "evaluatorName" TEXT,
    "submittedById" TEXT NOT NULL,
    "practicalEvaluationId" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "decidedAt" TIMESTAMP(3),

    CONSTRAINT "employee_assessments_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "practical_evaluations" (
    "id" TEXT NOT NULL,
    "employeeId" TEXT NOT NULL,
    "competencyId" TEXT NOT NULL,
    "trainingAssignmentId" TEXT,
    "evaluatorUserId" TEXT NOT NULL,
    "evaluatedAt" TIMESTAMP(3) NOT NULL,
    "result" "EvaluationResult" NOT NULL,
    "independentOperation" BOOLEAN NOT NULL DEFAULT false,
    "notes" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "practical_evaluations_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "trainers" (
    "id" TEXT NOT NULL,
    "userId" TEXT NOT NULL,
    "areaId" TEXT NOT NULL,
    "authorizedAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "validUntil" TIMESTAMP(3),

    CONSTRAINT "trainers_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "training_assignments" (
    "id" TEXT NOT NULL,
    "employeeId" TEXT NOT NULL,
    "competencyId" TEXT NOT NULL,
    "trainerUserId" TEXT,
    "assignedById" TEXT NOT NULL,
    "status" "TrainingStatus" NOT NULL DEFAULT 'ASSIGNED',
    "twiStep" "TwiStep" NOT NULL DEFAULT 'PREPARE',
    "startDate" TIMESTAMP(3),
    "dueDate" TIMESTAMP(3),
    "completedDate" TIMESTAMP(3),
    "notes" TEXT,
    "version" INTEGER NOT NULL DEFAULT 1,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "training_assignments_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "training_evidence" (
    "id" TEXT NOT NULL,
    "employeeId" TEXT NOT NULL,
    "competencyId" TEXT NOT NULL,
    "assessmentId" TEXT,
    "trainingAssignmentId" TEXT,
    "practicalEvaluationId" TEXT,
    "originalName" TEXT NOT NULL,
    "storedName" TEXT NOT NULL,
    "mimeType" TEXT NOT NULL,
    "sizeBytes" INTEGER NOT NULL,
    "sha256" TEXT NOT NULL,
    "uploadedById" TEXT NOT NULL,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "training_evidence_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "qualification_expiration_rules" (
    "id" TEXT NOT NULL,
    "areaId" TEXT,
    "competencyId" TEXT,
    "months" INTEGER NOT NULL,
    "notes" TEXT,

    CONSTRAINT "qualification_expiration_rules_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "shift_coverage_targets" (
    "id" TEXT NOT NULL,
    "areaId" TEXT NOT NULL,
    "shiftId" TEXT NOT NULL,
    "targetQualified" INTEGER NOT NULL,

    CONSTRAINT "shift_coverage_targets_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "promotion_rules" (
    "id" TEXT NOT NULL,
    "careerLevelId" TEXT NOT NULL,
    "version" INTEGER NOT NULL,
    "isActive" BOOLEAN NOT NULL DEFAULT true,
    "effectiveFrom" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "notes" TEXT,
    "createdById" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "promotion_rules_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "promotion_requirements" (
    "id" TEXT NOT NULL,
    "ruleSetId" TEXT NOT NULL,
    "type" "RequirementType" NOT NULL,
    "areaId" TEXT,
    "competencyId" TEXT,
    "minCount" INTEGER,
    "label" TEXT NOT NULL,
    "sortOrder" INTEGER NOT NULL DEFAULT 0,

    CONSTRAINT "promotion_requirements_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "promotion_requests" (
    "id" TEXT NOT NULL,
    "employeeId" TEXT NOT NULL,
    "fromLevelId" TEXT,
    "toLevelId" TEXT NOT NULL,
    "ruleSetId" TEXT NOT NULL,
    "status" "PromotionStatus" NOT NULL DEFAULT 'SUBMITTED',
    "justification" TEXT NOT NULL,
    "evaluationSnapshot" JSONB NOT NULL,
    "submittedById" TEXT NOT NULL,
    "submittedAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "decisionDate" TIMESTAMP(3),
    "effectiveDate" TIMESTAMP(3),
    "version" INTEGER NOT NULL DEFAULT 1,

    CONSTRAINT "promotion_requests_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "approval_records" (
    "id" TEXT NOT NULL,
    "promotionRequestId" TEXT,
    "assessmentId" TEXT,
    "step" "ApprovalStep" NOT NULL,
    "decision" "Decision" NOT NULL,
    "comments" TEXT,
    "decidedById" TEXT NOT NULL,
    "decidedAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "approval_records_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "pay_grades" (
    "id" TEXT NOT NULL,
    "code" TEXT NOT NULL,
    "name" TEXT NOT NULL,
    "careerLevelId" TEXT,
    "minRate" DECIMAL(10,2),
    "maxRate" DECIMAL(10,2),
    "currency" TEXT NOT NULL DEFAULT 'USD',
    "effectiveDate" TIMESTAMP(3),
    "endDate" TIMESTAMP(3),
    "approvalAuthority" TEXT,
    "notes" TEXT,
    "isActive" BOOLEAN NOT NULL DEFAULT true,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "pay_grades_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "skill_premiums" (
    "id" TEXT NOT NULL,
    "name" TEXT NOT NULL,
    "areaId" TEXT,
    "amount" DECIMAL(10,2),
    "unit" TEXT NOT NULL DEFAULT 'per hour',
    "effectiveDate" TIMESTAMP(3),
    "endDate" TIMESTAMP(3),
    "approvalAuthority" TEXT,
    "notes" TEXT,
    "isActive" BOOLEAN NOT NULL DEFAULT true,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "skill_premiums_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "audit_logs" (
    "id" BIGSERIAL NOT NULL,
    "at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "actorUserId" TEXT,
    "actorName" TEXT,
    "action" TEXT NOT NULL,
    "entityType" TEXT NOT NULL,
    "entityId" TEXT,
    "summary" TEXT NOT NULL,
    "before" JSONB,
    "after" JSONB,

    CONSTRAINT "audit_logs_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE UNIQUE INDEX "users_email_key" ON "users"("email");

-- CreateIndex
CREATE UNIQUE INDEX "users_employeeId_key" ON "users"("employeeId");

-- CreateIndex
CREATE UNIQUE INDEX "user_roles_userId_role_key" ON "user_roles"("userId", "role");

-- CreateIndex
CREATE INDEX "sessions_userId_idx" ON "sessions"("userId");

-- CreateIndex
CREATE UNIQUE INDEX "shifts_code_key" ON "shifts"("code");

-- CreateIndex
CREATE UNIQUE INDEX "departments_code_key" ON "departments"("code");

-- CreateIndex
CREATE UNIQUE INDEX "departments_name_key" ON "departments"("name");

-- CreateIndex
CREATE UNIQUE INDEX "production_areas_code_key" ON "production_areas"("code");

-- CreateIndex
CREATE UNIQUE INDEX "production_areas_name_key" ON "production_areas"("name");

-- CreateIndex
CREATE INDEX "production_areas_isActive_sortOrder_idx" ON "production_areas"("isActive", "sortOrder");

-- CreateIndex
CREATE UNIQUE INDEX "job_roles_title_key" ON "job_roles"("title");

-- CreateIndex
CREATE UNIQUE INDEX "competencies_code_key" ON "competencies"("code");

-- CreateIndex
CREATE INDEX "competencies_areaId_sortOrder_idx" ON "competencies"("areaId", "sortOrder");

-- CreateIndex
CREATE UNIQUE INDEX "competencies_areaId_name_key" ON "competencies"("areaId", "name");

-- CreateIndex
CREATE UNIQUE INDEX "competency_requirements_jobRoleId_competencyId_key" ON "competency_requirements"("jobRoleId", "competencyId");

-- CreateIndex
CREATE UNIQUE INDEX "career_levels_rank_key" ON "career_levels"("rank");

-- CreateIndex
CREATE UNIQUE INDEX "career_levels_name_key" ON "career_levels"("name");

-- CreateIndex
CREATE UNIQUE INDEX "employees_employeeNumber_key" ON "employees"("employeeNumber");

-- CreateIndex
CREATE INDEX "employees_shiftId_idx" ON "employees"("shiftId");

-- CreateIndex
CREATE INDEX "employees_primaryAreaId_idx" ON "employees"("primaryAreaId");

-- CreateIndex
CREATE INDEX "employees_supervisorUserId_idx" ON "employees"("supervisorUserId");

-- CreateIndex
CREATE INDEX "employees_lastName_firstName_idx" ON "employees"("lastName", "firstName");

-- CreateIndex
CREATE INDEX "qualification_records_competencyId_status_idx" ON "qualification_records"("competencyId", "status");

-- CreateIndex
CREATE INDEX "qualification_records_expiresAt_idx" ON "qualification_records"("expiresAt");

-- CreateIndex
CREATE UNIQUE INDEX "qualification_records_employeeId_competencyId_key" ON "qualification_records"("employeeId", "competencyId");

-- CreateIndex
CREATE INDEX "employee_assessments_employeeId_competencyId_createdAt_idx" ON "employee_assessments"("employeeId", "competencyId", "createdAt");

-- CreateIndex
CREATE INDEX "employee_assessments_state_idx" ON "employee_assessments"("state");

-- CreateIndex
CREATE INDEX "practical_evaluations_employeeId_competencyId_idx" ON "practical_evaluations"("employeeId", "competencyId");

-- CreateIndex
CREATE UNIQUE INDEX "trainers_userId_areaId_key" ON "trainers"("userId", "areaId");

-- CreateIndex
CREATE INDEX "training_assignments_employeeId_status_idx" ON "training_assignments"("employeeId", "status");

-- CreateIndex
CREATE INDEX "training_assignments_trainerUserId_status_idx" ON "training_assignments"("trainerUserId", "status");

-- CreateIndex
CREATE INDEX "training_assignments_dueDate_idx" ON "training_assignments"("dueDate");

-- CreateIndex
CREATE UNIQUE INDEX "training_evidence_storedName_key" ON "training_evidence"("storedName");

-- CreateIndex
CREATE INDEX "training_evidence_employeeId_competencyId_idx" ON "training_evidence"("employeeId", "competencyId");

-- CreateIndex
CREATE UNIQUE INDEX "qualification_expiration_rules_areaId_competencyId_key" ON "qualification_expiration_rules"("areaId", "competencyId");

-- CreateIndex
CREATE UNIQUE INDEX "shift_coverage_targets_areaId_shiftId_key" ON "shift_coverage_targets"("areaId", "shiftId");

-- CreateIndex
CREATE UNIQUE INDEX "promotion_rules_careerLevelId_version_key" ON "promotion_rules"("careerLevelId", "version");

-- CreateIndex
CREATE INDEX "promotion_requirements_ruleSetId_idx" ON "promotion_requirements"("ruleSetId");

-- CreateIndex
CREATE INDEX "promotion_requests_status_idx" ON "promotion_requests"("status");

-- CreateIndex
CREATE INDEX "promotion_requests_employeeId_idx" ON "promotion_requests"("employeeId");

-- CreateIndex
CREATE INDEX "approval_records_promotionRequestId_idx" ON "approval_records"("promotionRequestId");

-- CreateIndex
CREATE INDEX "approval_records_assessmentId_idx" ON "approval_records"("assessmentId");

-- CreateIndex
CREATE UNIQUE INDEX "pay_grades_code_key" ON "pay_grades"("code");

-- CreateIndex
CREATE UNIQUE INDEX "skill_premiums_name_key" ON "skill_premiums"("name");

-- CreateIndex
CREATE INDEX "audit_logs_entityType_entityId_idx" ON "audit_logs"("entityType", "entityId");

-- CreateIndex
CREATE INDEX "audit_logs_at_idx" ON "audit_logs"("at");

-- AddForeignKey
ALTER TABLE "users" ADD CONSTRAINT "users_employeeId_fkey" FOREIGN KEY ("employeeId") REFERENCES "employees"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "user_roles" ADD CONSTRAINT "user_roles_userId_fkey" FOREIGN KEY ("userId") REFERENCES "users"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "sessions" ADD CONSTRAINT "sessions_userId_fkey" FOREIGN KEY ("userId") REFERENCES "users"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "competencies" ADD CONSTRAINT "competencies_areaId_fkey" FOREIGN KEY ("areaId") REFERENCES "production_areas"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "competency_requirements" ADD CONSTRAINT "competency_requirements_jobRoleId_fkey" FOREIGN KEY ("jobRoleId") REFERENCES "job_roles"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "competency_requirements" ADD CONSTRAINT "competency_requirements_competencyId_fkey" FOREIGN KEY ("competencyId") REFERENCES "competencies"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "employees" ADD CONSTRAINT "employees_shiftId_fkey" FOREIGN KEY ("shiftId") REFERENCES "shifts"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "employees" ADD CONSTRAINT "employees_departmentId_fkey" FOREIGN KEY ("departmentId") REFERENCES "departments"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "employees" ADD CONSTRAINT "employees_primaryAreaId_fkey" FOREIGN KEY ("primaryAreaId") REFERENCES "production_areas"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "employees" ADD CONSTRAINT "employees_jobRoleId_fkey" FOREIGN KEY ("jobRoleId") REFERENCES "job_roles"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "employees" ADD CONSTRAINT "employees_supervisorUserId_fkey" FOREIGN KEY ("supervisorUserId") REFERENCES "users"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "employees" ADD CONSTRAINT "employees_careerLevelId_fkey" FOREIGN KEY ("careerLevelId") REFERENCES "career_levels"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "qualification_records" ADD CONSTRAINT "qualification_records_employeeId_fkey" FOREIGN KEY ("employeeId") REFERENCES "employees"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "qualification_records" ADD CONSTRAINT "qualification_records_competencyId_fkey" FOREIGN KEY ("competencyId") REFERENCES "competencies"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "employee_assessments" ADD CONSTRAINT "employee_assessments_employeeId_fkey" FOREIGN KEY ("employeeId") REFERENCES "employees"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "employee_assessments" ADD CONSTRAINT "employee_assessments_competencyId_fkey" FOREIGN KEY ("competencyId") REFERENCES "competencies"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "practical_evaluations" ADD CONSTRAINT "practical_evaluations_employeeId_fkey" FOREIGN KEY ("employeeId") REFERENCES "employees"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "practical_evaluations" ADD CONSTRAINT "practical_evaluations_competencyId_fkey" FOREIGN KEY ("competencyId") REFERENCES "competencies"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "practical_evaluations" ADD CONSTRAINT "practical_evaluations_trainingAssignmentId_fkey" FOREIGN KEY ("trainingAssignmentId") REFERENCES "training_assignments"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "trainers" ADD CONSTRAINT "trainers_userId_fkey" FOREIGN KEY ("userId") REFERENCES "users"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "trainers" ADD CONSTRAINT "trainers_areaId_fkey" FOREIGN KEY ("areaId") REFERENCES "production_areas"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "training_assignments" ADD CONSTRAINT "training_assignments_employeeId_fkey" FOREIGN KEY ("employeeId") REFERENCES "employees"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "training_assignments" ADD CONSTRAINT "training_assignments_competencyId_fkey" FOREIGN KEY ("competencyId") REFERENCES "competencies"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "training_evidence" ADD CONSTRAINT "training_evidence_employeeId_fkey" FOREIGN KEY ("employeeId") REFERENCES "employees"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "training_evidence" ADD CONSTRAINT "training_evidence_competencyId_fkey" FOREIGN KEY ("competencyId") REFERENCES "competencies"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "training_evidence" ADD CONSTRAINT "training_evidence_assessmentId_fkey" FOREIGN KEY ("assessmentId") REFERENCES "employee_assessments"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "training_evidence" ADD CONSTRAINT "training_evidence_trainingAssignmentId_fkey" FOREIGN KEY ("trainingAssignmentId") REFERENCES "training_assignments"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "training_evidence" ADD CONSTRAINT "training_evidence_practicalEvaluationId_fkey" FOREIGN KEY ("practicalEvaluationId") REFERENCES "practical_evaluations"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "qualification_expiration_rules" ADD CONSTRAINT "qualification_expiration_rules_areaId_fkey" FOREIGN KEY ("areaId") REFERENCES "production_areas"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "qualification_expiration_rules" ADD CONSTRAINT "qualification_expiration_rules_competencyId_fkey" FOREIGN KEY ("competencyId") REFERENCES "competencies"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "shift_coverage_targets" ADD CONSTRAINT "shift_coverage_targets_areaId_fkey" FOREIGN KEY ("areaId") REFERENCES "production_areas"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "shift_coverage_targets" ADD CONSTRAINT "shift_coverage_targets_shiftId_fkey" FOREIGN KEY ("shiftId") REFERENCES "shifts"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "promotion_rules" ADD CONSTRAINT "promotion_rules_careerLevelId_fkey" FOREIGN KEY ("careerLevelId") REFERENCES "career_levels"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "promotion_requirements" ADD CONSTRAINT "promotion_requirements_ruleSetId_fkey" FOREIGN KEY ("ruleSetId") REFERENCES "promotion_rules"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "promotion_requests" ADD CONSTRAINT "promotion_requests_employeeId_fkey" FOREIGN KEY ("employeeId") REFERENCES "employees"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "promotion_requests" ADD CONSTRAINT "promotion_requests_fromLevelId_fkey" FOREIGN KEY ("fromLevelId") REFERENCES "career_levels"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "promotion_requests" ADD CONSTRAINT "promotion_requests_toLevelId_fkey" FOREIGN KEY ("toLevelId") REFERENCES "career_levels"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "promotion_requests" ADD CONSTRAINT "promotion_requests_ruleSetId_fkey" FOREIGN KEY ("ruleSetId") REFERENCES "promotion_rules"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "approval_records" ADD CONSTRAINT "approval_records_promotionRequestId_fkey" FOREIGN KEY ("promotionRequestId") REFERENCES "promotion_requests"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "approval_records" ADD CONSTRAINT "approval_records_assessmentId_fkey" FOREIGN KEY ("assessmentId") REFERENCES "employee_assessments"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "pay_grades" ADD CONSTRAINT "pay_grades_careerLevelId_fkey" FOREIGN KEY ("careerLevelId") REFERENCES "career_levels"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- Append-only enforcement: audit and approval history can never be changed or removed.
CREATE OR REPLACE FUNCTION prevent_history_change() RETURNS trigger AS $$
BEGIN
  IF current_setting('chg.allow_history_purge', true) = 'on' THEN
    RETURN OLD; -- only used by the guarded "delete sample data" admin action
  END IF;
  RAISE EXCEPTION 'History table % is append-only', TG_TABLE_NAME;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER audit_logs_no_update BEFORE UPDATE OR DELETE ON "audit_logs"
  FOR EACH ROW EXECUTE FUNCTION prevent_history_change();
CREATE TRIGGER approval_records_no_update BEFORE UPDATE OR DELETE ON "approval_records"
  FOR EACH ROW EXECUTE FUNCTION prevent_history_change();
