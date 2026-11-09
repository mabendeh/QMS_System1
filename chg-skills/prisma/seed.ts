/**
 * Seeds configuration plus FICTIONAL demonstration data.
 * Every demo employee has isSample = true, an employee number starting with
 * "DEMO-", and shows a "Sample" badge. Remove it all from
 * Administration → Sample data, or with `npm run demo:purge`.
 */
import { PrismaClient, type Role, type SkillStatus } from "@prisma/client";
import bcrypt from "bcryptjs";
import { randomBytes } from "node:crypto";
import { AREAS, CAREER_LEVELS, JOB_ROLES, PRIMARY_AREA_FOR_ROLE } from "./config-data";

const prisma = new PrismaClient();
const now = new Date();
const monthsAgo = (m: number) => new Date(now.getTime() - m * 30.44 * 86_400_000);
const daysFromNow = (d: number) => new Date(now.getTime() + d * 86_400_000);
const addMonths = (d: Date, m: number) => { const x = new Date(d); x.setMonth(x.getMonth() + m); return x; };

let seed = 20261109;
const rnd = () => { seed = (seed * 16807) % 2147483647; return seed / 2147483647; };

async function main() {
  const existing = await prisma.productionArea.count();
  if (existing > 0) {
    console.log("Database already has configuration. Run `npm run db:reset` to start over.");
    return;
  }
  const withDemo = process.env.SEED_SKIP_DEMO !== "true";

  // ---------------------------------------------------------------- reference data
  const shifts = await Promise.all(
    [["A", "Shift A"], ["B", "Shift B"], ["C", "Shift C"]].map(([code, name]) => prisma.shift.create({ data: { code, name } })),
  );
  const depts = await Promise.all(
    [["PROD", "Production"], ["SAN", "Sanitation"], ["WH", "Warehouse & Shipping"]].map(([code, name]) => prisma.department.create({ data: { code, name } })),
  );
  const deptByCode = Object.fromEntries(depts.map((d) => [d.code, d]));

  const areaByCode: Record<string, { id: string; name: string; kind: string }> = {};
  const compsByArea: Record<string, { id: string; name: string; isCritical: boolean }[]> = {};
  for (const [i, a] of AREAS.entries()) {
    const area = await prisma.productionArea.create({
      data: { code: a.code, name: a.name, kind: a.kind, isActive: a.active, description: a.description, sortOrder: (i + 1) * 10 },
    });
    areaByCode[a.code] = area;
    compsByArea[a.code] = [];
    for (const [j, [name, description, flags]] of a.competencies.entries()) {
      const critical = flags === "critical";
      const c = await prisma.competency.create({
        data: {
          areaId: area.id, code: `${a.code}-${String(j + 1).padStart(2, "0")}`, name, description,
          sortOrder: (j + 1) * 10, isCritical: critical, requiresEvidence: critical,
        },
      });
      compsByArea[a.code].push(c);
    }
  }
  const compByName = (area: string, name: string) => compsByArea[area].find((c) => c.name === name)!;

  // Requalification (demonstration values – confirm with Quality / Safety)
  await prisma.qualificationExpirationRule.createMany({
    data: [
      { areaId: areaByCode.FS.id, months: 12, notes: "Annual food safety refresher (demo assumption)" },
      { areaId: areaByCode.SA.id, months: 12, notes: "Annual safety refresher (demo assumption)" },
      { areaId: areaByCode.MD.id, months: 12, notes: "Annual CCP requalification (demo assumption)" },
      { competencyId: compByName("MX", "Allergen changeover").id, months: 12, notes: "Annual allergen changeover requalification (demo assumption)" },
    ],
  });
  const expiryMonths = (areaCode: string, compId: string) =>
    compId === compByName("MX", "Allergen changeover").id ? 12 : ["FS", "SA", "MD"].includes(areaCode) ? 12 : null;

  const jobRoles: Record<string, { id: string; title: string }> = {};
  for (const jr of JOB_ROLES) {
    const r = await prisma.jobRole.create({ data: { title: jr.title } });
    jobRoles[jr.title] = r;
    await prisma.competencyRequirement.createMany({
      data: jr.areas.flatMap((code) => compsByArea[code].map((c) => ({ jobRoleId: r.id, competencyId: c.id }))),
    });
  }

  const levels: Record<number, { id: string }> = {};
  for (const l of CAREER_LEVELS) levels[l.rank] = await prisma.careerLevel.create({ data: l });

  const rule = async (rank: number, notes: string, reqs: { type: string; label: string; areaId?: string; competencyId?: string; minCount?: number }[]) =>
    prisma.promotionRuleSet.create({
      data: {
        careerLevelId: levels[rank].id, version: 1, notes,
        requirements: { create: reqs.map((r, i) => ({ ...r, type: r.type as never, sortOrder: i })) },
      },
    });
  const DEMO_NOTE = "Demonstration rule set – NOT official CHG policy. HR must confirm before use.";
  await rule(1, DEMO_NOTE, [
    { type: "PRIMARY_AREA_QUALIFIED", label: "Qualified in assigned production area" },
    { type: "AREA_QUALIFIED", label: "Food safety training verified", areaId: areaByCode.FS.id },
    { type: "AREA_QUALIFIED", label: "Safety training verified", areaId: areaByCode.SA.id },
    { type: "MANUAL_VERIFICATION", label: "Supervisor approval" },
  ]);
  await rule(2, DEMO_NOTE, [
    { type: "MIN_PRODUCTION_AREAS_QUALIFIED", label: "Independently qualified in 2+ production areas", minCount: 2 },
    { type: "AREA_QUALIFIED", label: "Food safety training verified", areaId: areaByCode.FS.id },
    { type: "AREA_QUALIFIED", label: "Safety training verified", areaId: areaByCode.SA.id },
    { type: "MANUAL_VERIFICATION", label: "Supervisor approval" },
  ]);
  await rule(3, DEMO_NOTE, [
    { type: "CURRENT_LEVEL_AT_LEAST", label: "Currently Multi-Skilled Operator or above", minCount: 2 },
    { type: "MIN_PRODUCTION_AREAS_QUALIFIED", label: "Independently qualified in 3+ production areas", minCount: 3 },
    { type: "AREA_QUALIFIED", label: "Food safety training verified", areaId: areaByCode.FS.id },
    { type: "MANUAL_VERIFICATION", label: "Supervisor confirms independent problem-solving" },
  ]);
  await rule(4, DEMO_NOTE, [
    { type: "CURRENT_LEVEL_AT_LEAST", label: "Currently Advanced Operator or above", minCount: 3 },
    { type: "MANUAL_VERIFICATION", label: "Completed approved train-the-trainer program" },
    { type: "MANUAL_VERIFICATION", label: "Supervisor and HR approval" },
  ]);

  // Pay structure placeholders – NO amounts. HR enters approved values.
  for (const l of CAREER_LEVELS) {
    await prisma.payGrade.create({
      data: { code: `PG-${l.rank}`, name: `${l.name} (placeholder)`, careerLevelId: levels[l.rank].id, notes: "Not configured – HR to enter approved range." },
    });
  }

  for (const s of shifts) {
    for (const code of ["PW", "MX", "BL", "DD", "DF", "CL", "SF", "PK", "MD", "PL", "LS"]) {
      await prisma.shiftCoverageTarget.create({ data: { areaId: areaByCode[code].id, shiftId: s.id, targetQualified: code === "PK" ? 3 : 2 } });
    }
  }

  // ---------------------------------------------------------------- user accounts
  let password = process.env.SEED_DEMO_PASSWORD?.trim();
  let generated = false;
  if (!password) { password = randomBytes(9).toString("base64url"); generated = true; }
  const hash = await bcrypt.hash(password, 12);
  const users: Record<string, { id: string; name: string }> = {};
  const mk = async (key: string, email: string, name: string, roles: Role[]) => {
    users[key] = await prisma.user.create({
      data: { email, name, passwordHash: hash, roles: { create: roles.map((role) => ({ role })) } },
    });
  };
  await mk("hr", "hr@demo.local", "Demo HR Administrator", ["HR_ADMIN"]);
  await mk("pm", "manager@demo.local", "Demo Production Manager", ["PRODUCTION_MANAGER"]);
  await mk("supA", "supervisor.a@demo.local", "Demo Supervisor (Shift A)", ["SUPERVISOR"]);
  await mk("supB", "supervisor.b@demo.local", "Demo Supervisor (Shift B)", ["SUPERVISOR"]);
  await mk("supC", "supervisor.c@demo.local", "Demo Supervisor (Shift C)", ["SUPERVISOR"]);
  await mk("trainer", "trainer@demo.local", "Demo Trainer", ["TRAINER"]);
  await mk("ci", "ci@demo.local", "Demo CI Manager", ["CI_MANAGER"]);
  await mk("qa", "quality@demo.local", "Demo Quality Lead", ["QUALITY"]);
  await mk("lead", "leadership@demo.local", "Demo Plant Manager", ["LEADERSHIP"]);
  await mk("sys", "sysadmin@demo.local", "Demo System Administrator", ["SYSTEM_ADMIN"]);
  for (const code of ["PW", "MX", "BL", "DD", "DF", "SF", "PK", "MD", "PL", "FS", "SA", "LS"]) {
    await prisma.trainerAuthorization.create({ data: { userId: users.trainer.id, areaId: areaByCode[code].id } });
  }

  if (withDemo) await seedDemo();

  await prisma.auditLog.create({
    data: { action: "system.seed", entityType: "System", summary: withDemo ? "Configuration and sample data created" : "Configuration created" },
  });
  console.log("Seed complete.");
  console.log("Demo accounts: hr@, manager@, supervisor.a@ / .b@ / .c@, trainer@, ci@, quality@, leadership@, sysadmin@, employee@  (all @demo.local)");
  if (generated) console.log(`Generated demo password (shown once): ${password}`);
  else console.log("Demo password: the value of SEED_DEMO_PASSWORD in your .env file.");

  // ---------------------------------------------------------------- demo people
  async function seedDemo() {
    const first = ["Alex", "Jordan", "Taylor", "Casey", "Morgan", "Riley", "Jamie", "Avery", "Quinn", "Reese", "Skyler", "Dakota", "Rowan", "Emerson", "Finley", "Harper", "Kendall", "Logan", "Parker", "Sawyer", "Blake", "Cameron", "Drew", "Elliot", "Hayden", "Jesse", "Kai", "Lane", "Marley", "Noel", "Oakley", "Peyton", "Remy", "Sage", "Tatum", "Wren", "Ari", "Bailey"];
    const last = ["Sample", "Demo", "Example", "Testcase", "Placeholder", "Fictional", "Mockley", "Trialson"];
    const plan: [string, number][] = [];
    for (let s = 0; s < 3; s++) {
      for (const [role, n] of [["Pre-Weigh Operator", 2], ["Mixer Operator", 2], ["Blending Operator", 1], ["Forming Operator", 1], ["Freezer Operator", 1], ["Packaging Operator", 3], ["Palletizer Operator", 1], ["Line Lead", 1]] as [string, number][]) {
        for (let k = 0; k < n; k++) plan.push([role, s]);
      }
    }
    plan.push(["Production Associate", 2], ["Sanitation Technician", 1]);
    const sup = [users.supA, users.supB, users.supC];
    const cross: Record<string, string> = { PW: "MX", MX: "BL", BL: "PW", PK: "PL", PL: "PK", DF: "SF", SF: "DF", LS: "PK" };
    const skills: { employeeId: string; competencyId: string; status: SkillStatus; qualifiedAt?: Date | null; expiresAt?: Date | null; trainingDate?: Date | null; lastAssessedAt: Date; lastAssessedById: string }[] = [];
    const evals: { employeeId: string; competencyId: string; evaluatorUserId: string; evaluatedAt: Date; result: "PASS"; independentOperation: boolean; notes: string }[] = [];
    const history: { employeeId: string; competencyId: string; requestedStatus: SkillStatus; state: "APPLIED" | "VERIFIED"; submittedById: string; createdAt: Date; decidedAt: Date; notes: string }[] = [];
    const created: { id: string; role: string; shift: number; primary: string; name: string; supervisor: string }[] = [];

    for (const [i, [role, s]] of plan.entries()) {
      const primary = PRIMARY_AREA_FOR_ROLE[role];
      const name = { first: first[i % first.length], last: last[i % last.length] };
      const trainee = role === "Production Associate" || i % 13 === 7;
      const hireMonths = trainee ? 1 + rnd() * 3 : 6 + rnd() * 60;
      const emp = await prisma.employee.create({
        data: {
          employeeNumber: `DEMO-${1001 + i}`, firstName: name.first, lastName: name.last, isSample: true,
          shiftId: shifts[s].id, departmentId: deptByCode[role === "Sanitation Technician" ? "SAN" : role === "Palletizer Operator" ? "WH" : "PROD"].id,
          primaryAreaId: areaByCode[primary].id, jobRoleId: jobRoles[role].id, supervisorUserId: sup[s].id,
          hireDate: monthsAgo(hireMonths), careerLevelId: levels[trainee ? 0 : role === "Line Lead" ? 2 : 1].id,
        },
      });
      created.push({ id: emp.id, role, shift: s, primary, name: `${name.first} ${name.last}`, supervisor: sup[s].id });
      const areas = JOB_ROLES.find((j) => j.title === role)!.areas.slice();
      const crossArea = !trainee && rnd() < 0.35 ? cross[primary] : undefined;
      if (crossArea) areas.push(crossArea);
      for (const code of areas) {
        for (const c of compsByArea[code]) {
          const x = rnd();
          let status: SkillStatus;
          if (trainee) status = code === "FS" || code === "SA" ? (x < 0.6 ? "COMPETENT" : "IN_TRAINING") : x < 0.15 ? "TRAINED" : x < 0.7 ? "IN_TRAINING" : "NEEDS_TRAINING";
          else status = x < 0.86 ? "COMPETENT" : x < 0.92 ? "TRAINED" : x < 0.97 ? "IN_TRAINING" : "NEEDS_TRAINING";
          if (code === crossArea && !trainee && x > 0.9) status = "IN_TRAINING";
          if (status === "NEEDS_TRAINING") continue; // no record = red automatically for required skills
          const months = expiryMonths(code, c.id);
          const qualifiedAt = status === "COMPETENT" ? monthsAgo(months ? 0.5 + rnd() * 10.5 : 1 + rnd() * 30) : null;
          const assessedBy = sup[s].id;
          skills.push({
            employeeId: emp.id, competencyId: c.id, status, qualifiedAt,
            expiresAt: qualifiedAt && months ? addMonths(qualifiedAt, months) : null,
            trainingDate: status === "TRAINED" || status === "COMPETENT" ? new Date((qualifiedAt ?? monthsAgo(0.5)).getTime() - 7 * 86_400_000) : null,
            lastAssessedAt: qualifiedAt ?? monthsAgo(rnd()), lastAssessedById: assessedBy,
          });
          if (status === "COMPETENT" && qualifiedAt) {
            evals.push({ employeeId: emp.id, competencyId: c.id, evaluatorUserId: users.trainer.id, evaluatedAt: new Date(qualifiedAt.getTime() - 86_400_000), result: "PASS", independentOperation: true, notes: "Sample historical evaluation" });
            history.push({ employeeId: emp.id, competencyId: c.id, requestedStatus: "COMPETENT", state: "VERIFIED", submittedById: assessedBy, createdAt: qualifiedAt, decidedAt: qualifiedAt, notes: "Sample historical record" });
          }
        }
      }
    }

    // ---- named scenarios -------------------------------------------------------
    const pick = (role: string, shift: number, nth = 0) => created.filter((e) => e.role === role && e.shift === shift)[nth];
    const setAll = (empId: string, code: string, status: SkillStatus, qualifiedAt: Date | null) => {
      for (const c of compsByArea[code]) {
        const idx = skills.findIndex((k) => k.employeeId === empId && k.competencyId === c.id);
        const months = expiryMonths(code, c.id);
        const rec = {
          employeeId: empId, competencyId: c.id, status, qualifiedAt,
          expiresAt: qualifiedAt && months ? addMonths(qualifiedAt, months) : null,
          trainingDate: status === "COMPETENT" || status === "TRAINED" ? monthsAgo(4) : null,
          lastAssessedAt: qualifiedAt ?? monthsAgo(0.3), lastAssessedById: users.supA.id,
        };
        if (idx >= 0) skills[idx] = rec; else skills.push(rec);
        if (status === "COMPETENT" && qualifiedAt && !evals.some((e) => e.employeeId === empId && e.competencyId === c.id)) {
          evals.push({ employeeId: empId, competencyId: c.id, evaluatorUserId: users.trainer.id, evaluatedAt: qualifiedAt, result: "PASS", independentOperation: true, notes: "Sample historical evaluation" });
        }
      }
    };
    // 1. Multi-skilled: Pre-Weigh operator (Shift A) fully qualified in Pre-Weigh AND Mixing, still Level 1
    const multi = pick("Pre-Weigh Operator", 0);
    for (const code of ["PW", "MX", "FS", "SA"]) setAll(multi.id, code, "COMPETENT", monthsAgo(3));
    await prisma.employee.update({ where: { id: multi.id }, data: { firstName: "Alex", lastName: "Multiskill-Sample", careerLevelId: levels[1].id, hireDate: monthsAgo(30) } });
    // 2. Fully qualified Mixer operator (Shift B)
    const mixer = pick("Mixer Operator", 1);
    for (const code of ["MX", "FS", "SA"]) setAll(mixer.id, code, "COMPETENT", monthsAgo(5));
    // 3. Employee in training on Mixing (Shift A, second mixer)
    const trainee = pick("Mixer Operator", 0, 1);
    setAll(trainee.id, "MX", "IN_TRAINING", null);
    await prisma.employee.update({ where: { id: trainee.id }, data: { firstName: "Jordan", lastName: "Trainee-Sample", careerLevelId: levels[0].id, hireDate: monthsAgo(2) } });
    // E2E target: Mixing "Process troubleshooting" Trained, no practical yet (Shift A mixer #1)
    const e2eMixer = pick("Mixer Operator", 0);
    const pt = compByName("MX", "Process troubleshooting");
    const ptIdx = skills.findIndex((k) => k.employeeId === e2eMixer.id && k.competencyId === pt.id);
    const ptRec = { employeeId: e2eMixer.id, competencyId: pt.id, status: "TRAINED" as SkillStatus, qualifiedAt: null, expiresAt: null, trainingDate: monthsAgo(0.2), lastAssessedAt: monthsAgo(0.2), lastAssessedById: users.supA.id };
    if (ptIdx >= 0) skills[ptIdx] = ptRec; else skills.push(ptRec);
    for (let k = evals.length - 1; k >= 0; k--) if (evals[k].employeeId === e2eMixer.id && evals[k].competencyId === pt.id) evals.splice(k, 1);
    await prisma.employee.update({ where: { id: e2eMixer.id }, data: { firstName: "Casey", lastName: "Mixer-Sample" } });
    // 4. Pending qualification: Blending operator (Shift B), "Blend time monitoring" awaiting verification
    const blender = pick("Blending Operator", 1);
    for (const code of ["BL", "FS", "SA"]) setAll(blender.id, code, "COMPETENT", monthsAgo(6));
    const btm = compByName("BL", "Blend time monitoring");
    const bIdx = skills.findIndex((k) => k.employeeId === blender.id && k.competencyId === btm.id);
    skills[bIdx] = { ...skills[bIdx], status: "TRAINED", qualifiedAt: null, expiresAt: null, trainingDate: monthsAgo(0.4) };
    for (let k = evals.length - 1; k >= 0; k--) if (evals[k].employeeId === blender.id && evals[k].competencyId === btm.id) evals.splice(k, 1);
    await prisma.employee.update({ where: { id: blender.id }, data: { firstName: "Taylor", lastName: "Pending-Sample" } });
    // 5. Pending promotion: Packaging operator (Shift C), qualified in Packaging + Palletizing
    const packer = pick("Packaging Operator", 2);
    for (const code of ["PK", "MD", "PL", "FS", "SA"]) setAll(packer.id, code, "COMPETENT", monthsAgo(4));
    await prisma.employee.update({ where: { id: packer.id }, data: { firstName: "Morgan", lastName: "Promotion-Sample", careerLevelId: levels[1].id } });
    // 6. Expired: Freezer operator (Shift A) HACCP awareness lapsed; another expiring soon
    const freezer = pick("Freezer Operator", 0);
    for (const code of ["SF", "FS", "SA"]) setAll(freezer.id, code, "COMPETENT", monthsAgo(5));
    const haccp = compByName("FS", "HACCP / CCP awareness");
    const hIdx = skills.findIndex((k) => k.employeeId === freezer.id && k.competencyId === haccp.id);
    skills[hIdx] = { ...skills[hIdx], qualifiedAt: monthsAgo(13), expiresAt: addMonths(monthsAgo(13), 12) };
    const gmp = compByName("FS", "GMPs and personal hygiene");
    const gIdx = skills.findIndex((k) => k.employeeId === freezer.id && k.competencyId === gmp.id);
    skills[gIdx] = { ...skills[gIdx], qualifiedAt: addMonths(daysFromNow(14), -12), expiresAt: daysFromNow(14) };
    await prisma.employee.update({ where: { id: freezer.id }, data: { firstName: "Riley", lastName: "Expired-Sample" } });

    await prisma.employeeSkill.createMany({ data: skills });
    await prisma.practicalEvaluation.createMany({ data: evals });
    await prisma.skillAssessment.createMany({ data: history });

    // Pending verification for scenario 4
    const pe = await prisma.practicalEvaluation.create({
      data: { employeeId: blender.id, competencyId: btm.id, evaluatorUserId: users.trainer.id, evaluatedAt: monthsAgo(0.1), result: "PASS", independentOperation: true, notes: "Ran three batches independently (sample)." },
    });
    const pending = await prisma.skillAssessment.create({
      data: { employeeId: blender.id, competencyId: btm.id, previousStatus: "TRAINED", requestedStatus: "COMPETENT", state: "PENDING_VERIFICATION", submittedById: users.supB.id, practicalEvaluationId: pe.id, notes: "Ready for sign-off (sample)." },
    });
    await prisma.employeeSkill.update({
      where: { employeeId_competencyId: { employeeId: blender.id, competencyId: btm.id } },
      data: { pendingStatus: "COMPETENT", pendingAssessmentId: pending.id },
    });

    // Training assignments
    const ta = (employeeId: string, competencyId: string, status: "ASSIGNED" | "IN_PROGRESS" | "PRACTICAL_PASSED" | "COMPLETED", twiStep: "PREPARE" | "PRESENT" | "TRY_OUT" | "FOLLOW_UP", start: Date | null, due: Date, completed: Date | null = null) =>
      prisma.trainingAssignment.create({ data: { employeeId, competencyId, trainerUserId: users.trainer.id, assignedById: users.supA.id, status, twiStep, startDate: start, dueDate: due, completedDate: completed, notes: "Sample training assignment" } });
    for (const [k, c] of compsByArea.MX.slice(0, 5).entries()) await ta(trainee.id, c.id, "IN_PROGRESS", k < 2 ? "TRY_OUT" : "PRESENT", monthsAgo(1), daysFromNow(10 + k * 7));
    await ta(trainee.id, compsByArea.MX[5].id, "ASSIGNED", "PREPARE", null, daysFromNow(-5)); // overdue
    await ta(e2eMixer.id, pt.id, "IN_PROGRESS", "FOLLOW_UP", monthsAgo(0.5), daysFromNow(7));
    await ta(blender.id, btm.id, "PRACTICAL_PASSED", "FOLLOW_UP", monthsAgo(1), daysFromNow(3));
    await ta(freezer.id, haccp.id, "ASSIGNED", "PREPARE", null, daysFromNow(-2)); // requalification overdue
    for (const [k, c] of compsByArea.PL.slice(0, 3).entries()) await ta(packer.id, c.id, "COMPLETED", "FOLLOW_UP", monthsAgo(6 - k), monthsAgo(5 - k), monthsAgo(4.5 - k));

    // Pending promotion request (scenario: pending promotion review)
    const rs2 = await prisma.promotionRuleSet.findFirstOrThrow({ where: { careerLevelId: levels[2].id, isActive: true } });
    await prisma.promotionRequest.create({
      data: {
        employeeId: packer.id, fromLevelId: levels[1].id, toLevelId: levels[2].id, ruleSetId: rs2.id, status: "SUBMITTED",
        justification: "Qualified in Packaging and Palletizing; covers both stations on Shift C. (Sample request)",
        evaluationSnapshot: { note: "Sample request created by seed; re-evaluated live in the review screen." },
        submittedById: users.supC.id,
      },
    });

    // Employee self-service demo account linked to the multi-skilled sample employee
    await prisma.user.create({
      data: { email: "employee@demo.local", name: "Alex Multiskill-Sample", passwordHash: hash, employeeId: multi.id, roles: { create: [{ role: "EMPLOYEE" }] } },
    });
  }
}

main()
  .catch((e) => { console.error(e); process.exit(1); })
  .finally(() => prisma.$disconnect());
