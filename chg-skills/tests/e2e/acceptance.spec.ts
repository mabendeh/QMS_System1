import { expect, test, devices } from "@playwright/test";
import ExcelJS from "exceljs";
import { cell, chipStatus, comp, db, empByLast, expectAudit, login, pickStatus } from "./helpers";

test.describe.configure({ mode: "serial" });
test.afterAll(async () => { await db.$disconnect(); });

test("Scenario 1: supervisor moves a skill from Needs Training to In Training; cell turns yellow and is audited", async ({ page }) => {
  const supA = await db.user.findUniqueOrThrow({ where: { email: "supervisor.a@demo.local" } });
  // Find a required, unassessed (red) skill for one of Supervisor A's employees
  const emps = await db.employee.findMany({ where: { supervisorUserId: supA.id, isSample: true }, include: { jobRole: { include: { requirements: { include: { competency: { include: { area: true } } } } } }, skills: true } });
  type Emp = (typeof emps)[number];
  type Req = NonNullable<Emp["jobRole"]>["requirements"][number];
  let emp: Emp | undefined;
  let c: Req | undefined;
  for (const e of emps) {
    const r = e.jobRole?.requirements.find((q) => q.competency.area.kind === "PRODUCTION" && !e.skills.some((s) => s.competencyId === q.competencyId));
    if (r) { emp = e; c = r; break; }
  }
  if (!emp || !c) throw new Error("No red cell found for Supervisor A in sample data");
  const name = `${emp.firstName} ${emp.lastName}`;
  await login(page, "supervisor.a");
  await page.goto(`/matrix?area=${c.competency.area.code}&focus=${emp.id}`);
  expect(await chipStatus(page, name, c.competency.name)).toBe("NEEDS_TRAINING");
  await cell(page, name, c.competency.name).click();
  await pickStatus(page, "In Training");
  await page.getByRole("button", { name: "Save assessment" }).click();
  await expect(page.getByText(/updated to In Training/)).toBeVisible();
  await page.keyboard.press("Escape");
  await expect.poll(() => chipStatus(page, name, c.competency.name)).toBe("IN_TRAINING");
  await expectAudit(`set ${name} – ${c.competency.name} → In Training`);
});

test("Scenario 2: trainer passes practical, supervisor requests Competent, authorized evaluator verifies; cell turns blue", async ({ browser }) => {
  const emp = await empByLast("Mixer-Sample");
  const c = await comp("MX", "Process troubleshooting");
  const name = `${emp.firstName} ${emp.lastName}`;
  const url = `/matrix?area=MX&focus=${emp.id}`;

  const trainer = await browser.newPage();
  await login(trainer, "trainer");
  await trainer.goto(url);
  expect(await chipStatus(trainer, name, c.name)).toBe("TRAINED");
  await cell(trainer, name, c.name).click();
  await trainer.getByRole("tab", { name: "Practical evaluation" }).click();
  await trainer.getByLabel("Passed", { exact: true }).check();
  await trainer.getByLabel("Operated independently without help").check();
  await trainer.getByRole("button", { name: "Record evaluation" }).click();
  await expect(trainer.getByText(/Practical evaluation recorded/)).toBeVisible();

  const sup = await browser.newPage();
  await login(sup, "supervisor.a");
  await sup.goto(url);
  await cell(sup, name, c.name).click();
  await pickStatus(sup, "Competent");
  await sup.getByRole("button", { name: "Submit for verification" }).click();
  await expect(sup.getByText("Submitted for verification")).toBeVisible();
  await sup.keyboard.press("Escape");
  await expect(cell(sup, name, c.name).locator("[data-pending=true]")).toBeVisible();
  expect(await chipStatus(sup, name, c.name)).toBe("TRAINED"); // not prematurely blue

  const mgr = await browser.newPage();
  await login(mgr, "manager");
  await mgr.goto(url);
  await cell(mgr, name, c.name).click();
  await mgr.getByRole("button", { name: "Verify as Competent" }).click();
  await expect(mgr.getByText(/is now Competent/)).toBeVisible();
  await mgr.keyboard.press("Escape");
  await expect.poll(() => chipStatus(mgr, name, c.name)).toBe("COMPETENT");
  await expectAudit(`verified Competent for ${name} – ${c.name}`);
  const approvals = await db.approvalRecord.count({ where: { assessment: { employeeId: emp.id, competencyId: c.id }, step: "QUALIFICATION_VERIFICATION" } });
  expect(approvals).toBe(1);
});

test("Scenario 3: Pre-Weigh + Mixing qualification flags progression eligibility without any pay change", async ({ page }) => {
  const emp = await empByLast("Multiskill-Sample");
  await login(page, "manager");
  await page.goto(`/employees/${emp.id}`);
  await expect(page.getByText("To reach")).toBeVisible();
  await expect(page.getByText("Multi-Skilled Operator").first()).toBeVisible();
  await expect(page.getByText(/All system-checked requirements are met/)).toBeVisible();
  await expect(page.getByText("Level 1 – Qualified Operator")).toBeVisible();
  await page.goto("/matrix");
  await expect(page.locator(`#row-${emp.id}`).getByText("Promotion-ready")).toBeVisible();
  const after = await db.employee.findUniqueOrThrow({ where: { id: emp.id }, include: { careerLevel: true } });
  expect(after.careerLevel?.rank).toBe(1);
  expect(await db.promotionRequest.count({ where: { employeeId: emp.id } })).toBe(0);
});

test("Scenario 4: a supervisor cannot reach compensation data", async ({ page }) => {
  await login(page, "supervisor.a");
  await page.goto("/hr/compensation");
  await expect(page).toHaveURL(/\/denied$/);
  await expect(page.getByRole("heading", { name: "Access denied" })).toBeVisible();
  await page.goto("/");
  await expect(page.getByRole("link", { name: "HR Compensation" })).toHaveCount(0);
  const emp = await empByLast("Multiskill-Sample");
  await page.goto(`/employees/${emp.id}`);
  await expect(page.getByText("Pay grade (HR only)")).toHaveCount(0);
  // System administrators do not see pay either
  await login(page, "sysadmin");
  await page.goto("/hr/compensation");
  await expect(page).toHaveURL(/\/denied$/);
});

test("Scenario 5: an expired qualification stops counting toward qualified coverage", async ({ page }) => {
  await login(page, "manager");
  const md = await db.productionArea.findUniqueOrThrow({ where: { code: "MD" } });
  const mdComps = await db.competency.findMany({ where: { areaId: md.id, isActive: true } });
  // someone currently qualified in every Metal Detection competency
  const holders = await db.employee.findMany({ where: { skills: { some: { competencyId: mdComps[0].id, status: "COMPETENT" } } }, include: { skills: true, shift: true } });
  const holder = holders.find((h) => mdComps.every((c) => h.skills.some((s) => s.competencyId === c.id && s.status === "COMPETENT" && (!s.expiresAt || s.expiresAt > new Date()))))!;
  expect(holder).toBeTruthy();
  const coverageCell = () => page.getByRole("row").filter({ hasText: "Metal Detection / Inspection" }).locator("td").nth(1 + ["A", "B", "C"].indexOf(holder.shift.code));
  await page.goto("/coverage");
  const before = Number((await coverageCell().innerText()).split("/")[0].trim());
  const rec = holder.skills.find((s) => s.competencyId === mdComps[0].id)!;
  await db.employeeSkill.update({ where: { id: rec.id }, data: { expiresAt: new Date(Date.now() - 86_400_000) } });
  await page.reload();
  const after = Number((await coverageCell().innerText()).split("/")[0].trim());
  expect(after).toBe(before - 1);
  await page.goto(`/matrix?area=MD&focus=${holder.id}`);
  expect(await chipStatus(page, `${holder.firstName} ${holder.lastName}`, mdComps[0].name)).toBe("EXPIRED");
});

test("Scenario 6: HR exports the matrix to Excel with accurate data, department columns, statuses and colours", async ({ page }) => {
  await login(page, "hr");
  const res = await page.request.get("/api/export/matrix?format=xlsx");
  expect(res.status()).toBe(200);
  expect(res.headers()["content-type"]).toContain("spreadsheetml");
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load((await res.body()) as unknown as ArrayBuffer);
  const overview = wb.getWorksheet("Overview")!;
  const header = (overview.getRow(4).values as string[]).filter(Boolean);
  for (const h of ["Employee ID", "Employee", "Shift", "Department", "Job title", "Pre-Weigh", "Mixing", "Blending", "Packaging"]) expect(header).toContain(h);
  const active = await db.employee.count({ where: { status: "ACTIVE" } });
  let rows = 0;
  overview.eachRow((r, n) => { if (n > 4 && String(r.getCell(1).value ?? "").startsWith("DEMO-")) rows++; });
  expect(rows).toBe(active);
  // A known fully-qualified department cell is blue "C"
  const emp = await empByLast("Multiskill-Sample");
  const mixing = wb.getWorksheet("Mixing")!;
  let found = false;
  mixing.eachRow((r, n) => {
    if (n > 4 && r.getCell(1).value === emp.employeeNumber) {
      const c = r.getCell(6);
      expect(c.value).toBe("C");
      expect((c.fill as ExcelJS.FillPattern).fgColor?.argb).toBe("FF1E63C4");
      expect(c.note).toBeTruthy();
      found = true;
    }
  });
  expect(found).toBe(true);
  const pdf = await page.request.get("/api/export/matrix?format=pdf&area=MX");
  expect(pdf.status()).toBe(200);
  expect((await pdf.body()).subarray(0, 5).toString()).toBe("%PDF-");
});

test("Scenario 7: on a tablet the matrix stays readable and an assessment is easy to submit", async ({ browser }) => {
  const ctx = await browser.newContext({ ...devices["iPad (gen 7)"] });
  const page = await ctx.newPage();
  await login(page, "supervisor.b");
  await page.goto("/matrix?area=PK");
  const overflow = await page.evaluate(() => document.documentElement.scrollWidth - window.innerWidth);
  expect(overflow).toBeLessThanOrEqual(1);
  await expect(page.getByRole("table", { name: "Skills matrix" })).toBeVisible();
  const supB = await db.user.findUniqueOrThrow({ where: { email: "supervisor.b@demo.local" } });
  const emp = await db.employee.findFirstOrThrow({ where: { supervisorUserId: supB.id, jobRole: { title: "Packaging Operator" } } });
  await page.goto(`/assess?employee=${emp.id}`);
  await page.getByLabel("2. Production area").selectOption({ label: "Packaging" });
  await page.getByLabel("3. Competency").selectOption({ index: 7 });
  const radio = page.locator("label").filter({ hasText: /^ITIn Training$/ });
  const box = await radio.boundingBox();
  expect(box!.height).toBeGreaterThanOrEqual(44); // touch-friendly
  await radio.click();
  await page.getByLabel("5. Comments (optional)").fill("Retraining after line change");
  await page.getByRole("button", { name: "6. Submit assessment" }).click();
  await expect(page.getByText(/updated to In Training/)).toBeVisible();
  await ctx.close();
});

test("Scenario 8: Competent without a practical evaluation is blocked with an explanation", async ({ page }) => {
  const emp = await empByLast("Trainee-Sample");
  const c = await comp("MX", "Mixer startup");
  const name = `${emp.firstName} ${emp.lastName}`;
  await login(page, "manager");
  await page.goto(`/matrix?area=MX&focus=${emp.id}`);
  await cell(page, name, c.name).click();
  await pickStatus(page, "Competent");
  await page.getByRole("button", { name: "Save assessment" }).click();
  await expect(page.getByText("Not saved")).toBeVisible();
  await expect(page.getByText(/A practical evaluation is required/)).toBeVisible();
  const rec = await db.employeeSkill.findUnique({ where: { employeeId_competencyId: { employeeId: emp.id, competencyId: c.id } } });
  expect(rec?.status).not.toBe("COMPETENT");
});

test("Scenario 9: two users editing the same assessment – the second is warned and nothing is lost", async ({ browser }) => {
  const emp = await empByLast("Trainee-Sample");
  const c = await comp("MX", "Recipe selection");
  const name = `${emp.firstName} ${emp.lastName}`;
  const url = `/matrix?area=MX&focus=${emp.id}`;
  const a = await browser.newPage();
  const b = await browser.newPage();
  await login(a, "supervisor.a");
  await login(b, "manager");
  await a.goto(url); await b.goto(url);
  await cell(a, name, c.name).click();
  await cell(b, name, c.name).click();
  const v0 = (await db.employeeSkill.findUnique({ where: { employeeId_competencyId: { employeeId: emp.id, competencyId: c.id } } }))?.version ?? 0;

  await pickStatus(a, "Trained");
  await a.getByLabel("Training date").fill(new Date().toISOString().slice(0, 10));
  await a.getByLabel("Assessment notes").fill("Completed TWI sessions (supervisor A)");
  await a.getByRole("button", { name: "Save assessment" }).click();
  await expect(a.getByText(/updated to Trained/)).toBeVisible();

  await pickStatus(b, "Needs Training");
  await b.getByRole("button", { name: "Save assessment" }).click();
  await expect(b.getByText("Changed by someone else")).toBeVisible();
  await expect(b.getByText(/after you opened it/)).toBeVisible();

  const rec = await db.employeeSkill.findUniqueOrThrow({ where: { employeeId_competencyId: { employeeId: emp.id, competencyId: c.id } } });
  expect(rec.status).toBe("TRAINED");
  expect(rec.version).toBe(v0 + 1);
  await expectAudit(`set ${name} – ${c.name} → Trained`);
  expect(await db.auditLog.count({ where: { summary: { contains: `${name} – ${c.name} → Needs Training` } } })).toBe(0);
});

test("Scenario 10: a rejected promotion keeps reason, approver and full history", async ({ browser }) => {
  const emp = await empByLast("Promotion-Sample");
  const req = await db.promotionRequest.findFirstOrThrow({ where: { employeeId: emp.id, status: "SUBMITTED" } });
  const mgr = await browser.newPage();
  await login(mgr, "manager");
  await mgr.goto(`/promotions/${req.id}`);
  await mgr.getByLabel("Comments").fill("Covers both stations reliably on Shift C.");
  await mgr.getByRole("button", { name: "Approve" }).click();
  await mgr.getByRole("button", { name: "Yes, confirm" }).click();
  await expect(mgr.getByText(/sent to HR/)).toBeVisible();

  const hr = await browser.newPage();
  await login(hr, "hr");
  await hr.goto(`/promotions/${req.id}`);
  await hr.getByRole("button", { name: "Reject" }).click();
  await hr.getByRole("button", { name: "Yes, confirm" }).click();
  await expect(hr.getByText("A reason is required when rejecting.")).toBeVisible();
  await hr.getByLabel("Comments").fill("Pay policy requires a 90-day performance review on file first.");
  await hr.getByRole("button", { name: "Reject" }).click();
  await hr.getByRole("button", { name: "Yes, confirm" }).click();
  await expect(hr.getByText(/Rejected. The reason and decision history are kept/)).toBeVisible();
  await hr.reload();
  await expect(hr.getByText("“Pay policy requires a 90-day performance review on file first.”")).toBeVisible();
  await expect(hr.getByText(/Rejected by Demo HR Administrator/)).toBeVisible();
  await expect(hr.getByText(/Approved by Demo Production Manager/)).toBeVisible();

  const final = await db.promotionRequest.findUniqueOrThrow({ where: { id: req.id }, include: { approvals: true } });
  expect(final.status).toBe("REJECTED");
  expect(final.decisionDate).not.toBeNull();
  expect(final.approvals.map((x) => x.decision).sort()).toEqual(["APPROVED", "REJECTED"]);
  expect((await db.employee.findUniqueOrThrow({ where: { id: emp.id }, include: { careerLevel: true } })).careerLevel?.rank).toBe(1);
  // History is immutable at the database level
  await expect(db.approvalRecord.deleteMany({ where: { promotionRequestId: req.id } })).rejects.toThrow(/append-only/);
  await expect(db.auditLog.deleteMany({})).rejects.toThrow(/append-only/);
});
