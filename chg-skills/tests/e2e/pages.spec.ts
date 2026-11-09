import { expect, test } from "@playwright/test";
import { cell, chipStatus, db, login, pickStatus } from "./helpers";

test.describe.configure({ mode: "serial" });
test.afterAll(async () => { await db.$disconnect(); });

const PAGES: [string, string[]][] = [
  ["hr", ["/", "/matrix", "/matrix?area=PW", "/employees", "/training", "/training?tab=gaps", "/promotions", "/coverage", "/analytics", "/hr/compensation", "/admin", "/admin/areas", "/admin/jobs", "/admin/rules", "/admin/import", "/admin/audit"]],
  ["manager", ["/", "/assess", "/training?tab=verify", "/coverage?edit=1", "/promotions?view=closed"]],
  ["supervisor.c", ["/", "/matrix?area=PK", "/assess", "/training?tab=assign", "/promotions"]],
  ["trainer", ["/training", "/assess"]],
  ["ci", ["/analytics", "/coverage"]],
  ["quality", ["/training?tab=verify", "/admin/audit"]],
  ["leadership", ["/promotions", "/hr/compensation", "/analytics"]],
  ["sysadmin", ["/admin", "/admin/users", "/admin/areas", "/admin/sample"]],
];

for (const [who, urls] of PAGES) {
  test(`${who} can open every page their role allows`, async ({ page }) => {
    const errors: string[] = [];
    page.on("pageerror", (e) => errors.push(e.message));
    await login(page, who);
    for (const url of urls) {
      const res = await page.goto(url);
      expect(res?.status(), url).toBeLessThan(400);
      await expect(page, url).not.toHaveURL(/\/denied|\/login/);
      await expect(page.locator("h1").first(), url).toBeVisible();
    }
    expect(errors).toEqual([]);
  });
}

test("employee self-service sees only their own record", async ({ page }) => {
  await login(page, "employee");
  await expect(page).toHaveURL(/\/employees\/.+/);
  await expect(page.getByText("My qualifications", { exact: true })).toBeVisible();
  for (const url of ["/matrix", "/employees", "/training", "/admin"]) {
    await page.goto(url);
    await expect(page, url).toHaveURL(/\/denied/);
  }
  const other = await db.employee.findFirstOrThrow({ where: { lastName: "Promotion-Sample" } });
  await page.goto(`/employees/${other.id}`);
  await expect(page).toHaveURL(/\/denied/);
  const res = await page.request.get("/api/export/matrix?format=xlsx");
  expect(res.status()).toBe(403);
});


const PNG = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==", "base64");

test("critical skill: evidence upload is validated, required, stored and downloadable", async ({ browser }) => {
  const emp = await db.employee.findFirstOrThrow({ where: { lastName: "Trainee-Sample" } });
  const c = await db.competency.findFirstOrThrow({ where: { name: "Allergen changeover", area: { code: "MX" } } });
  const name = `${emp.firstName} ${emp.lastName}`;
  const url = `/matrix?area=MX&focus=${emp.id}`;

  const t = await browser.newPage();
  await login(t, "trainer");
  await t.goto(url);
  await cell(t, name, c.name).click();
  await t.getByRole("tab", { name: "Practical evaluation" }).click();
  await t.getByLabel("Passed", { exact: true }).check();
  await t.getByLabel("Operated independently without help").check();
  await t.getByRole("button", { name: "Record evaluation" }).click();
  await expect(t.getByText(/Practical evaluation recorded/)).toBeVisible();

  const m = await browser.newPage();
  await login(m, "manager");
  await m.goto(url);
  await cell(m, name, c.name).click();
  await pickStatus(m, "Competent");
  await m.getByRole("button", { name: "Save assessment" }).click();
  await expect(m.getByText(/critical skill: upload supporting evidence/i)).toBeVisible();
  // A disguised executable is rejected by content check
  await m.setInputFiles("#a-file", { name: "checklist.png", mimeType: "image/png", buffer: Buffer.from("MZ\x90\x00 not really a png") });
  await m.getByRole("button", { name: "Save assessment" }).click();
  await expect(m.getByText("Only PDF, PNG, JPEG or WEBP files are accepted.")).toBeVisible();
  await m.setInputFiles("#a-file", { name: "signed-checklist.png", mimeType: "image/png", buffer: PNG });
  await m.getByRole("button", { name: "Save assessment" }).click();
  await expect(m.getByText(/updated to Competent/)).toBeVisible();
  const ev = await db.evidence.findFirstOrThrow({ where: { employeeId: emp.id, competencyId: c.id } });
  expect(ev.storedName).toMatch(/^[a-f0-9]{32}\.png$/);
  const dl = await m.request.get(`/api/evidence/${ev.id}`);
  expect(dl.status()).toBe(200);
  expect(dl.headers()["content-type"]).toBe("image/png");
  await m.keyboard.press("Escape");
  await expect.poll(() => chipStatus(m, name, c.name)).toBe("COMPETENT");
  const rec = await db.employeeSkill.findUniqueOrThrow({ where: { employeeId_competencyId: { employeeId: emp.id, competencyId: c.id } } });
  expect(rec.expiresAt).not.toBeNull(); // 12-month requalification rule applied
});

test("HR imports employees with preview, validation and duplicate detection", async ({ page }) => {
  await login(page, "hr");
  await page.goto("/admin/import");
  const csv = [
    "Employee ID,First Name,Last Name,Shift,Department,Job Title,Hire Date",
    "T-9001,Test,Importone,A,Mixing,Mixer Operator,2026-01-05",
    "T-9001,Dup,Licate,B,,,",
    "DEMO-1001,Existing,Person,B,,,",
    "T-9002,Bad,Shift,Z,,,",
  ].join("\n");
  await page.setInputFiles("#imp-file", { name: "employees.csv", mimeType: "text/csv", buffer: Buffer.from(csv) });
  await page.getByRole("button", { name: "Preview import" }).click();
  await expect(page.getByText(/Duplicate Employee ID \(also on row 2\)/)).toBeVisible();
  await expect(page.getByText(/Unknown shift “Z”/)).toBeVisible();
  await expect(page.getByText("Exists")).toBeVisible();
  const before = await db.employeeSkill.count();
  await page.getByRole("button", { name: /^Import 1 new/ }).click();
  await expect(page.getByText(/Import complete: 1 created, 0 updated/)).toBeVisible();
  expect(await db.employee.count({ where: { employeeNumber: "T-9001" } })).toBe(1);
  const existing = await db.employee.findUniqueOrThrow({ where: { employeeNumber: "DEMO-1001" } });
  expect(existing.firstName).not.toBe("Existing");
  expect(await db.employeeSkill.count()).toBe(before);
});

test("HR configures a pay grade; amounts are validated and audited", async ({ page }) => {
  await login(page, "hr");
  await page.goto("/hr/compensation");
  await page.getByRole("row").filter({ hasText: "PG-1" }).getByRole("button", { name: "Edit" }).click();
  await page.getByLabel("Minimum hourly rate (USD)").fill("30");
  await page.getByLabel("Maximum hourly rate (USD)").fill("20");
  await page.getByLabel("Effective date").first().fill("2027-01-01");
  await page.getByRole("button", { name: "Save" }).click();
  await expect(page.getByText("Minimum cannot be greater than maximum.")).toBeVisible();
  await page.getByLabel("Maximum hourly rate (USD)").fill("40");
  await page.getByRole("button", { name: "Save" }).click();
  await expect(page.getByText("Pay grade saved.")).toBeVisible();
  expect(await db.auditLog.count({ where: { action: "comp.paygrade_updated" } })).toBe(1);
});

test("HR publishes a new promotion rule version; the old version is kept", async ({ page }) => {
  await login(page, "hr");
  await page.goto("/admin/rules");
  const card = page.locator("div.rounded-lg").filter({ hasText: "Level 2 – Multi-Skilled Operator" }).first();
  await card.getByRole("button", { name: "Edit rules (new version)" }).click();
  await page.getByLabel("Change notes").fill("Pilot review: same rules, new wording");
  await page.getByRole("button", { name: "Publish new version" }).click();
  await expect(page.getByText(/Rules v2 published/)).toBeVisible();
  const lvl = await db.careerLevel.findUniqueOrThrow({ where: { rank: 2 } });
  const sets = await db.promotionRuleSet.findMany({ where: { careerLevelId: lvl.id }, orderBy: { version: "asc" } });
  expect(sets.map((s) => [s.version, s.isActive])).toEqual([[1, false], [2, true]]);
});

test("system administrator deletes all sample data with typed confirmation", async ({ page }) => {
  await login(page, "sysadmin");
  await page.goto("/admin/sample");
  await page.getByLabel("Type DELETE SAMPLE DATA to confirm").fill("delete");
  await page.getByRole("button", { name: "Delete all sample data" }).click();
  await expect(page.getByText("Type DELETE SAMPLE DATA exactly to confirm.")).toBeVisible();
  expect(await db.employee.count({ where: { isSample: true } })).toBeGreaterThan(0);
  await page.getByLabel("Type DELETE SAMPLE DATA to confirm").fill("DELETE SAMPLE DATA");
  await page.getByRole("button", { name: "Delete all sample data" }).click();
  await expect(page.getByText(/Deleted \d+ sample employees/)).toBeVisible();
  expect(await db.employee.count({ where: { isSample: true } })).toBe(0);
  expect(await db.employee.count({ where: { employeeNumber: "T-9001" } })).toBe(1); // real (imported) data untouched
  expect(await db.auditLog.count({ where: { action: "system.sample_deleted" } })).toBe(1);
  expect(await db.productionArea.count()).toBeGreaterThan(10);
});
