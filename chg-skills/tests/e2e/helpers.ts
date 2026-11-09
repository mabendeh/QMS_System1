import { readFileSync } from "node:fs";
import path from "node:path";
import { PrismaClient } from "@prisma/client";
import { expect, type Page } from "@playwright/test";

const env = Object.fromEntries(
  readFileSync(path.join(__dirname, "../../.env"), "utf8").split("\n").filter((l) => l.includes("=")).map((l) => {
    const i = l.indexOf("=");
    return [l.slice(0, i).trim(), l.slice(i + 1).trim().replace(/^"|"$/g, "")];
  }),
);
export const PASSWORD = env.SEED_DEMO_PASSWORD;
/** Direct access to this run's test database (created by start-test-server.sh) for arranging and checking data. */
const runDb = readFileSync(path.join(__dirname, "../../var/e2e-db-url"), "utf8").trim();
export const db = new PrismaClient({ datasources: { db: { url: runDb } } });

export async function login(page: Page, who: string) {
  await page.context().clearCookies();
  await page.goto("/login");
  await page.fill("#email", `${who}@demo.local`);
  await page.fill("#password", PASSWORD);
  await page.click("button[type=submit]");
  await page.waitForURL((u) => !u.pathname.startsWith("/login"));
}

export async function empByLast(lastName: string) {
  return db.employee.findFirstOrThrow({ where: { lastName } });
}
export async function comp(areaCode: string, name: string) {
  return db.competency.findFirstOrThrow({ where: { name, area: { code: areaCode } }, include: { area: true } });
}
export function cell(page: Page, empName: string, compName: string) {
  return page.getByRole("button", { name: new RegExp(`^${escape(empName)} – ${escape(compName)}: .*Open assessment$`) });
}
const escape = (s: string) => s.replace(/[.*+?^${}()|[\]\\/]/g, "\\$&");
export async function chipStatus(page: Page, empName: string, compName: string) {
  return cell(page, empName, compName).locator("[data-status]").getAttribute("data-status");
}
export async function expectAudit(match: string) {
  await expect.poll(async () => db.auditLog.count({ where: { summary: { contains: match } } })).toBeGreaterThan(0);
}

/** Picks a skill level inside the open assessment dialog. */
export async function pickStatus(page: Page, label: "Needs Training" | "In Training" | "Trained" | "Competent" | "Not Applicable") {
  await page.getByRole("dialog").getByRole("radio", { name: label, exact: true }).check({ force: true });
}
