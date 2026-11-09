import { existsSync, readFileSync, unlinkSync } from "node:fs";
import { execFileSync } from "node:child_process";
import path from "node:path";

/** Drops only the database created by this test run (name recorded in var/e2e-db-url). */
export default async function teardown() {
  const file = path.join(__dirname, "../../var/e2e-db-url");
  if (!existsSync(file)) return;
  const url = readFileSync(file, "utf8").trim();
  const name = url.split("/").pop()!.split("?")[0];
  if (!/^chg_e2e_\d+$/.test(name)) return;
  const admin = url.replace(/\/[^/]+$/, "/postgres");
  try {
    execFileSync("psql", [admin, "-qc", `DROP DATABASE IF EXISTS ${name} WITH (FORCE)`], { stdio: "ignore" });
  } catch { /* leave it for manual cleanup */ }
  unlinkSync(file);
}
