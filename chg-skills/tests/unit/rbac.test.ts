import { describe, expect, it } from "vitest";
import { can, ROLE_PERMISSIONS } from "@/lib/rbac";

describe("role-based access", () => {
  it("Scenario 4: supervisors cannot see compensation", () => {
    expect(can(["SUPERVISOR"], "comp:view")).toBe(false);
    expect(can(["SUPERVISOR"], "comp:manage")).toBe(false);
  });
  it("system administrators do not automatically see pay data", () => {
    expect(can(["SYSTEM_ADMIN"], "comp:view")).toBe(false);
    expect(can(["SYSTEM_ADMIN"], "users:manage")).toBe(true);
  });
  it("only HR manages pay; leadership can view it", () => {
    const managers = Object.entries(ROLE_PERMISSIONS).filter(([, p]) => p.includes("comp:manage")).map(([r]) => r);
    expect(managers).toEqual(["HR_ADMIN"]);
    expect(can(["LEADERSHIP"], "comp:view")).toBe(true);
  });
  it("employees only get self-service", () => {
    expect(ROLE_PERMISSIONS.EMPLOYEE).toEqual(["self:view"]);
  });
  it("supervisors cannot verify their own Competent requests or approve promotions", () => {
    expect(can(["SUPERVISOR"], "qualification:verify")).toBe(false);
    expect(can(["SUPERVISOR"], "promotion:final_approve")).toBe(false);
  });
});
