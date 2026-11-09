import { pageUser } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { ROLE_LABELS } from "@/lib/rbac";
import { saveUser } from "@/app/actions/admin";
import { ActionForm } from "@/components/ActionForm";
import { Badge, Button, Card, CardHeader, Field, Input, PageHeader, Select } from "@/components/ui";
import type { Role } from "@prisma/client";
import { fmtDateTime } from "@/lib/utils";

export const metadata = { title: "Users & permissions" };

export default async function UsersAdmin() {
  await pageUser("users:manage");
  const [users, areas, employees, sessions] = await Promise.all([
    prisma.user.findMany({ orderBy: { name: "asc" }, include: { roles: true, trainerAuths: true } }),
    prisma.productionArea.findMany({ where: { isActive: true }, orderBy: { sortOrder: "asc" } }),
    prisma.employee.findMany({ where: { status: "ACTIVE" }, orderBy: { lastName: "asc" }, select: { id: true, firstName: true, lastName: true, employeeNumber: true } }),
    prisma.session.groupBy({ by: ["userId"], _max: { lastSeenAt: true } }),
  ]);
  const lastSeen = new Map(sessions.map((s) => [s.userId, s._max.lastSeenAt]));
  const roles = Object.keys(ROLE_LABELS) as Role[];
  const form = (u?: (typeof users)[number]) => {
    const has = new Set(u?.roles.map((r) => r.role));
    const ta = new Set(u?.trainerAuths.map((t) => t.areaId));
    return (
      <ActionForm action={saveUser} className="grid gap-4 p-5 md:grid-cols-2" resetOnSuccess={!u}>
        <input type="hidden" name="id" value={u?.id ?? ""} />
        <Field label="Name" htmlFor={`n-${u?.id ?? "new"}`}><Input id={`n-${u?.id ?? "new"}`} name="name" defaultValue={u?.name} required /></Field>
        <Field label="Work email" htmlFor={`e-${u?.id ?? "new"}`}><Input id={`e-${u?.id ?? "new"}`} name="email" type="email" defaultValue={u?.email} required /></Field>
        <Field label={u ? "Reset password (optional)" : "Initial password"} htmlFor={`p-${u?.id ?? "new"}`} hint="At least 12 characters. Changing it signs the user out everywhere.">
          <Input id={`p-${u?.id ?? "new"}`} name="password" type="password" autoComplete="new-password" minLength={12} required={!u} />
        </Field>
        <Field label="Linked employee record" htmlFor={`l-${u?.id ?? "new"}`} hint="Required for the Employee role (self-service).">
          <Select id={`l-${u?.id ?? "new"}`} name="employeeId" defaultValue={u?.employeeId ?? ""}>
            <option value="">None</option>
            {employees.map((e) => <option key={e.id} value={e.id}>{e.lastName}, {e.firstName} ({e.employeeNumber})</option>)}
          </Select>
        </Field>
        <fieldset className="md:col-span-2">
          <legend className="mb-2 text-sm font-semibold">Roles</legend>
          <div className="grid gap-1 sm:grid-cols-2 lg:grid-cols-3">
            {roles.map((r) => <label key={r} className="flex items-center gap-2 text-sm"><input type="checkbox" name="roles" value={r} defaultChecked={has.has(r)} className="h-4 w-4" /> {ROLE_LABELS[r]}</label>)}
          </div>
        </fieldset>
        <fieldset className="md:col-span-2">
          <legend className="mb-2 text-sm font-semibold">Trainer authorized for (Trainer role only)</legend>
          <div className="grid gap-1 sm:grid-cols-2 lg:grid-cols-4">
            {areas.map((a) => <label key={a.id} className="flex items-center gap-2 text-sm"><input type="checkbox" name="trainerAreas" value={a.id} defaultChecked={ta.has(a.id)} className="h-4 w-4" /> {a.name}</label>)}
          </div>
        </fieldset>
        <label className="flex items-center gap-2 text-sm font-semibold"><input type="checkbox" name="isActive" defaultChecked={u?.isActive ?? true} className="h-5 w-5" /> Active</label>
        <div className="md:col-span-2"><Button type="submit">{u ? "Save user" : "Create user"}</Button></div>
      </ActionForm>
    );
  };
  return (
    <div className="grid gap-5">
      <PageHeader title="Users & permissions" subtitle="Least privilege: give each person only the roles they need. System administrators do not see pay data unless also given an HR role." />
      {users.map((u) => (
        <Card key={u.id}>
          <details>
            <summary className="flex cursor-pointer flex-wrap items-center gap-2 px-5 py-4">
              <span className="font-bold">{u.name}</span><span className="text-sm text-muted">{u.email}</span>
              {u.roles.map((r) => <Badge key={r.role} tone="navy">{ROLE_LABELS[r.role]}</Badge>)}
              {!u.isActive && <Badge tone="red">Disabled</Badge>}
              <span className="ml-auto text-xs text-muted">Last active {fmtDateTime(lastSeen.get(u.id) ?? null)}</span>
            </summary>
            <div className="border-t border-line">{form(u)}</div>
          </details>
        </Card>
      ))}
      <Card><CardHeader title="New user" />{form()}</Card>
    </div>
  );
}
