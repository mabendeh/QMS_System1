import { pageUser } from "@/lib/auth";
import { ROLE_LABELS } from "@/lib/rbac";
import { AppShell } from "@/components/layout/AppShell";

export const dynamic = "force-dynamic";

export default async function AppLayout({ children }: { children: React.ReactNode }) {
  const u = await pageUser();
  return (
    <AppShell user={{ name: u.name, roleLabel: u.roles.map((r) => ROLE_LABELS[r]).join(", "), permissions: u.permissions, employeeId: u.employeeId }}>
      {children}
    </AppShell>
  );
}
