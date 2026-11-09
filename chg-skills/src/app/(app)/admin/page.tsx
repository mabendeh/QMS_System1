import Link from "next/link";
import { Boxes, Briefcase, FileUp, History, Trash2, TrendingUp, Users } from "lucide-react";
import { pageUser } from "@/lib/auth";
import type { Permission } from "@/lib/rbac";
import { PageHeader } from "@/components/ui";

export const metadata = { title: "Administration" };

const ITEMS: { href: string; title: string; text: string; perm: Permission; icon: React.ElementType }[] = [
  { href: "/admin/areas", title: "Departments & competencies", text: "Add, rename, disable or reorder departments and edit each competency checklist.", perm: "config:manage", icon: Boxes },
  { href: "/admin/jobs", title: "Job titles", text: "Which departments each job title must be qualified in.", perm: "config:manage", icon: Briefcase },
  { href: "/admin/rules", title: "Career levels & promotion rules", text: "Level titles and version-controlled promotion requirements.", perm: "rules:manage", icon: TrendingUp },
  { href: "/admin/import", title: "Import employees", text: "Upload a CSV or Excel employee list with preview and validation.", perm: "employee:import", icon: FileUp },
  { href: "/admin/users", title: "Users & permissions", text: "Accounts, roles and trainer authorizations.", perm: "users:manage", icon: Users },
  { href: "/admin/audit", title: "Audit log", text: "Every change, who made it and when.", perm: "audit:view", icon: History },
  { href: "/admin/sample", title: "Sample data", text: "Remove all fictional demonstration records before going live.", perm: "sample:delete", icon: Trash2 },
];

export default async function AdminHome() {
  const u = await pageUser(["config:manage", "users:manage", "rules:manage", "employee:import", "audit:view", "sample:delete"]);
  return (
    <div>
      <PageHeader title="Administration" subtitle="Configuration changes take effect immediately and are recorded in the audit log." />
      <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-3">
        {ITEMS.filter((i) => u.permissions.includes(i.perm)).map((i) => (
          <Link key={i.href} href={i.href} className="flex gap-4 rounded-lg border border-line bg-white p-5 shadow-sm hover:border-navy-600">
            <i.icon className="h-6 w-6 shrink-0 text-navy-700" aria-hidden />
            <span><span className="block font-bold">{i.title}</span><span className="mt-1 block text-sm text-muted">{i.text}</span></span>
          </Link>
        ))}
      </div>
    </div>
  );
}
