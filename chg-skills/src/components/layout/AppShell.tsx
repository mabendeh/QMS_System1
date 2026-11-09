"use client";
import Link from "next/link";
import { usePathname } from "next/navigation";
import { useState } from "react";
import {
  LayoutDashboard, Grid3x3, Users, ClipboardCheck, GraduationCap, TrendingUp, Wallet, ShieldCheck, BarChart3, Settings, LogOut, Menu, X, UserCircle,
} from "lucide-react";
import { cn } from "@/lib/utils";
import { logoutAction } from "@/app/actions/auth";
import type { Permission } from "@/lib/rbac";

const NAV: { href: string; label: string; icon: React.ElementType; any: Permission[] }[] = [
  { href: "/", label: "Home", icon: LayoutDashboard, any: ["matrix:view"] },
  { href: "/matrix", label: "Skills Matrix", icon: Grid3x3, any: ["matrix:view"] },
  { href: "/employees", label: "Employees", icon: Users, any: ["employee:view_all"] },
  { href: "/assess", label: "Quick Assessment", icon: ClipboardCheck, any: ["skill:assess"] },
  { href: "/training", label: "Training", icon: GraduationCap, any: ["training:view"] },
  { href: "/promotions", label: "Promotion Review", icon: TrendingUp, any: ["promotion:view"] },
  { href: "/coverage", label: "Shift Coverage", icon: ShieldCheck, any: ["coverage:view"] },
  { href: "/analytics", label: "CI Analytics", icon: BarChart3, any: ["analytics:view"] },
  { href: "/hr/compensation", label: "HR Compensation", icon: Wallet, any: ["comp:view"] },
  { href: "/admin", label: "Administration", icon: Settings, any: ["config:manage", "users:manage", "rules:manage", "employee:import", "audit:view", "sample:delete"] },
];

export function AppShell({
  user, children,
}: { user: { name: string; roleLabel: string; permissions: Permission[]; employeeId: string | null }; children: React.ReactNode }) {
  const path = usePathname();
  const [open, setOpen] = useState(false);
  const items = NAV.filter((n) => n.any.some((p) => user.permissions.includes(p)));
  if (user.permissions.includes("self:view") && user.employeeId) {
    items.unshift({ href: `/employees/${user.employeeId}`, label: "My Qualifications", icon: UserCircle, any: [] });
  }
  const active = (href: string) => (href === "/" ? path === "/" : path.startsWith(href));
  const nav = (
    <nav aria-label="Main" className="grid gap-0.5 p-3">
      {items.map((n) => (
        <Link
          key={n.href}
          href={n.href}
          onClick={() => setOpen(false)}
          aria-current={active(n.href) ? "page" : undefined}
          className={cn(
            "flex items-center gap-3 rounded-md px-3 py-2.5 text-[15px] font-medium text-navy-100 hover:bg-navy-700 hover:text-white",
            active(n.href) && "bg-white text-navy-800 hover:bg-white hover:text-navy-800",
          )}
        >
          <n.icon className="h-5 w-5 shrink-0" aria-hidden />
          {n.label}
        </Link>
      ))}
    </nav>
  );
  return (
    <div className="min-h-screen lg:grid lg:grid-cols-[240px_1fr]">
      <aside className="no-print hidden bg-navy-800 lg:sticky lg:top-0 lg:flex lg:h-screen lg:flex-col">
        <Brand />
        <div className="flex-1 overflow-y-auto">{nav}</div>
        <UserBox user={user} />
      </aside>
      <header className="no-print sticky top-0 z-30 flex items-center justify-between bg-navy-800 px-4 py-2 lg:hidden">
        <Brand small />
        <button onClick={() => setOpen(!open)} className="rounded p-2 text-white" aria-expanded={open} aria-label={open ? "Close menu" : "Open menu"}>
          {open ? <X /> : <Menu />}
        </button>
      </header>
      {open && (
        <div className="no-print fixed inset-x-0 top-[56px] z-30 max-h-[calc(100vh-56px)] overflow-y-auto bg-navy-800 shadow-xl lg:hidden">
          {nav}
          <UserBox user={user} />
        </div>
      )}
      <main className="min-w-0 px-4 py-5 sm:px-6 lg:px-8 lg:py-7">{children}</main>
    </div>
  );
}

function Brand({ small }: { small?: boolean }) {
  return (
    <Link href="/" className={cn("flex items-center gap-3 text-white", small ? "" : "px-5 py-5")}>
      <span className="flex h-9 w-9 items-center justify-center rounded border-2 border-white/80 text-xs font-bold">CHG</span>
      <span className="leading-tight">
        <span className="block text-sm font-bold">Skills &amp; Careers</span>
        <span className="block text-xs text-navy-200">Duncanville · Pilot</span>
      </span>
    </Link>
  );
}

function UserBox({ user }: { user: { name: string; roleLabel: string } }) {
  return (
    <div className="border-t border-navy-700 p-4 text-sm text-navy-100">
      <p className="font-semibold text-white">{user.name}</p>
      <p className="text-xs">{user.roleLabel}</p>
      <form action={logoutAction} className="mt-3">
        <button className="flex items-center gap-2 text-xs font-semibold text-navy-100 hover:text-white">
          <LogOut className="h-4 w-4" aria-hidden /> Sign out
        </button>
      </form>
    </div>
  );
}
