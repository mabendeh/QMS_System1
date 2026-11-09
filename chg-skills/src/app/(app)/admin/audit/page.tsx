import Link from "next/link";
import { pageUser } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { Card, Input, PageHeader, Table, Td, Th } from "@/components/ui";
import { fmtDateTime } from "@/lib/utils";

export const metadata = { title: "Audit log" };

export default async function AuditPage({ searchParams }: { searchParams: Promise<{ q?: string; page?: string }> }) {
  const u = await pageUser("audit:view");
  const { q = "", page = "1" } = await searchParams;
  const p = Math.max(1, Number(page) || 1);
  const canSeeComp = u.permissions.includes("comp:view");
  const where = {
    ...(q ? { OR: [{ summary: { contains: q, mode: "insensitive" as const } }, { action: { contains: q, mode: "insensitive" as const } }] } : {}),
    ...(canSeeComp ? {} : { NOT: { action: { startsWith: "comp." } } }),
  };
  const [rows, total] = await Promise.all([
    prisma.auditLog.findMany({ where, orderBy: { at: "desc" }, take: 100, skip: (p - 1) * 100 }),
    prisma.auditLog.count({ where }),
  ]);
  return (
    <div>
      <PageHeader title="Audit log" subtitle={`Append-only record of every change (${total.toLocaleString()} entries).${canSeeComp ? "" : " Compensation changes are hidden for your role."}`} />
      <form className="mb-4 max-w-md"><Input name="q" defaultValue={q} placeholder="Search actions or names" aria-label="Search audit log" /></form>
      <Card>
        <Table>
          <thead><tr><Th>When</Th><Th>Who</Th><Th>Action</Th><Th>What happened</Th></tr></thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r.id.toString()}>
                <Td className="whitespace-nowrap text-muted">{fmtDateTime(r.at)}</Td>
                <Td>{r.actorName ?? "system"}</Td>
                <Td className="font-mono text-xs">{r.action}</Td>
                <Td>{r.summary}</Td>
              </tr>
            ))}
          </tbody>
        </Table>
        <div className="flex justify-between p-4 text-sm">
          {p > 1 ? <Link className="font-semibold text-navy-700 underline" href={`/admin/audit?q=${encodeURIComponent(q)}&page=${p - 1}`}>Newer</Link> : <span />}
          {p * 100 < total && <Link className="font-semibold text-navy-700 underline" href={`/admin/audit?q=${encodeURIComponent(q)}&page=${p + 1}`}>Older</Link>}
        </div>
      </Card>
    </div>
  );
}
