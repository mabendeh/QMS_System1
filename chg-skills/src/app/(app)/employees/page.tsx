import Link from "next/link";
import { pageUser } from "@/lib/auth";
import { loadPlant } from "@/lib/data/plant";
import { Badge, Card, PageHeader, SampleBadge, Table, Td, Th } from "@/components/ui";
import { EmployeeSearch } from "./EmployeeSearch";

export const metadata = { title: "Employees" };

export default async function EmployeesPage({ searchParams }: { searchParams: Promise<{ q?: string; shift?: string }> }) {
  const u = await pageUser("employee:view_all");
  const { q = "", shift = "" } = await searchParams;
  const plant = await loadPlant();
  const needle = q.trim().toLowerCase();
  const list = plant.employees.filter((e) =>
    (!needle || `${e.firstName} ${e.lastName} ${e.employeeNumber}`.toLowerCase().includes(needle)) && (!shift || e.shift.code === shift));
  return (
    <div>
      <PageHeader
        title="Employees"
        subtitle="Open a profile to see qualifications, training needs and career progression."
        actions={u.permissions.includes("employee:import") ? <Link href="/admin/import" className="inline-flex h-10 items-center rounded-md border border-line bg-white px-4 text-sm font-semibold hover:bg-canvas">Import employees</Link> : null}
      />
      <EmployeeSearch q={q} shift={shift} />
      <Card className="mt-4">
        <Table>
          <thead>
            <tr><Th>Employee</Th><Th>ID</Th><Th>Shift</Th><Th>Department</Th><Th>Job title</Th><Th>Level</Th><Th>Job skills</Th><Th>Areas qualified</Th></tr>
          </thead>
          <tbody>
            {list.map((e) => {
              const s = plant.employeeSummary(e.id);
              return (
                <tr key={e.id} className="hover:bg-canvas">
                  <Td><Link href={`/employees/${e.id}`} className="font-semibold text-navy-700 hover:underline">{e.firstName} {e.lastName}</Link> {e.isSample && <SampleBadge />}</Td>
                  <Td className="font-mono text-xs">{e.employeeNumber}</Td>
                  <Td>{e.shift.code}</Td>
                  <Td>{e.primaryArea?.name ?? "—"}</Td>
                  <Td>{e.jobRole?.title ?? "—"}</Td>
                  <Td>{e.careerLevel ? `L${e.careerLevel.rank} ${e.careerLevel.name}` : "—"}</Td>
                  <Td>
                    {s.fullyQualified ? <Badge tone="blue">All {s.requiredCount} Competent</Badge> : <Badge tone={s.needsTraining ? "red" : "yellow"}>{s.competentRequired}/{s.requiredCount} Competent</Badge>}
                  </Td>
                  <Td>{s.productionAreasQualified}{s.crossTrained && <Badge tone="green" className="ml-2">Cross-trained</Badge>}</Td>
                </tr>
              );
            })}
            {list.length === 0 && <tr><Td colSpan={8} className="py-10 text-center text-muted">No employees match.</Td></tr>}
          </tbody>
        </Table>
      </Card>
    </div>
  );
}
