import { pageUser } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { deleteSampleData } from "@/app/actions/admin";
import { ActionForm } from "@/components/ActionForm";
import { Alert, Button, Card, Field, Input, PageHeader } from "@/components/ui";

export const metadata = { title: "Sample data" };

export default async function SamplePage() {
  await pageUser("sample:delete");
  const [emps, promos] = await Promise.all([
    prisma.employee.count({ where: { isSample: true } }),
    prisma.promotionRequest.count({ where: { employee: { isSample: true } } }),
  ]);
  return (
    <div className="mx-auto max-w-2xl">
      <PageHeader title="Sample data" subtitle="All demonstration employees are fictional, have IDs starting with DEMO- and show a Sample badge." />
      <Card className="grid gap-4 p-5">
        <p><b>{emps}</b> sample employees and <b>{promos}</b> sample promotion requests are in the system.</p>
        <Alert tone="warn" title="This cannot be undone">
          Deleting removes every sample employee with their sample qualifications, assessments, training, evidence records and promotion requests, and disables demo logins linked to them.
          Configuration (departments, competencies, rules) and real employees are not touched. The audit log keeps a record that the deletion happened.
        </Alert>
        <ActionForm action={deleteSampleData} className="grid gap-3">
          <Field label="Type DELETE SAMPLE DATA to confirm" htmlFor="confirm"><Input id="confirm" name="confirm" autoComplete="off" /></Field>
          <div><Button type="submit" variant="danger" disabled={emps === 0}>Delete all sample data</Button></div>
        </ActionForm>
      </Card>
    </div>
  );
}
