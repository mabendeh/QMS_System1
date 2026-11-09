import { pageUser } from "@/lib/auth";
import { PageHeader } from "@/components/ui";
import { ImportClient } from "./ImportClient";

export const metadata = { title: "Import employees" };

export default async function ImportPage() {
  await pageUser("employee:import");
  return (
    <div className="mx-auto max-w-5xl">
      <PageHeader title="Import employees" subtitle="Upload CSV or XLSX. You will see a preview with every problem listed before anything is saved." />
      <ImportClient />
    </div>
  );
}
