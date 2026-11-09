import Link from "next/link";
import { notFound, redirect } from "next/navigation";
import { CheckCircle2, Circle } from "lucide-react";
import { pageUser } from "@/lib/auth";
import { loadPlant } from "@/lib/data/plant";
import { Alert, Card, CardHeader, PageHeader, SampleBadge } from "@/components/ui";
import { NewPromotionForm } from "./NewPromotionForm";

export const metadata = { title: "Recommend for promotion" };

export default async function NewPromotion({ searchParams }: { searchParams: Promise<{ employee?: string; level?: string }> }) {
  const u = await pageUser("promotion:submit");
  const { employee, level } = await searchParams;
  if (!employee) redirect("/promotions");
  const plant = await loadPlant({ employeeIds: [employee] });
  const e = plant.employees[0];
  if (!e) notFound();
  if (!u.roles.includes("PRODUCTION_MANAGER") && e.supervisorUserId !== u.id) redirect("/denied");
  const higher = plant.levels.filter((l) => l.rank > (e.careerLevel?.rank ?? -1));
  const target = higher.find((l) => l.id === level) ?? higher[0];
  const ev = target ? plant.evaluateFor(e.id, target.id) : null;
  return (
    <div className="mx-auto grid max-w-3xl gap-5">
      <PageHeader title="Recommend for promotion" subtitle={<><Link href={`/employees/${e.id}`} className="font-semibold text-navy-700 hover:underline">{e.firstName} {e.lastName}</Link> {e.isSample && <SampleBadge />} · currently {e.careerLevel ? `L${e.careerLevel.rank} ${e.careerLevel.name}` : "no level"}</>} />
      {!target ? <Alert tone="info">This employee is already at the highest configured level.</Alert> : (
        <>
          <Card>
            <CardHeader title={`Requirements for ${target.name}`} subtitle={ev ? `Promotion rules v${ev.version} · checked live from verified qualifications` : "No rules configured"} />
            {ev ? (
              <ul className="grid gap-3 p-5 text-sm">
                {ev.result.items.map((it) => (
                  <li key={it.id} className="flex gap-2">
                    {it.met ? <CheckCircle2 className="h-5 w-5 shrink-0 text-st-blue" /> : <Circle className={`h-5 w-5 shrink-0 ${it.manual ? "text-muted" : "text-st-red"}`} />}
                    <span><b>{it.label}</b> – <span className="text-muted">{it.detail}</span></span>
                  </li>
                ))}
              </ul>
            ) : <p className="p-5 text-sm text-muted">HR has not configured rules for this level.</p>}
          </Card>
          <NewPromotionForm employeeId={e.id} levels={higher.map((l) => ({ id: l.id, name: `L${l.rank} ${l.name}` }))} targetId={target.id} ready={!!ev?.result.readyForReview} />
        </>
      )}
    </div>
  );
}
