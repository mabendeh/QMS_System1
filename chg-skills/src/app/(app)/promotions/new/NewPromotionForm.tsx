"use client";
import { useRouter } from "next/navigation";
import { useState, useTransition } from "react";
import { submitPromotion } from "@/app/actions/promotions";
import type { ActionResult } from "@/app/actions/skills";
import { Alert, Button, Card, Field, Select, Textarea } from "@/components/ui";

export function NewPromotionForm({ employeeId, levels, targetId, ready }: { employeeId: string; levels: { id: string; name: string }[]; targetId: string; ready: boolean }) {
  const router = useRouter();
  const [res, setRes] = useState<ActionResult | null>(null);
  const [pending, start] = useTransition();
  return (
    <Card className="p-5">
      <form
        className="grid gap-4"
        onSubmit={(e) => {
          e.preventDefault();
          const f = e.currentTarget;
          start(async () => {
            const r = await submitPromotion(new FormData(f));
            setRes(r);
            if (r.ok && "id" in r && r.id) router.push(`/promotions/${r.id}`);
          });
        }}
      >
        <input type="hidden" name="employeeId" value={employeeId} />
        <Field label="Recommended level" htmlFor="p-level">
          <Select id="p-level" name="toLevelId" defaultValue={targetId} onChange={(e) => router.replace(`/promotions/new?employee=${employeeId}&level=${e.target.value}`)}>
            {levels.map((l) => <option key={l.id} value={l.id}>{l.name}</option>)}
          </Select>
        </Field>
        <Field label="Why are you recommending this employee?" htmlFor="p-just" hint="Describe what they have demonstrated. Do not include pay amounts.">
          <Textarea id="p-just" name="justification" required minLength={20} rows={5} />
        </Field>
        {!ready && <Alert tone="warn">Not all documented requirements are met yet. You can still try – the system will list what is missing.</Alert>}
        {res && !res.ok && (
          <Alert tone="error" title={res.error}>{res.reasons && <ul className="mt-1 list-disc pl-5">{res.reasons.map((r) => <li key={r}>{r}</li>)}</ul>}</Alert>
        )}
        <Button type="submit" size="lg" disabled={pending}>{pending ? "Submitting…" : "Submit recommendation"}</Button>
        <p className="text-xs text-muted">Submitting starts the review workflow. It does not change the employee&apos;s level or pay.</p>
      </form>
    </Card>
  );
}
