"use client";
import { useRouter } from "next/navigation";
import { useState, useTransition } from "react";
import { decidePromotion, withdrawPromotion } from "@/app/actions/promotions";
import type { ActionResult } from "@/app/actions/skills";
import { Alert, Button, Card, CardHeader, Field, Input, Textarea } from "@/components/ui";

export function DecisionPanel({ id, version, stepLabel, canDecide, isFinal, canWithdraw }: { id: string; version: number; stepLabel: string; canDecide: boolean; isFinal: boolean; canWithdraw: boolean }) {
  const router = useRouter();
  const [res, setRes] = useState<ActionResult | null>(null);
  const [confirm, setConfirm] = useState<"APPROVED" | "REJECTED" | "WITHDRAW" | null>(null);
  const [pending, start] = useTransition();
  return (
    <Card>
      <CardHeader title={canDecide ? `Your decision: ${stepLabel}` : "Your request"} />
      <form
        className="grid gap-4 p-5"
        onSubmit={(e) => {
          e.preventDefault();
          if (!confirm) return;
          const f = new FormData(e.currentTarget);
          f.set("decision", confirm === "WITHDRAW" ? "" : confirm);
          start(async () => {
            const r = confirm === "WITHDRAW" ? await withdrawPromotion(f) : await decidePromotion(f);
            setRes(r);
            setConfirm(null);
            router.refresh();
          });
        }}
      >
        <input type="hidden" name="id" value={id} />
        <input type="hidden" name="expectedVersion" value={version} />
        {res && <Alert tone={res.ok ? "success" : res.conflict ? "warn" : "error"}>{res.ok ? res.message : res.error}</Alert>}
        <Field label="Comments" htmlFor="d-comments" hint="Required to reject or withdraw. Kept permanently in the decision history.">
          <Textarea id="d-comments" name="comments" />
        </Field>
        {canDecide && isFinal && (
          <Field label="Effective date" htmlFor="d-eff" hint="Required for approval."><Input id="d-eff" type="date" name="effectiveDate" /></Field>
        )}
        {confirm ? (
          <div className="rounded-md border-2 border-navy-200 p-3">
            <p className="text-sm font-semibold">
              {confirm === "APPROVED" ? (isFinal ? "Approve this promotion and record the new career level?" : "Approve this step and send it on?") : confirm === "REJECTED" ? "Reject this recommendation? The reason will be kept permanently." : "Withdraw this recommendation?"}
            </p>
            <div className="mt-3 flex gap-2">
              <Button type="submit" variant={confirm === "APPROVED" ? "success" : "danger"} disabled={pending}>{pending ? "Saving…" : "Yes, confirm"}</Button>
              <Button type="button" variant="secondary" onClick={() => setConfirm(null)}>Go back</Button>
            </div>
          </div>
        ) : (
          <div className="flex flex-wrap gap-2">
            {canDecide && <Button type="button" variant="success" onClick={() => setConfirm("APPROVED")}>Approve</Button>}
            {canDecide && <Button type="button" variant="danger" onClick={() => setConfirm("REJECTED")}>Reject</Button>}
            {canWithdraw && <Button type="button" variant="secondary" onClick={() => setConfirm("WITHDRAW")}>Withdraw</Button>}
          </div>
        )}
      </form>
    </Card>
  );
}
