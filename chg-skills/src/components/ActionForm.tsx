"use client";
import { useRouter } from "next/navigation";
import { useState, useTransition } from "react";
import type { ActionResult } from "@/app/actions/skills";
import { Alert } from "@/components/ui";

/** Small wrapper: submits a server action, shows the result, refreshes the page. */
export function ActionForm({ action, children, className, onDone, resetOnSuccess }: {
  action: (f: FormData) => Promise<ActionResult>; children: React.ReactNode; className?: string; onDone?: (r: ActionResult) => void; resetOnSuccess?: boolean;
}) {
  const router = useRouter();
  const [res, setRes] = useState<ActionResult | null>(null);
  const [pending, start] = useTransition();
  return (
    <form
      className={className}
      onSubmit={(e) => {
        e.preventDefault();
        const f = e.currentTarget;
        const sub = (e.nativeEvent as SubmitEvent).submitter as HTMLButtonElement | null;
        const fd = new FormData(f);
        if (sub?.name) fd.set(sub.name, sub.value);
        start(async () => {
          const r = await action(fd);
          setRes(r);
          onDone?.(r);
          if (r.ok) { if (resetOnSuccess) f.reset(); router.refresh(); }
        });
      }}
    >
      <fieldset disabled={pending} className="contents">{children}</fieldset>
      {res && <div className="col-span-full mt-2"><Alert tone={res.ok ? "success" : "error"}>{res.ok ? res.message : res.error}{!res.ok && res.reasons && <ul className="list-disc pl-5">{res.reasons.map((x) => <li key={x}>{x}</li>)}</ul>}</Alert></div>}
    </form>
  );
}
