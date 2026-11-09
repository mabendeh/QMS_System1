import { redirect } from "next/navigation";
import { getCurrentUser } from "@/lib/auth";
import { LoginForm } from "./LoginForm";

export const metadata = { title: "Sign in" };

export default async function LoginPage({ searchParams }: { searchParams: Promise<{ next?: string }> }) {
  if (await getCurrentUser()) redirect("/");
  const { next } = await searchParams;
  return (
    <main className="flex min-h-screen items-center justify-center bg-navy-800 px-4 py-10">
      <div className="w-full max-w-md">
        <div className="mb-6 text-center text-white">
          <div className="mx-auto mb-3 flex h-12 w-12 items-center justify-center rounded-lg border-2 border-white/80 text-lg font-bold">CHG</div>
          <h1 className="text-2xl font-bold">Production Skills &amp; Career Progression</h1>
          <p className="mt-1 text-sm text-navy-100">Duncanville plant · private pilot</p>
        </div>
        <div className="rounded-lg bg-white p-6 shadow-xl">
          <LoginForm next={next ?? "/"} />
        </div>
        <p className="mt-4 text-center text-xs text-navy-100">Authorized users only. Activity is logged.</p>
      </div>
    </main>
  );
}
