import Link from "next/link";

export const metadata = { title: "Access denied" };

export default function Denied() {
  return (
    <main className="flex min-h-screen items-center justify-center p-6">
      <div className="max-w-md rounded-lg border border-line bg-white p-8 text-center shadow-sm">
        <h1 className="text-xl font-bold">Access denied</h1>
        <p className="mt-2 text-sm text-muted">Your role does not include access to this page. If you need it, ask HR or the system administrator.</p>
        <Link href="/" className="mt-5 inline-block font-semibold text-navy-700 underline">Back to home</Link>
      </div>
    </main>
  );
}
