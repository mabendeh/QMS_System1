import { NextResponse, type NextRequest } from "next/server";

// Cheap gate: no session cookie → login. Real validation happens server-side on every request.
export function middleware(req: NextRequest) {
  const has = req.cookies.has("chg_session");
  const { pathname, search } = req.nextUrl;
  if (!has && !pathname.startsWith("/login")) {
    if (pathname.startsWith("/api/")) return NextResponse.json({ error: "Sign in required" }, { status: 401 });
    const url = req.nextUrl.clone();
    url.pathname = "/login";
    url.search = `?next=${encodeURIComponent(pathname + search)}`;
    return NextResponse.redirect(url);
  }
  return NextResponse.next();
}

export const config = { matcher: ["/((?!_next/static|_next/image|favicon.ico).*)"] };
