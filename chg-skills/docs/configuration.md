# Configuration

## Environment variables (`.env`)

| Variable | Required | Description |
|---|---|---|
| `DATABASE_URL` | Yes | PostgreSQL connection string, e.g. `postgresql://chg:PASSWORD@db-host:5432/chg_skills?schema=public` |
| `SESSION_SECRET` | Recommended | 64 random hex characters (`openssl rand -hex 32`). Not used by the current version (sessions are random tokens stored hashed in the database); reserved for future signed links or SSO state. Keep it secret |
| `UPLOAD_DIR` | Yes | Folder for evidence files, outside the web root. Back it up with the database |
| `COOKIE_SECURE` | Production | `true` when served over HTTPS |
| `SEED_DEMO_PASSWORD` | Demo only | Password for the fictional demo accounts created by `npm run db:seed` |
| `SEED_SKIP_DEMO` | Optional | `true` loads configuration only (no sample people) |
| `TEST_DATABASE_URL` | Tests only | Points at the test server; each e2e run creates and drops its own `chg_e2e_*` database there |

Never commit `.env`. It is in `.gitignore`.

## What is configurable in the app, without code changes

| Item | Where | Who |
|---|---|---|
| Departments: add, rename, disable, reorder, plant-wide vs production | Administration → Departments & competencies | HR, System Admin |
| Competency checklists, critical flag, order | same | HR, System Admin |
| Requalification intervals (per department or competency) | same | HR, System Admin |
| Job titles and the departments they require | Administration → Job titles | HR, System Admin |
| Career level titles and descriptions | Administration → Career levels & promotion rules | HR |
| Promotion rules (version-controlled) | same | HR |
| Pay grades and skill premiums | HR Compensation | HR |
| Shift staffing targets | Shift Coverage → Edit staffing targets | Production Manager |
| Users, roles, trainer authorizations | Administration → Users & permissions | System Admin |

Shift codes are A, B and C. Changing them requires a database edit of the `shifts` table.

## Single sign-on (optional)

The project does not assume any identity provider. To connect one:

1. Register the app with the provider (OIDC is recommended) and set the redirect URL to `https://<host>/api/auth/callback`.
2. Add a route handler that completes the OIDC code flow, using a maintained library such as `openid-client`.
3. On success, look up or create the `User` by email, map provider groups to `Role` values, then create the session the same way `login()` does in `src/lib/auth.ts`.
4. Set `passwordHash` to `null` for SSO users so password sign-in is disabled for them.

Every role check stays the same, because it reads roles from the `user_roles` table.

## Scheduling and attendance (optional)

The coverage page reports **total qualified headcount** by home shift. "Scheduled" and "Available" are shown as **Unknown** until a scheduling or attendance feed is integrated. The values are deliberately never estimated. To add one, populate scheduled and available counts in `computeCoverage()` (`src/lib/domain/coverage.ts`) from the integration.
