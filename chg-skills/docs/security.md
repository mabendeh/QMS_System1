# Security and access control

## Principles

- **Least privilege.** Each role gets only the permissions it needs. Permissions are checked **on the server** in every page, server action and API route (`src/lib/rbac.ts`, `src/lib/auth.ts`, `src/lib/scope.ts`). Hiding a menu item is never the security boundary.
- **Pay data is isolated.** Only HR Administrators (edit) and Plant Leadership (view) hold `comp:view`. System Administrators deliberately do not.
- **Evidence-based qualification.** Competent requires a passed practical evaluation that confirms independent operation, evidence for critical skills, and verification by an authorized evaluator who is not the person who ran the practical.
- **Separation of duties in promotions.** The submitter cannot review their own request, one person cannot decide two steps, and only the authorized role can decide each step.
- **Nothing automatic about pay.** The system records recommendations and decisions. It never changes pay.

## Permission table

| Permission | HR | Prod. Mgr | Supervisor | Trainer | CI | Quality | Leadership | Employee | Sys Admin |
|---|---|---|---|---|---|---|---|---|---|
| View matrix / all employees / training / coverage, export | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | – | ✓ |
| Assess skills | – | ✓ all | ✓ own team | ✓ authorized areas | – | – | – | – | – |
| Verify Competent | – | ✓ | – | – | – | ✓ | – | – | – |
| Record practical evaluation | – | ✓ | – | ✓ authorized areas | – | – | – | – | – |
| Assign training | ✓ | ✓ | ✓ own team | ✓ | – | – | – | – | – |
| Submit promotion | – | ✓ | ✓ own team | – | – | – | – | – | – |
| Promotion: manager review / HR verification / final approval | HR verify | Mgr review | – | – | – | – | Final | – | – |
| Career levels & promotion rules | ✓ | – | – | – | – | – | – | – | – |
| View pay | ✓ | – | – | – | – | – | ✓ | – | – |
| Edit pay | ✓ | – | – | – | – | – | – | – | – |
| Departments, competencies, job titles | ✓ | – | – | – | – | – | – | – | ✓ |
| Users & roles | – | – | – | – | – | – | – | – | ✓ |
| Employee import | ✓ | – | – | – | – | – | – | – | – |
| Coverage targets | – | ✓ | – | – | – | – | – | – | – |
| Analytics | ✓ | ✓ | – | – | ✓ | ✓ | ✓ | – | – |
| Audit log | ✓ | – | – | – | – | ✓ | – | – | ✓ |
| Delete sample data | – | – | – | – | – | – | – | – | ✓ |
| Own record only | – | – | – | – | – | – | – | ✓ | – |

Automated tests cover these rules, including Scenario 4 (a supervisor is denied compensation) and employee self-service isolation.

## Authentication and sessions

- Email and password, hashed with **bcrypt** (cost 12). There are **no hardcoded passwords**: demo passwords come from `SEED_DEMO_PASSWORD` or are generated, and real admins are created with `npm run admin:create`.
- Sessions use a random 256-bit token in an `HttpOnly`, `SameSite=Lax` cookie (`Secure` when `COOKIE_SECURE=true`). Only a SHA-256 hash of the token is stored. Sessions expire after 12 hours. Disabling a user or changing their password ends their sessions.
- After 5 failed sign-ins, that email is locked out for 15 minutes. The counter is per app instance, so add a reverse-proxy rate limit in production.
- **SSO:** the identity layer is isolated in `src/lib/auth.ts`. To use the company identity provider (OIDC/SAML, for example Entra ID or Okta), replace `login()` with the provider callback and map groups to roles. No provider is assumed. See [configuration.md](configuration.md).

## Data protection

- **Transport:** serve only over HTTPS (TLS 1.2+) behind the company reverse proxy, and set `COOKIE_SECURE=true`.
- **Headers:** `X-Frame-Options: DENY`, `X-Content-Type-Options: nosniff`, `Referrer-Policy: same-origin`, `Permissions-Policy`. Search engines are told not to index.
- **Input validation:** every server action validates input with Zod. Prisma uses parameterized queries. React escapes all output.
- **CSRF:** server actions are POST-only, and Next.js checks that the Origin header matches the host.
- **File uploads:** maximum 10 MB. The type is detected **by content** (PDF, PNG, JPEG, WEBP only), not by file extension. Files are stored outside the web root under random names with mode 600, and served only through an authenticated route that checks permission to view that employee, with a restrictive `Content-Security-Policy` and `nosniff`.
- **Concurrency:** qualification records and training assignments use optimistic locking (`version` column). Concurrent edits are detected and reported, never silently overwritten.
- **Audit:** `audit_logs` and `approval_records` are append-only. A PostgreSQL trigger rejects `UPDATE` and `DELETE`. The only exception is the guarded *delete sample data* action, which turns it off inside a single transaction to remove fictional approval rows.
- **Least data:** the system stores no SSNs, home addresses, dates of birth or bank details. Exports never include pay data.

## Retention and deletion (HR to approve)

Retention periods are a company decision and have **not** been set. Agree with HR and Legal:

- how long training and qualification records are kept after employment ends,
- how long audit logs and promotion decisions are kept,
- how uploaded evidence is retained and destroyed.

Until a policy is approved, mark leavers **Inactive** instead of deleting them, so their history stays available for audits.

## Deployment boundary

Run the system **only on a private network or VPN**. Do not expose it to the internet, and do not load real employee data until company IT and HR have approved it. See the [deployment checklist](deployment-checklist.md).
