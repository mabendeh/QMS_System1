# Production deployment checklist

Work through this with **company IT, HR and the project owner**. Do not load real employee data until every item marked **Required** is signed off.

## A. Approvals (Required)

- [ ] IT security has reviewed [security.md](security.md) and approved hosting on the private network.
- [ ] HR has approved processing employee qualification data in this system.
- [ ] HR and Legal have approved **retention and deletion periods** for qualification records, evidence files, promotion decisions and audit logs.
- [ ] HR has confirmed who holds each role, especially HR Administrator and Plant Leadership, which can see pay data.

## B. Configuration (Required)

- [ ] Departments match the actual Duncanville plant layout (Administration → Departments & competencies).
- [ ] Each department's competency checklist has been reviewed by its area owner, and critical skills are flagged.
- [ ] The decision on **Proofing** and **Baking** is recorded. They ship disabled because of the frozen-only scope.
- [ ] Requalification intervals are confirmed by Quality and Safety.
- [ ] Job titles and their required departments are confirmed.
- [ ] Career level titles and promotion rules are approved by HR. The starting values are demonstration only.
- [ ] Pay grades and premiums have been entered by HR from the **approved** compensation policy, or left blank.
- [ ] Shift staffing targets have been entered by Production Managers.

## C. Infrastructure (Required)

- [ ] Server is on the private network or VPN only, and not reachable from the internet.
- [ ] HTTPS via the company reverse proxy (TLS 1.2+), with `COOKIE_SECURE=true`.
- [ ] Reverse-proxy rate limiting on `/login`.
- [ ] PostgreSQL 14+ with a dedicated least-privilege database user and a strong password stored in the server's secret store.
- [ ] `.env` is readable only by the service account. `UPLOAD_DIR` is outside the web root, mode 700.
- [ ] The app runs as a service (systemd, Windows service or container) with automatic restart. Node 20 or 22.
- [ ] Nightly backups of PostgreSQL **and** `UPLOAD_DIR`, kept off-server. A test restore has been completed.
- [ ] Server time zone and NTP are set (the app displays America/Chicago).
- [ ] Monitoring or alerting on service health and disk space.

## D. Accounts (Required)

- [ ] Configuration-only database created: `npm run db:migrate && SEED_SKIP_DEMO=true npm run db:seed`, **or** sample data deleted from a demo database (Administration → Sample data, or `CONFIRM=DELETE npm run demo:purge`).
- [ ] No `@demo.local` accounts are active.
- [ ] First real administrator created with `npm run admin:create`, and the password changed after first sign-in.
- [ ] SSO connected (optional; see [configuration.md](configuration.md)), or a password policy agreed for local accounts.
- [ ] Every user has only the roles they need. Trainers are authorized only for their areas.

## E. Data load

- [ ] Employee list imported via Administration → Import employees, preview checked, duplicates resolved.
- [ ] Existing paper or Excel qualifications entered through assessments, with practical evaluations and evidence for critical skills. The system will not mark anyone Competent without them.

## F. Verification before go-live

- [ ] `npm test` passes (unit and component tests).
- [ ] `npm run test:e2e` passes against a **non-production** database server (all 10 acceptance scenarios plus role smoke tests).
- [ ] A supervisor completes an assessment on the floor tablet in under 60 seconds.
- [ ] HR confirms a supervisor account cannot open HR Compensation.
- [ ] An Excel export has been opened and checked by HR.
- [ ] Pilot users have received the [supervisor guide](supervisor-guide.md) and [HR guide](hr-guide.md).

## G. After go-live

- [ ] Monthly report reviewed at the leadership meeting (CI Analytics → Monthly report).
- [ ] Audit log reviewed monthly by HR or Quality.
- [ ] Quarterly access review: remove leavers and adjust roles.
- [ ] Quarterly backup restore test.
