# CHG Production Skills & Career Progression System

A private web application for the C.H. Guenther / Pioneer Frozen Foods plant in Duncanville, Texas. It lets you:

- keep a colour-coded **skills matrix** for every production employee,
- **assess** a skill in under a minute, from a PC, tablet or phone,
- run **training** using TWI steps (Prepare, Present, Try Out, Follow Up), with practical evaluations, evidence and verification,
- make **promotion recommendations** based on documented skill rules, with an approval workflow and a permanent decision history,
- see **shift coverage** and **continuous-improvement analytics**,
- keep **pay configuration** visible only to HR and plant leadership.

> **Status: private pilot.** All employees in the starting data are fictional (IDs start with `DEMO-`, and each shows a *Sample* badge). Do not load real employee data until company IT and HR have approved the deployment (see [docs/deployment-checklist.md](docs/deployment-checklist.md)).

---

## Quick start (local pilot)

Requirements: **Node.js 20 or 22**, **PostgreSQL 14+**.

```bash
cd chg-skills
npm install

# 1. Create a database and a user (example with psql)
psql -U postgres -c "CREATE ROLE chg LOGIN PASSWORD 'choose-a-strong-password' CREATEDB;"
psql -U postgres -c "CREATE DATABASE chg_skills OWNER chg;"

# 2. Configure
cp .env.example .env
#   edit .env: DATABASE_URL, SEED_DEMO_PASSWORD, UPLOAD_DIR (and COOKIE_SECURE=true behind HTTPS)

# 3. Create the tables and load configuration + sample data
npm run db:migrate
npm run db:seed

# 4. Build and start
npm run build
npm start            # http://localhost:3000
```

Sign in with any demo account below. The password is the `SEED_DEMO_PASSWORD` you set in `.env`. If you left it empty, the seed generated a password and printed it once.

| Account | Role | What to try |
|---|---|---|
| `supervisor.a@demo.local` (also `.b`, `.c`) | Supervisor, shifts A/B/C | Matrix, Quick Assessment, recommending a promotion |
| `trainer@demo.local` | Trainer | Training list, recording practical evaluations |
| `manager@demo.local` | Production Manager | Verifying qualifications, coverage targets, first promotion review |
| `quality@demo.local` | Quality & Food Safety | Verifying qualifications, audit log |
| `hr@demo.local` | HR Administrator | Pay configuration, promotion rules, employee import, HR verification |
| `leadership@demo.local` | Plant Leadership | Final promotion approval, analytics |
| `ci@demo.local` | CI Manager | CI analytics, monthly report |
| `sysadmin@demo.local` | System Administrator | Users, departments, sample-data removal (cannot see pay) |
| `employee@demo.local` | Employee | Own qualifications and career path only |

To create your first **real** administrator without any demo data:

```bash
SEED_SKIP_DEMO=true npm run db:seed        # configuration only, on an empty database
ADMIN_EMAIL=you@company.com ADMIN_NAME="Your Name" ADMIN_ROLES=SYSTEM_ADMIN,HR_ADMIN npm run admin:create
```

## Everyday commands

| Command | What it does |
|---|---|
| `npm run dev` | Development server with live reload |
| `npm run build` / `npm start` | Production build / start (port 3000) |
| `npm run db:migrate` | Apply database migrations (safe, non-destructive) |
| `npm run db:seed` | Load configuration and sample data into an empty database |
| `npm run admin:create` | Create or reset an administrator account (password from `ADMIN_PASSWORD`, or generated) |
| `CONFIRM=DELETE npm run demo:purge` | Delete all sample employees and disable `@demo.local` accounts |
| `npm test` | Unit and component tests (Vitest + React Testing Library) |
| `npm run test:e2e` | Browser acceptance tests (Playwright). They create a throw-away database for each run and drop it afterwards |
| `npm run typecheck` | TypeScript check |

## Documentation

| Guide | For |
|---|---|
| [Supervisor guide](docs/supervisor-guide.md) | Supervisors, trainers, leads |
| [HR guide](docs/hr-guide.md) | HR, plant leadership |
| [Administrator guide](docs/admin-guide.md) | System administrators |
| [Configuration](docs/configuration.md) | Environment variables, departments, rules, SSO |
| [Security & access control](docs/security.md) | IT, HR, audit |
| [Database schema](docs/database.md) | IT, developers |
| [Production deployment checklist](docs/deployment-checklist.md) | IT, HR, project owner |

## Technology

Next.js 15 (App Router, server actions), React 19, TypeScript, Tailwind CSS with shadcn-style components, PostgreSQL with Prisma ORM, Recharts, ExcelJS, PDFKit, Vitest, React Testing Library and Playwright.

## Status colours

| Colour | Label | Meaning |
|---|---|---|
| 🔴 Red `NT` | Needs Training | Required for the job, not yet demonstrated |
| 🟡 Yellow `IT` | In Training | Learning; works under supervision |
| 🟢 Green `TR` | Trained | Training complete; waiting for independent competency verification |
| 🔵 Blue `C` | Competent | Verified to operate independently; passed qualification |
| ⚪ Grey `—` | Not Applicable / Not Assessed | Not required for this job, or not assessed |
| Dark red `EXP` | Expired | Requalification overdue; not counted as qualified |
| Striped + `P` | Pending | A change is waiting for verification |

Every cell also shows its abbreviation and has a tooltip and screen-reader label, so the matrix never relies on colour alone.

## Starting assumptions you must confirm

- **Departments and competency checklists** come from the project brief and are *starting assumptions*. Verify them against the actual Duncanville layout and procedures (Administration → Departments & competencies).
- **Proofing** and **Baking** are configured but **inactive**, because the earlier scope was frozen operations only. Turn them on in Administration if you need them.
- **Career levels, promotion rules, requalification intervals and coverage targets** are demonstration values, not official CHG policy.
- **No pay amounts** are pre-filled. HR enters only approved values.
- **No company logo** is used. Add one only when an authorized file is provided.
- The brief mentioned a reference screenshot of a skills matrix, but no screenshot was attached. The matrix layout follows the written description: fixed employee columns on the left and angled skill headings across the top.
