# Administrator guide

## Departments and competencies

**Administration → Departments & competencies** (System Administrators and HR):

- **Add** a department (production department, or plant-wide training such as Food Safety & GMP).
- **Rename** it, edit its operating notes, or **disable** it. Disabled departments disappear from the matrix but their history is kept.
- **Reorder** departments and competencies with the arrows. The matrix follows this order.
- **Edit the checklist:** add competencies, rename them, mark them **critical** (practical evaluation and evidence become mandatory), or deactivate them.
- **Requalification:** set months per department, or per competency (the competency setting wins). New intervals apply to qualifications verified from then on.

None of this needs a code change.

## Job titles

**Administration → Job titles** sets which departments each job title must be qualified in. Required competencies that haven't been assessed show red (*Needs Training*) and appear in training gaps.

## Users and permissions

**Administration → Users & permissions:**

- Create users with a work email, name and an initial password (12+ characters). Changing a password signs the user out everywhere.
- Give each person **only the roles they need** (least privilege). The roles are HR Administrator, Production Manager, Supervisor, Trainer / Training Coordinator, CI Manager, Quality & Food Safety, Plant Leadership, Employee and System Administrator.
- **Trainers** must be authorized per department. They can only record practical evaluations and assessments there.
- **Employee** accounts must be linked to an employee record. They only ever see that person's own data.
- Disabling a user ends their sessions immediately.
- A System Administrator role does **not** include pay access. Add HR Administrator only if that person is approved for it.

See [security.md](security.md) for the full permission table.

## Coverage targets

Production Managers set the minimum qualified headcount per department and shift under **Shift Coverage → Edit staffing targets**.

## Sample data

Before real use, remove the fictional records. Either:

- go to **Administration → Sample data**, type `DELETE SAMPLE DATA` and confirm, or
- run `CONFIRM=DELETE npm run demo:purge`, which also disables every `@demo.local` account.

Both remove sample employees with their skills, assessments, training, evidence records and promotion requests. Configuration and real employees are untouched, and an audit entry records the deletion. Any evidence files uploaded against sample employees remain in `UPLOAD_DIR` and can be deleted manually.

## Audit log

**Administration → Audit log** lists every sign-in, assessment, verification, training change, promotion decision, configuration change, export and import. The `audit_logs` and `approval_records` tables are protected by database triggers that reject `UPDATE` and `DELETE`.

## Backups

Back up PostgreSQL **and** the upload directory together. For example, nightly:

```bash
pg_dump --format=custom --file=/backups/chg_skills_$(date +%F).dump "$DATABASE_URL"
tar czf /backups/chg_uploads_$(date +%F).tgz -C "$(dirname "$UPLOAD_DIR")" "$(basename "$UPLOAD_DIR")"
```

Test a restore into a separate database at least once a quarter.

## Upgrades

```bash
git pull
npm ci
npm run db:migrate      # applies new migrations; never resets data
npm run build
# restart the service
```
