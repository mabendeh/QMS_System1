# HR guide

## What HR can do

| Area | Where |
|---|---|
| See every employee's qualifications and training | Skills Matrix, Employees |
| Import or update the employee list | Administration → Import employees |
| Career level titles and promotion rules | Administration → Career levels & promotion rules |
| Verify promotion recommendations against qualification and pay policy | Promotion Review |
| Pay grades and skill premiums | HR Compensation (HR and Plant Leadership only) |
| Reports | Export buttons on the Matrix, Employee profile, Training, Promotion Review and Coverage pages |
| Audit history | Administration → Audit log |

## Importing employees

1. Prepare a CSV or Excel file with the columns **Employee ID, First Name, Last Name, Shift** (A, B or C). Optional columns: **Department, Job Title, Hire Date** (YYYY-MM-DD) and **Supervisor Email**.
2. Go to **Administration → Import employees** and click **Preview import**.
3. The preview lists every row and every problem: unknown shift, department or job title, duplicate IDs in the file, and IDs that already exist. Rows with problems are skipped.
4. Existing employees are **left unchanged** unless you tick *Also update details of existing employees*. **An import never changes qualifications.**
5. Click **Import**. The result, and who ran it, is recorded in the audit log.

## Career levels and promotion rules

The starting levels (Trainee → Qualified Operator → Multi-Skilled Operator → Advanced Operator → Certified Trainer) and their rules are **demonstration values**. Rename levels and set the real requirements before use.

Rule types you can combine:

- Qualified in the employee's assigned department
- Qualified in a specific department (for example Food Safety & GMP)
- Qualified in at least *N* production departments
- A specific competency is Competent
- Currently at career level *N* or above
- Manual confirmation during review (for example "Completed approved train-the-trainer program")

Clicking **Publish new version** never edits the old rules. It creates version *N+1*, and requests already submitted keep the version they were checked against. Time employed is deliberately not available as a rule.

## Promotion workflow

1. **Supervisor** submits a recommendation. The system will only accept it once every system-checked requirement is met, and it stores a snapshot of the evaluation.
2. **Production Manager** reviews and approves or rejects.
3. **HR** verifies qualification and pay policy, then approves or rejects.
4. **Plant Leadership** gives final approval and enters the effective date.

Rules that apply at every step:

- A reason is required to reject.
- The same person cannot decide two steps, and the person who submitted cannot review their own request.
- Every decision is stored permanently, with reviewer, comments and date. The database refuses edits or deletions of decision records.
- Final approval records the new career level. **It does not change pay.** Process approved pay changes through payroll under CHG compensation policy.

## Compensation configuration

**HR Compensation** is visible only to HR Administrators (can edit) and Plant Leadership (view only). Supervisors, trainers and system administrators cannot open it. The server refuses the request, not just the menu.

- **Pay grades:** code, name, linked career level, hourly minimum and maximum, effective and end dates, approval authority, notes.
- **Skill premiums:** name, optional department, amount, unit, effective dates, approval authority.
- No amounts are pre-filled. An amount requires an effective date, and the minimum cannot exceed the maximum.
- Every change is audited with before and after values. Compensation audit entries are hidden from roles without pay access.
- Exports never include compensation data.

## Reports

| Report | Where | Formats |
|---|---|---|
| Complete skills matrix (overview plus one sheet per department) | Skills Matrix → Excel / PDF | XLSX, PDF |
| Department skills matrix | Skills Matrix, department tab → Excel / PDF | XLSX, PDF |
| Employee qualification report | Employee profile → PDF report / Excel | XLSX, PDF |
| Training gaps | Training → Export gaps | XLSX, PDF |
| Promotion review report | Promotion Review → Export report | XLSX |
| Shift coverage report | Shift Coverage → Excel / PDF | XLSX, PDF |
| Monthly skills & training report | CI Analytics → Monthly report | XLSX, PDF |

Excel exports keep the status colours **and** the text labels, with a legend and a tooltip on each cell.

## Before going live

Work through the [deployment checklist](deployment-checklist.md) with IT. In particular: confirm departments, checklists, rules and requalification intervals; agree retention and deletion periods; and delete the sample data.
