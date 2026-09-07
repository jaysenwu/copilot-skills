# Resume playbook

Read the current user request, overview, task rows and relevant journal events.
Resolve the workbook and goal before updating anything.

| Observed state | Recovery |
| --- | --- |
| Completed task; relevant output and inputs match evidence | Preserve completion and reuse the verified result. |
| In progress; output absent | Continue from recorded intent using current inputs. |
| In progress; output partially exists | Inspect owned ranges/stable keys; perform only missing work. |
| In progress; expected output already complete | Verify acceptance and commit completion without repeating the business action. |
| Completed task; output missing or wrong | Journal previous evidence/discrepancy; reopen task and affected dependents. |
| Inputs/scope changed | Identify affected work; preserve history, reopen affected tasks, update Context. |
| Checkmark disagrees with Status | Recompute display; do not infer completion from a checkmark. |
| Blocked task | Read error/unblock action; continue independent tasks only if authorized. |
| Paused/plan_only mode | Execute only if the current user request authorizes continuation. |
| Journal behind task row | Inspect output; reconcile from current state and append recovery event. |
| Unexpected IDs, headers or dependencies | Repair the discrepancy before business edits; preserve history. |

Before a non-idempotent operation such as append, chart creation or increment,
record a stable identifier and output target. On recovery, inspect that identifier
before repeating the operation. If there is no reliable way to determine whether
it occurred, mark blocked and ask for the missing fact instead of guessing.

Verify source content as well as existence. For large inputs, use appropriate
content checks, totals, selected keys, formulas or trusted version identifiers;
disclose incomplete verification. The scripts do not hash inputs or verify
business results automatically.

Give a short completed/total update, last reliable checkpoint and next action,
then perform authorized work. Do not rebuild the plan merely because the chat is new.
