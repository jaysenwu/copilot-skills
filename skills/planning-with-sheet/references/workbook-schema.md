# Workbook schema and write protocol

## Overview

Use visible worksheet `_Plan`, A1=`PWS/1`. Do not move tables or insert rows above
row 14. Keep English headers/tokens; task prose can be Chinese.

| Cell | Field | Maintenance |
| --- | --- | --- |
| B3 | Goal | User objective and deliverable. |
| B4 | Context / source snapshot | Input names, period, units, filters, assumptions, source/content checks. |
| B5 | Plan ID | Created once; retained on resume. |
| B6 | Run mode | `plan_only`, `run`, `paused`; follows user intent. |
| B7 | Last checkpoint UTC | Actual ISO UTC time of last progress write. |
| B8 | Resume summary | Last confirmed operation, partial outputs, precise next action. |
| B9 | Completed / total | Script-derived; skipped shown separately. |
| B10 | Next task | Script-derived candidate, not permission to execute. |
| B11 | Plan validation | Script-derived schema result, not output verification. |
| B12 | Revision | Nonnegative integer; increment for each logical progress write. |

## Tasks

Table `PWS_Tasks` has header A14:M14 and body from row 15. Keep one blank body row
when empty. This version supports up to 500 body rows.

| Column | Field | Rule |
| --- | --- | --- |
| A | Check | Derived ☑, ☐, — or !. |
| B | TaskID | Stable unique `T001` style, 3–6 digits. |
| C | Task | Independently verifiable work. |
| D | Status | `pending`, `in_progress`, `blocked`, `completed`, `skipped`. |
| E | DependsOn | Existing TaskIDs separated by commas; blank if none. |
| F | Output | Exact owned range/table/chart or explicitly named planning finding. |
| G | Acceptance | Observable completion test. |
| H | Checkpoint | Last successful substep, partial outputs and/or recorded intent. |
| I | NextAction | One operation/unblock action; may be blank after completion. |
| J | Evidence | Expected versus observed result, address/object and check time. |
| K | UpdatedUTC | Actual ISO UTC timestamp. |
| L | Findings | Decisions, assumptions, errors, skip/reopen reasons. |
| M | Attempts | Numeric consecutive failures, starting at 0. |

`completed` requires verified output, Checkpoint and Evidence. Active/blocked rows
need Checkpoint and NextAction. Skipped rows require a reason grounded in user
scope in Findings. A skipped prerequisite does not satisfy a dependency. After
sorting, locate every task by ID again; never use a remembered row number.

## Journal

Table `PWS_Journal`: header P14:U14, body from row 15, up to 5000 rows.
Columns: LogID, UTC, TaskID, Event, Detail, NextAction.
Use unique LogIDs (`L0001`, ...), actual timestamps, and events such as `plan`,
`start`, `checkpoint`, `complete`, `error`, `pause`, `resume`, `reopen`, `scope_change`.
Append events. The check script returns the last 40 rows; read older relevant
events separately. Do not silently archive or truncate either table.

## Write protocol

1. Reload the marker, Revision and tasks. Confirm exact headers and current mode.
   Find rows by TaskID; reconcile edits since your last snapshot.
2. For initial capture, replace the entirely blank body row with the first task,
   then append other tasks with `table.rows.add`. Never clear an existing plan.
3. Before business edits, write intent (`in_progress`, Output, Checkpoint,
   NextAction, UpdatedUTC), overview/revision and a `start` journal event.
   Synchronize and read back this intent before performing the business action.
4. After verification, write Evidence/Checkpoint and Status in the same logical
   update. Update B7, B8 and numeric B12; append the corresponding journal event.
   Synchronize and run `checkPlanningSheet()` to read back progress.
5. Reconcile failures before retrying. An Office.js batch is not a multi-operation
   transaction. A journal may lag its task row after a crash; append a recovery event.

### Office.js editing idioms

These are fragments for Copilot's native editing tool, not additional entry points.
Use actual task content as literals inside generated code. Task-specific values
are never passed to a bundled skill function.

```javascript
// Within await Excel.run(async (context) => { ... }):
const sheet = context.workbook.worksheets.getItem("_Plan");
const table = sheet.tables.getItem("PWS_Tasks");
const body = table.getDataBodyRange();
body.load(["values", "rowIndex"]);
await context.sync();
const index = body.values.findIndex((row) => row[1] === "T001");
if (index < 0) throw new Error("T001 missing; reconcile the plan.");
// After checking state/dependencies, use body.rowIndex + index for targeted edits.
// Write the complete intent/checkpoint described above, not Status alone.
```

```javascript
function asLiteral(value) {
  return typeof value === "string" && /^[\s]*[=+@-]/.test(value)
    ? "'" + value : value;
}
// Apply to planning text assigned through .values; keep numeric fields numeric.
```

Wrap task/journal body text after adding rows, and use readable row heights (for
example 54 points). Freeze the first 14 rows if the host supports it. Cosmetic
APIs must not be prerequisites for saving task progress.
