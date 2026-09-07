# Excel execution boundary

Run inside Copilot in Excel with workbook editing access. The host supplies
Office.js. Bundled entry points take zero arguments; internal Office.js API calls
take their normal arguments. Do not confuse these restrictions.

Run `initializePlanningSheet()` for a new tracked task only. It preserves an
existing initialized plan. Run `checkPlanningSheet()` for planning, checkpoints,
resumption and completion checks. It changes only derived Check, Completed / total,
Next task and Plan validation cells. It does not execute business tasks or change
their statuses. For a read-only status request, use native read tools instead.

Write task-specific content with normal Copilot editing tools. In generated
Office.js, use the real values as literals inside the code, not parameters passed
to a bundled entry point. No filesystem import, shared global state, dynamic code
loading, `eval`, external libraries or interactive prompts are required.

If Plan/Chat mode cannot edit, say the user must switch to Edit mode to save the
plan. Outside Excel, explain where to invoke the skill and do not claim workbook
changes. A draft in chat is not a saved checkpoint.

Office.js synchronizes changes to the open workbook; Excel separately saves the
file. This package has no background daemon, event listener, native Plan hook,
transaction across business/status writes, or guaranteed automatic invocation.

After any script failure, read current state before retrying: an earlier sync
may have succeeded. Preserve partial initialization and resolve schema errors
explicitly instead of clearing or recreating existing data.
