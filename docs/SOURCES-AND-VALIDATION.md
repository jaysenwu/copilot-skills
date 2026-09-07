# Sources and validation

Package: planning-with-sheet 1.0.1 · Reviewed 2026-09-07

## Primary sources

| Source | What was used |
| --- | --- |
| [Microsoft: Use skills with Copilot in Excel](https://support.microsoft.com/en-us/excel/copilot/copilot-in-excel-skills) | OneDrive personal skill installation, matching folder/name, SKILL.md frontmatter, optional scripts/references, refresh, @ invocation, English display-language requirement. |
| [Microsoft: Excel skills overview](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-skills) | Excel metadata tag, workbook outputs, Office.js execution/runtime, execution boundaries and Preview developer requirements. |
| [Microsoft: Office.js skill tutorial](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-copilot-skill) | Parameterless entry points in the current preview; no Office.onReady/Office.initialize; scripts supplied as skill resources. |
| [Microsoft: Excel Copilot FAQ](https://support.microsoft.com/en-us/excel/copilot/frequently-asked-questions-about-copilot-in-excel) | Plan/Edit/Chat distinction; skills use existing authorized execution capabilities. |
| [Microsoft: Workbook rules](https://support.microsoft.com/en-us/excel/copilot/copilot-in-excel-rules) | Optional visible .Rules sheet with one rule per cell in column A. |
| [Agent Skills specification](https://agentskills.io/specification) | Portable folder layout and naming/frontmatter conventions. |
| [Original planning-with-files skill](https://github.com/OthmanAdi/planning-with-files/blob/master/skills/planning-with-files/SKILL.md) | Persistent plan/findings/progress, checkpoints, recovery, error history and completion checks. Inspected version marker: 3.17.0, on the mutable master branch. |

## Adaptation decisions

This is an independent Excel implementation inspired by and adapting the upstream
persistence pattern. It is not an official Microsoft product, is not endorsed by
the upstream author, and is not a verbatim port of upstream scripts. The upstream
copyright and complete MIT license notice are included in the skill folder as
`UPSTREAM-LICENSE.txt`. The copyright and MIT license for the Excel adaptation
are included separately as `LICENSE.txt`; `NOTICE.md` explains the attribution
scope and provides recommended wording for redistribution.

| Original concept | Excel adaptation |
| --- | --- |
| task_plan.md | `_Plan` overview and `PWS_Tasks` |
| findings.md | Context, Findings and evidence-bearing journal entries |
| progress.md | Checkpoint, NextAction, UpdatedUTC and `PWS_Journal` |
| Re-read before decisions | Skill instructions plus explicit parameterless check script |
| Host hooks / stop gate | No equivalent assumed; workflow reminders and explicit @ invocation |
| Session-file recovery | Saved workbook state plus reconciliation of actual outputs |

The OneDrive upload folder deliberately contains no M365 app manifest, installable
add-in, Python, shell script, Office Scripts TypeScript, or external dependency.
The Microsoft developer tutorial also describes a separate M365 plugin packaging
route; that package model is not required for this personal OneDrive deliverable.

## Validation performed

- Frontmatter, naming, relative references and package layout checked locally.
- Both JavaScript files parsed and executed in an in-memory Office.js test double.
  The test double checks requested property loads and range write dimensions.
- 24 behavioral cases passed: initialization, idempotence, existing-name collisions,
  renamed-plan protection, empty plans, blank-row filtering, ready dependencies,
  active-task recovery priority, preservation of business data and statuses,
  missing evidence, duplicate IDs, missing dependencies, cycles, skipped prerequisites,
  multiple active tasks, paused/plan-only modes, all-skipped plans, completion
  candidates, preservation of formula-like literal evidence, incomplete prerequisites, independent
  tasks with blocked branches, repeated failures, malformed headers, bounded reads,
  recent journal limits, and zero-argument entry points. Some cases cover multiple assertions.
- A separate forward-use exercise used a paused plan with a chart already created
  but still recorded as in_progress. The skill led to verifying/reusing the chart,
  committing recovery state, then running the remaining dependent task.

## What has not been verified

No live Microsoft 365 tenant, OneDrive Skills import, Copilot model behavior,
real Office.js host execution, AutoSave round trip, or native Plan handoff was
available for end-to-end verification. The local test double is not Excel.
There is no claim of Microsoft certification or guaranteed automatic invocation.

The script validates plan structure and evidence presence. It deliberately reports
`businessOutputsVerifiedByScript: false`; Copilot must inspect real outputs.
Check marks/summary are snapshots refreshed by the script, not live formulas or
native checkboxes. The user guide contains the final tenant acceptance procedure.
