# Copilot Skills

Reusable custom skills for Microsoft 365 Copilot, with a focus on practical
Excel workflows.

## Available skills

### planning-with-sheet

`planning-with-sheet` turns an Excel workbook into durable working memory for
multi-step Copilot tasks. It creates a dedicated `_Plan` worksheet, records task
dependencies and acceptance criteria, checkpoints progress after each verified
step, and helps Copilot resume after an interrupted session without duplicating
completed output.

Key capabilities:

- Save a Copilot-generated plan into the workbook.
- Track `pending`, `in_progress`, `blocked`, `completed`, and `skipped` tasks.
- Store checkpoints, findings, evidence, errors, and the next action.
- Reconcile actual workbook output before resuming.
- Use Office.js only; no Python, external services, or add-in deployment.

Go to [`skills/planning-with-sheet`](skills/planning-with-sheet) for the source.
Download [`planning-with-sheet-onedrive-v1.0.1.zip`](dist/planning-with-sheet-onedrive-v1.0.1.zip)
for the OneDrive upload package.

## Install in Copilot for Excel

1. Download and unzip the package.
2. In Excel, open Copilot → **…** → **Manage skills** → **Custom skills**.
3. Select **Create OneDrive folder** if this is your first custom skill, then
   select **Open skills folder**.
4. Upload the complete `planning-with-sheet` folder. Keep its `scripts` and
   `references` subfolders.
5. Return to Excel and select **Refresh** in the Custom skills dialog.
6. Use the skill in **Edit mode** with `@planning-with-sheet`.

Microsoft currently documents Excel skills as English-only, so the skill
instructions and fixed workbook field names are in English. The tracked task
descriptions can be written in Chinese.

See the [Chinese installation and usage guide](docs/START-HERE-zh-CN.md) for
ready-to-use prompts and the first-run acceptance test.

## Recommended workflow

Use Copilot's Plan mode to shape a plan if needed. Switch to Edit mode and ask:

```text
@planning-with-sheet Save the plan you just generated into a separate _Plan
worksheet in this workbook. Keep Run mode as plan_only and do not execute the
business tasks yet.
```

To resume later in the same saved workbook:

```text
@planning-with-sheet Resume from the saved _Plan sheet. Reconcile existing
outputs before continuing the remaining tasks.
```

The skill does not run automatically when Excel opens. Invoke it in an
editing-capable Copilot session, and save the workbook or allow AutoSave to
complete before closing.

## Repository structure

```text
copilot-skills/
├── skills/
│   └── planning-with-sheet/
│       ├── SKILL.md
│       ├── LICENSE.txt
│       ├── NOTICE.md
│       ├── scripts/
│       └── references/
├── docs/
├── dist/
├── NOTICE.md
└── LICENSE
```

## Validation

The JavaScript entry points passed 24 local behavioral scenarios using an
in-memory Office.js test double. This includes idempotent initialization,
dependency validation, interruption recovery, completion controls, bounded
reads, and preservation of existing workbook content. Live import and runtime
behavior must still be verified in the target Microsoft 365 tenant.

See [sources and validation](docs/SOURCES-AND-VALIDATION.md) for the source list,
adaptation decisions, and test boundary.

## 中文简介

这个仓库存放面向 Microsoft 365 Copilot 的自定义技能。
首个技能 `planning-with-sheet` 会在 Excel 中创建 `_Plan` 工作表，把任务计划、
执行状态、验收记录和下一步动作保存在工作簿内。即使会话中断，也可以在新会话中
先核对实际产出，再从断点继续。

实际使用时，请在 Excel Copilot 的 **Edit 模式**调用技能。Plan 模式可以用来
先生成计划，但写入 `_Plan`、更新状态和执行任务都需要 Edit 模式。

## Copyright, attribution, and upstream credit / 版权与转载说明

The Excel adaptation and original implementation in this repository are
Copyright © 2026 Jaysen Wu and are licensed under the MIT License.

If you copy, modify, redistribute, or incorporate this project or a substantial
portion of it, retain the copyright and permission notice in [LICENSE](LICENSE).
That required notice identifies `planning-with-sheet`, Jaysen Wu, and the source
repository URL. Please also keep [NOTICE.md](NOTICE.md) with the distribution.

本仓库中的 Excel 适配方案及新增实现版权归 Jaysen Wu 所有，并按 MIT License
开源。按照 MIT License，转载、复制、修改、再发布或整合本项目的实质性内容时，
须保留 [LICENSE](LICENSE) 中的完整版权声明与许可文本；该版权声明已明确包含
**Jaysen Wu、planning-with-sheet 及项目链接**。建议同时保留 [NOTICE.md](NOTICE.md)。

This project was inspired by and adapts the persistent-planning pattern from
[OthmanAdi/planning-with-files](https://github.com/OthmanAdi/planning-with-files),
Copyright © 2026 Ahmad Adi, also licensed under the MIT License. The upstream
copyright and license are preserved in
[`UPSTREAM-LICENSE.txt`](skills/planning-with-sheet/UPSTREAM-LICENSE.txt).
This repository is an independent Excel-oriented adaptation; it is not endorsed
by or affiliated with the upstream author or Microsoft.

See [NOTICE.md](NOTICE.md) for the recommended attribution wording and the scope
of each notice.
