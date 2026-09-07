# planning-with-sheet · Excel Copilot 技能包

版本：1.0.0 ｜ 文档核实日期：2026-09-07

把 Copilot 的多步骤计划保存到当前工作簿的 `_Plan` 工作表；执行时逐项更新，
中断后依据检查点和实际产出继续。运行文件全部使用 Office.js，不需要 Python。

## 安装到 OneDrive

1. 解压 ZIP。找到直接包含 `SKILL.md` 的 **planning-with-sheet** 文件夹。
2. Excel → Copilot → 右上角 `…` → **Manage skills** → **Custom skills**。
3. 首次使用时选择 **Create OneDrive folder**，然后点 **Open skills folder**。
   以这个入口打开的位置为准，不要猜测 OneDrive 路径。
4. 将 **planning-with-sheet 整个文件夹**上传到该位置，保留 scripts 和 references。
   上传后应是 `…/planning-with-sheet/SKILL.md`，不要多套一层同名文件夹。
5. 回到 **Custom skills** 点 **Refresh**，在 **Manage skills** 启用该技能。
6. 在提示框的 **Add work content → Choose skills** 选择它，或输入
   `@planning-with-sheet` 调用。

**不要只上传 ZIP。** 当前安装说明要求技能文件夹；本 ZIP 是交付容器。
文件夹名必须与 SKILL.md 的 `name` 一致，不能改名；注意文件不能变成 `SKILL.md.txt`。
这份中文说明和 SOURCES-AND-VALIDATION.md 无需上传至技能目录。

微软当前说明要求 Office 显示语言为英文以访问 Excel Skills，因此技能正文使用英文。
下方给出英文调用语句，任务内容可以要求用中文记录。如果找不到菜单，先检查语言、
租户开放情况和 Excel 版本。Office.js 扩展开发文档仍标注 Preview，并列出
2608（Build 20305.20002）及以上、Beta / Current Channel (Preview) 条件；
个人 OneDrive Skills 与开发者插件分发是不同入口，以你实际账号可用能力为准。
无需为了这个技能创建 manifest、部署网站或购买 Copilot Studio。

依据：[微软 Excel Skills 安装说明](https://support.microsoft.com/en-us/excel/copilot/copilot-in-excel-skills)、
[Office.js Skills 开发概览](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-skills)。

## 直接可用的提示词

**已有 Copilot Plan：先保存，不执行业务任务。**

```text
@planning-with-sheet Save the plan you just generated into a separate _Plan
worksheet in this workbook. Preserve its scope and order. Include task IDs,
dependencies, acceptance criteria, checkpoints and next actions.
Keep Run mode as plan_only. Do not execute the business tasks yet.
Write task descriptions in Chinese.
```

如果当前处于不能修改工作簿的 Plan / Chat 模式，切换到 **Edit** 再执行保存。
如果换了会话，Copilot 已看不到原 Plan，需把计划文本带入这条请求。

**保存后开始执行。**

```text
@planning-with-sheet Execute the saved plan in _Plan. Set Run mode to run.
Before each task, record the intended output and checkpoint. After each task,
verify the actual result, update the task status and check marks, and save the
next action. Continue until the authorized tasks are complete or genuinely blocked.
```

**中断后，在同一个已保存工作簿的新会话继续。**

```text
@planning-with-sheet Resume from the saved _Plan sheet. Reconcile the existing
outputs before continuing. Keep verified completed tasks, resolve any in-progress
task without duplicating its output, and continue the remaining work.
```

**主动暂停。**

```text
@planning-with-sheet Pause now. Save the last successful operation, partial
output locations, errors and the exact next action in _Plan. Set Run mode to paused.
```

关闭 Excel 前保存工作簿，或等待 AutoSave 完成。检查点存储在工作簿里；
如果修改尚未保存，下一次打开就不能保证读到最近的检查点。

## 工作表里有什么

| 区域 | 内容 |
| --- | --- |
| `_Plan` 顶部 | 目标、输入背景、运行模式、最近检查点、恢复摘要、完成数、下一任务 |
| A14:M 起的任务表 | ID、任务、状态、依赖、输出位置、验收条件、检查点、下一动作、验证记录 |
| P14:U 起的日志表 | 每次计划、开始、完成、异常、暂停、恢复和范围调整的记录 |

状态使用 `pending / in_progress / blocked / completed / skipped`。
勾选显示为 `☑ / ☐ / — / !`，由检查脚本刷新；它们是状态指示符，不是可点击的
原生复选框。手工勾选不能替代验收。脚本不会自动认定业务分析正确。

一份工作簿对应一个活动计划。相关追加需求可以增加新任务，保留原记录；
本版不提供多项目切换和多人并发写入控制。不要修改 `_Plan`、表名、表头和表格起点。
任务正文可中文；请保留状态与字段名为英文。

## 这个版本怎样处理断点

| 中断情况 | 恢复处理 |
| --- | --- |
| 任务尚未执行 | 从保存的下一动作继续 |
| 已写入部分结果 | 检查实际输出，只补缺失部分 |
| 结果已完成，状态仍是 in_progress | 验证结果后补记完成，不重复创建或追加 |
| 显示 completed，但结果被删除/数据已变化 | 保留历史证据，重新打开受影响任务及后续依赖 |

**恢复能力的边界：** 本技能没有后台运行或内置 Plan 事件钩子。Excel 重新打开后，
仍需让 Copilot 调用这个技能；它不会像程序调度器一样自行开工。执行阶段的逐项更新
由 Copilot 遵循技能流程完成，不能提供系统级强制保证。提示词中的显式 @ 调用最清楚。

## 可选：让工作簿在新会话更容易想起计划

如果你希望增加工作簿级提醒，可以自行在可见的 `.Rules` 工作表 A 列追加以下规则，
每条一个单元格，保留已有规则。技能不会自动创建或修改 `.Rules`。

```text
For requests concerning the tracked goal in _Plan, read its current task list,
source context and checkpoints before choosing the next action.
```

```text
When continuing the tracked goal, use planning-with-sheet if available. Reconcile
actual outputs before repeating an operation and update the checkpoint after work.
```

这是额外提醒，不是保证触发脚本的钩子。与本计划无关的简单问题不应强制走任务清单。
依据：[微软 workbook rules 文档](https://support.microsoft.com/en-us/excel/copilot/copilot-in-excel-rules)。

## 首次上线验收（建议用副本）

1. 给 Copilot 一个三步测试任务：检查现有小表 → 生成汇总 → 创建图表。
   先用“仅保存计划”提示词，确认 `_Plan` 存在且业务任务尚未执行。
2. 让它只执行第一步并暂停。确认第一步有 Evidence 和勾选、剩余任务未完成。
3. 保存并开启新会话，用恢复提示词。确认沿用同一计划，第一步没有重复执行。
4. 验证最终完成数、输出与验收条件一致。可在副本中将一个已生成图表的任务状态
   改为 `in_progress`，再恢复，检查是否复用已有图表而没有新增同名图表。

已做本地脚本与流程测试，但未连接你的 Microsoft 365 租户，也未在真实 Excel
Copilot 中完成导入和端到端运行；上面四步是你环境中的最终验收。

## 常见问题

- **脚本不能运行：** 技能文件不会额外提供宿主没有的工具。技能已包含指引，让
  Copilot 通过已有的授权 Office.js 编辑能力执行同一逻辑；若该能力也不可用，
  就需要在支持的 Excel Copilot 编辑环境中运行，不能声称已经写入。
- **提示 `_Plan` 冲突：** 保留现有工作表，确认它是否是其他用途或中途初始化失败。
  处理命名/结构后再运行，不要要求技能清空它。
- **新会话没读旧计划：** 确认打开的是同一个已保存文件，并显式 @ 调用技能。
- **手工改了任务：** 让 Copilot 刷新计划检查；修改范围会影响后续任务时需重新验收。
- **容量：** 当前每个计划最多 500 个任务行、5000 条日志；达到上限时显式归档，
  不会自动删掉旧记录。
