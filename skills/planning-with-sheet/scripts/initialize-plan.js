/* Copilot in Excel skill entry point. Call with no arguments. */
async function initializePlanningSheet() {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    const tables = context.workbook.tables;
    sheets.load("items/name");
    tables.load("items/name");
    await context.sync();
    const existing = sheets.items.find((s) => s.name === "_Plan");
    if (existing) {
      const marker = existing.getRange("A1");
      marker.load("values");
      await context.sync();
      if (marker.values[0][0] !== "PWS/1") {
        throw new Error("_Plan already exists without the PWS/1 marker. Preserve it; resolve this name collision before initialization.");
      }
      return JSON.stringify({status: "existing_plan_preserved", next: "Run checkPlanningSheet; read the saved plan before any task edit."});
    }
    if (tables.items.some((t) => ["PWS_Tasks", "PWS_Journal"].includes(t.name))) {
      throw new Error("A PWS table already exists, possibly on a renamed planning sheet. Restore the original _Plan sheet name or resolve the collision; do not create a second plan.");
    }
    const sheet = sheets.add("_Plan");
    const labels = ["Goal", "Context / source snapshot", "Plan ID", "Run mode", "Last checkpoint UTC", "Resume summary", "Completed / total", "Next task", "Plan validation", "Revision"];
    sheet.getRange("A3:A12").values = labels.map((s) => [s]);
    sheet.getRange("B3:M3").merge(false);
    sheet.getRange("B4:M4").merge(false);
    sheet.getRange("B8:M8").merge(false);
    sheet.getRange("B11:M11").merge(false);
    const now = new Date().toISOString();
    sheet.getRange("B5").values = [["PWS-" + now.replace(/[^0-9]/g, "")]];
    sheet.getRange("B6").values = [["plan_only"]];
    sheet.getRange("B7").values = [[now]];
    sheet.getRange("B8").values = [["Plan created. Add the user goal, source context, and tasks before execution."]];
    sheet.getRange("B9").values = [["0 / 0"]];
    sheet.getRange("B11").values = [["EMPTY PLAN"]];
    sheet.getRange("B12").values = [[0]];
    const taskHeaders = ["Check", "TaskID", "Task", "Status", "DependsOn", "Output", "Acceptance", "Checkpoint", "NextAction", "Evidence", "UpdatedUTC", "Findings", "Attempts"];
    sheet.getRange("A14:M14").values = [taskHeaders];
    const tasks = sheet.tables.add("A14:M15", true);
    tasks.name = "PWS_Tasks";
    tasks.style = "TableStyleMedium2";
    sheet.getRange("P14:U14").values = [["LogID", "UTC", "TaskID", "Event", "Detail", "NextAction"]];
    const journal = sheet.tables.add("P14:U15", true);
    journal.name = "PWS_Journal";
    journal.style = "TableStyleMedium2";
    sheet.getRange("P15:U15").values = [["L0001", now, "", "initialize", "Created planning sheet; business sheets unchanged.", "Capture the user goal and task plan."]];
    sheet.getRange("B1:M1").merge(false);
    sheet.getRange("B1").values = [["PLANNING WITH SHEET"]];
    sheet.getRange("A1:M1").format.fill.color = "#103C4A";
    sheet.getRange("A1:M1").format.font.color = "#FFFFFF";
    sheet.getRange("A1:M1").format.font.bold = true;
    sheet.getRange("B1:M1").format.font.size = 18;
    sheet.getRange("A1:M1").format.rowHeight = 34;
    sheet.getRange("A3:A12").format.font.bold = true;
    sheet.getRange("A3:A12").format.fill.color = "#E8F3F6";
    sheet.getRange("B3:M4").format.wrapText = true;
    sheet.getRange("B8:M8").format.wrapText = true;
    sheet.getRange("B11:M11").format.wrapText = true;
    sheet.getRange("A3:M4").format.rowHeight = 38;
    sheet.getRange("A8:M8").format.rowHeight = 38;
    sheet.getRange("A11:M11").format.rowHeight = 38;
    const widths = [145, 115, 210, 95, 115, 190, 220, 230, 220, 250, 150, 250, 70];
    for (let i = 0; i < widths.length; i++) {
      sheet.getRangeByIndexes(0, i, 1, 1).getEntireColumn().format.columnWidth = widths[i];
    }
    sheet.getRange("P:P").format.columnWidth = 90;
    sheet.getRange("Q:Q").format.columnWidth = 150;
    sheet.getRange("R:S").format.columnWidth = 95;
    sheet.getRange("T:U").format.columnWidth = 260;
    sheet.getRange("P12").values = [["SESSION JOURNAL →"]];
    sheet.getRange("P12:U12").format.font.bold = true;
    sheet.getRange("A13").values = [["Tasks ↓ | Journal: P14 | Checks refresh when the skill runs"]];
    await context.sync();
    // Set the marker only after core initialization succeeds. Never silently reset partial state.
    sheet.getRange("A1").values = [["PWS/1"]];
    await context.sync();
    return JSON.stringify({status: "initialized", sheet: "_Plan", mode: "plan_only", next: "Write the plan using the editing protocol, then run checkPlanningSheet."});
  });
}
