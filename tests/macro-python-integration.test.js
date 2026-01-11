import assert from "node:assert/strict";
import test from "node:test";

import { MacroRecorder, generatePythonMacro } from "../apps/desktop/src/macro-recorder/index.js";
import { Workbook } from "../packages/scripting/src/index.js";
import { NativePythonRuntime } from "../packages/python-runtime/src/native-python-runtime.js";
import { MockWorkbook } from "../packages/python-runtime/src/mock-workbook.js";

test("macro recorder generates runnable Python that replays simple edits", async () => {
  const workbook = new Workbook();
  workbook.addSheet("Sheet1");
  workbook.setActiveSheet("Sheet1");

  const recorder = new MacroRecorder(workbook);
  recorder.start();

  workbook.setSelection("Sheet1", "A1");
  const sheet = workbook.getActiveSheet();
  sheet.setCellValue("A1", 10);
  sheet.setCellValue("B1", 32);
  sheet.getRange("A1:B1").setFormat({ bold: true });
  sheet.setRangeValues("C1:D2", [
    [1, 2],
    [3, 4],
  ]);
  workbook.setSelection("Sheet1", "A2");

  const actions = recorder.stop();
  const script = generatePythonMacro(actions);

  const api = new MockWorkbook();
  const runtime = new NativePythonRuntime({ timeoutMs: 10_000 });
  await runtime.execute(script, { api, timeoutMs: 10_000 });

  const sheetId = api.get_sheet_id({ name: "Sheet1" });
  assert.ok(sheetId);

  assert.deepEqual(
    api.get_range_values({
      range: { sheet_id: sheetId, start_row: 0, start_col: 0, end_row: 0, end_col: 1 },
    }),
    [[10, 32]],
  );
  assert.deepEqual(
    api.get_range_values({
      range: { sheet_id: sheetId, start_row: 0, start_col: 2, end_row: 1, end_col: 3 },
    }),
    [
      [1, 2],
      [3, 4],
    ],
  );
  assert.deepEqual(
    api.get_range_format({
      range: { sheet_id: sheetId, start_row: 0, start_col: 0, end_row: 0, end_col: 1 },
    }),
    { bold: true },
  );
  assert.deepEqual(
    api.get_range_format({
      range: { sheet_id: sheetId, start_row: 0, start_col: 1, end_row: 0, end_col: 1 },
    }),
    { bold: true },
  );
  assert.deepEqual(api.get_selection(), {
    sheet_id: sheetId,
    start_row: 1,
    start_col: 0,
    end_row: 1,
    end_col: 0,
  });
});
