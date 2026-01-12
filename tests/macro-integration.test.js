import assert from "node:assert/strict";
import test from "node:test";

import { MacroRecorder, generateTypeScriptMacro } from "../apps/desktop/src/macro-recorder/index.js";
import { ScriptRuntime, Workbook } from "../packages/scripting/src/node.js";

test("macro recorder generates runnable TypeScript that replays simple edits", async () => {
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
  workbook.setSelection("Sheet1", "A2");

  const actions = recorder.stop();
  const script = generateTypeScriptMacro(actions);

  const freshWorkbook = new Workbook();
  freshWorkbook.addSheet("Sheet1");
  freshWorkbook.setActiveSheet("Sheet1");

  const runtime = new ScriptRuntime(freshWorkbook);
  const result = await runtime.run(script, { timeoutMs: 30_000 });
  assert.equal(result.error, undefined, result.error?.message);

  const freshSheet = freshWorkbook.getActiveSheet();
  assert.deepEqual(freshSheet.getRange("A1:B1").getValues(), [[10, 32]]);
  assert.deepEqual(freshSheet.getRange("A1").getFormat(), { bold: true });
  assert.deepEqual(freshWorkbook.getSelection(), { sheetName: "Sheet1", address: "A2" });
});
