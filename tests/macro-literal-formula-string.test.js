import assert from "node:assert/strict";
import test from "node:test";

import { DocumentController } from "../apps/desktop/src/document/documentController.js";
import { MacroRecorder, generateTypeScriptMacro } from "../apps/desktop/src/macro-recorder/index.js";
import { DocumentControllerWorkbookAdapter } from "../apps/desktop/src/scripting/documentControllerWorkbookAdapter.js";
import { ScriptRuntime } from "../packages/scripting/src/node.js";

test("macro recorder preserves literal strings that start with '=' (DocumentController)", async () => {
  const controller = new DocumentController();
  const workbook = new DocumentControllerWorkbookAdapter(controller, { activeSheetName: "Sheet1" });

  const recorder = new MacroRecorder(workbook);
  recorder.start();

  workbook.setSelection("Sheet1", "A1");
  const sheet = workbook.getActiveSheet();
  sheet.getRange("A1").setValue("'=A1+B1");

  const actions = recorder.stop();
  assert.deepEqual(actions, [
    { type: "setSelection", sheetName: "Sheet1", address: "A1" },
    { type: "setCellValue", sheetName: "Sheet1", address: "A1", value: "'=A1+B1" },
  ]);

  const script = generateTypeScriptMacro(actions);

  const freshController = new DocumentController();
  const freshWorkbook = new DocumentControllerWorkbookAdapter(freshController, { activeSheetName: "Sheet1" });

  const runtime = new ScriptRuntime(freshWorkbook);
  const result = await runtime.run(script, { timeoutMs: 30_000 });
  assert.equal(result.error, undefined, result.error?.message);

  const freshSheet = freshWorkbook.getActiveSheet();
  assert.equal(freshSheet.getRange("A1").getValue(), "'=A1+B1");

  const cell = freshController.getCell("Sheet1", "A1");
  assert.equal(cell.formula, null);
  assert.equal(cell.value, "=A1+B1");

  workbook.dispose();
  freshWorkbook.dispose();
});
