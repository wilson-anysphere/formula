import assert from "node:assert/strict";
import test from "node:test";

import { DocumentController } from "../apps/desktop/src/document/documentController.js";
import { MacroRecorder, generateTypeScriptMacro } from "../apps/desktop/src/macro-recorder/index.js";
import { DocumentControllerWorkbookAdapter } from "../apps/desktop/src/scripting/documentControllerWorkbookAdapter.js";
import { ScriptRuntime } from "../packages/scripting/src/node.js";

let typescriptAvailable = true;
try {
  await import("typescript");
} catch {
  typescriptAvailable = false;
}

test(
  "macro recorder generates runnable TypeScript that replays formula edits against DocumentController",
  { skip: typescriptAvailable ? false : "typescript not installed" },
  async () => {
  const controller = new DocumentController();
  const workbook = new DocumentControllerWorkbookAdapter(controller, { activeSheetName: "Sheet1" });

  const recorder = new MacroRecorder(workbook);
  recorder.start();

  workbook.setSelection("Sheet1", "A1");
  const sheet = workbook.getActiveSheet();
  sheet.getRange("A1:B1").setValues([[1, 2]]);
  sheet.getRange("C1").setValue("=A1+B1");
  workbook.setSelection("Sheet1", "C1");

  const actions = recorder.stop();
  assert.deepEqual(actions, [
    { type: "setSelection", sheetName: "Sheet1", address: "A1" },
    { type: "setRangeValues", sheetName: "Sheet1", address: "A1:B1", values: [[1, 2]] },
    { type: "setCellFormula", sheetName: "Sheet1", address: "C1", formula: "=A1+B1" },
    { type: "setSelection", sheetName: "Sheet1", address: "C1" },
  ]);
  const script = generateTypeScriptMacro(actions);

  const freshController = new DocumentController();
  const freshWorkbook = new DocumentControllerWorkbookAdapter(freshController, { activeSheetName: "Sheet1" });

  const runtime = new ScriptRuntime(freshWorkbook);
  const result = await runtime.run(script, { timeoutMs: 30_000 });
  assert.equal(result.error, undefined, result.error?.message);

  const freshSheet = freshWorkbook.getActiveSheet();
  assert.deepEqual(freshSheet.getRange("A1:B1").getValues(), [[1, 2]]);
  assert.equal(freshSheet.getRange("C1").getValue(), "=A1+B1");
  assert.deepEqual(freshWorkbook.getSelection(), { sheetName: "Sheet1", address: "C1" });

  const cell = freshController.getCell("Sheet1", "C1");
  assert.equal(cell.formula, "=A1+B1");
  assert.equal(cell.value, null);

  workbook.dispose();
  freshWorkbook.dispose();
},
);
