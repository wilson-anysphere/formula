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
  "macro recorder normalizes DocumentController formulas to include '=' prefix",
  { skip: typescriptAvailable ? false : "typescript not installed" },
  async () => {
  const controller = new DocumentController();
  const workbook = new DocumentControllerWorkbookAdapter(controller, { activeSheetName: "Sheet1" });

  const recorder = new MacroRecorder(workbook);
  recorder.start();

  workbook.setSelection("Sheet1", "A1");
  // DocumentController accepts formula expressions without a leading "=" and canonicalizes them.
  controller.setCellFormula("Sheet1", "A1", "A1+B1");

  const actions = recorder.stop();
  assert.deepEqual(actions, [
    { type: "setSelection", sheetName: "Sheet1", address: "A1" },
    { type: "setCellFormula", sheetName: "Sheet1", address: "A1", formula: "=A1+B1" },
  ]);

  const script = generateTypeScriptMacro(actions);

  const freshController = new DocumentController();
  const freshWorkbook = new DocumentControllerWorkbookAdapter(freshController, { activeSheetName: "Sheet1" });

  const runtime = new ScriptRuntime(freshWorkbook);
  const result = await runtime.run(script, { timeoutMs: 30_000 });
  assert.equal(result.error, undefined, result.error?.message);

  const cell = freshController.getCell("Sheet1", "A1");
  assert.equal(cell.formula, "=A1+B1");
  assert.equal(cell.value, null);

  workbook.dispose();
  freshWorkbook.dispose();
},
);
