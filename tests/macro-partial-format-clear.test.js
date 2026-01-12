import assert from "node:assert/strict";
import test from "node:test";

import { DocumentController } from "../apps/desktop/src/document/documentController.js";
import { MacroRecorder, generateTypeScriptMacro } from "../apps/desktop/src/macro-recorder/index.js";
import { DocumentControllerWorkbookAdapter } from "../apps/desktop/src/scripting/documentControllerWorkbookAdapter.js";
import { ScriptRuntime } from "../packages/scripting/src/node.js";

test("macro recorder can clear backgroundColor without clearing other formats (DocumentController)", async () => {
  const controller = new DocumentController();
  const workbook = new DocumentControllerWorkbookAdapter(controller, { activeSheetName: "Sheet1" });

  const recorder = new MacroRecorder(workbook);
  recorder.start();

  workbook.setSelection("Sheet1", "A1");
  const sheet = workbook.getActiveSheet();
  sheet.getRange("A1").setFormat({ bold: true, backgroundColor: "#FFFFFF00" });
  sheet.getRange("A1").setFormat({ backgroundColor: null });

  const actions = recorder.stop();
  assert.deepEqual(actions, [
    { type: "setSelection", sheetName: "Sheet1", address: "A1" },
    {
      type: "setFormat",
      sheetName: "Sheet1",
      address: "A1",
      format: { bold: true, backgroundColor: "#FFFFFF00" },
    },
    { type: "setFormat", sheetName: "Sheet1", address: "A1", format: { backgroundColor: null } },
  ]);

  const script = generateTypeScriptMacro(actions);

  const freshController = new DocumentController();
  const freshWorkbook = new DocumentControllerWorkbookAdapter(freshController, { activeSheetName: "Sheet1" });

  const runtime = new ScriptRuntime(freshWorkbook);
  const result = await runtime.run(script, { timeoutMs: 30_000 });
  assert.equal(result.error, undefined, result.error?.message);

  const freshSheet = freshWorkbook.getActiveSheet();
  assert.deepEqual(freshSheet.getRange("A1").getFormat(), { bold: true });

  const cell = freshController.getCell("Sheet1", "A1");
  const style = freshController.styleTable.get(cell.styleId);
  assert.equal(style.font?.bold, true);
  assert.equal(style.fill, null);

  workbook.dispose();
  freshWorkbook.dispose();
});

test("macro recorder can clear numberFormat without clearing other formats (DocumentController)", async () => {
  const controller = new DocumentController();
  const workbook = new DocumentControllerWorkbookAdapter(controller, { activeSheetName: "Sheet1" });

  const recorder = new MacroRecorder(workbook);
  recorder.start();

  workbook.setSelection("Sheet1", "A1");
  const sheet = workbook.getActiveSheet();
  sheet.getRange("A1").setFormat({ bold: true, numberFormat: "0%" });
  sheet.getRange("A1").setFormat({ numberFormat: null });

  const actions = recorder.stop();
  assert.deepEqual(actions, [
    { type: "setSelection", sheetName: "Sheet1", address: "A1" },
    { type: "setFormat", sheetName: "Sheet1", address: "A1", format: { bold: true, numberFormat: "0%" } },
    { type: "setFormat", sheetName: "Sheet1", address: "A1", format: { numberFormat: null } },
  ]);

  const script = generateTypeScriptMacro(actions);

  const freshController = new DocumentController();
  const freshWorkbook = new DocumentControllerWorkbookAdapter(freshController, { activeSheetName: "Sheet1" });

  const runtime = new ScriptRuntime(freshWorkbook);
  const result = await runtime.run(script, { timeoutMs: 30_000 });
  assert.equal(result.error, undefined, result.error?.message);

  const freshSheet = freshWorkbook.getActiveSheet();
  assert.deepEqual(freshSheet.getRange("A1").getFormat(), { bold: true });

  const cell = freshController.getCell("Sheet1", "A1");
  const style = freshController.styleTable.get(cell.styleId);
  assert.equal(style.font?.bold, true);
  assert.equal(style.numberFormat, null);

  workbook.dispose();
  freshWorkbook.dispose();
});
