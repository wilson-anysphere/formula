import assert from "node:assert/strict";
import test from "node:test";

import { DocumentController } from "../apps/desktop/src/document/documentController.js";
import { MacroRecorder, generateTypeScriptMacro } from "../apps/desktop/src/macro-recorder/index.js";
import { DocumentControllerWorkbookAdapter } from "../apps/desktop/src/scripting/documentControllerWorkbookAdapter.js";
import { ScriptRuntime, Workbook } from "../packages/scripting/src/node.js";

test("macro recorder can record/replay clear formatting (toy workbook)", async () => {
  const workbook = new Workbook();
  workbook.addSheet("Sheet1");
  workbook.setActiveSheet("Sheet1");

  const recorder = new MacroRecorder(workbook);
  recorder.start();

  workbook.setSelection("Sheet1", "A1");
  const sheet = workbook.getActiveSheet();
  sheet.getRange("A1").setFormat({ bold: true });
  sheet.getRange("A1").setFormat(null);

  const actions = recorder.stop();
  assert.deepEqual(actions, [
    { type: "setSelection", sheetName: "Sheet1", address: "A1" },
    { type: "setFormat", sheetName: "Sheet1", address: "A1", format: { bold: true } },
    { type: "setFormat", sheetName: "Sheet1", address: "A1", format: null },
  ]);

  const script = generateTypeScriptMacro(actions);

  const freshWorkbook = new Workbook();
  freshWorkbook.addSheet("Sheet1");
  freshWorkbook.setActiveSheet("Sheet1");

  const runtime = new ScriptRuntime(freshWorkbook);
  const result = await runtime.run(script, { timeoutMs: 30_000 });
  assert.equal(result.error, undefined, result.error?.message);

  const freshSheet = freshWorkbook.getActiveSheet();
  assert.deepEqual(freshSheet.getRange("A1").getFormat(), {});
});

test("macro recorder can record/replay clear formatting (DocumentController)", async () => {
  const controller = new DocumentController();
  const workbook = new DocumentControllerWorkbookAdapter(controller, { activeSheetName: "Sheet1" });

  const recorder = new MacroRecorder(workbook);
  recorder.start();

  workbook.setSelection("Sheet1", "A1");
  const sheet = workbook.getActiveSheet();
  sheet.getRange("A1").setFormat({ bold: true });
  sheet.getRange("A1").setFormat(null);

  const actions = recorder.stop();
  assert.deepEqual(actions, [
    { type: "setSelection", sheetName: "Sheet1", address: "A1" },
    { type: "setFormat", sheetName: "Sheet1", address: "A1", format: { bold: true } },
    { type: "setFormat", sheetName: "Sheet1", address: "A1", format: null },
  ]);

  const script = generateTypeScriptMacro(actions);

  const freshController = new DocumentController();
  const freshWorkbook = new DocumentControllerWorkbookAdapter(freshController, { activeSheetName: "Sheet1" });

  const runtime = new ScriptRuntime(freshWorkbook);
  const result = await runtime.run(script, { timeoutMs: 30_000 });
  assert.equal(result.error, undefined, result.error?.message);

  const freshSheet = freshWorkbook.getActiveSheet();
  assert.deepEqual(freshSheet.getRange("A1").getFormat(), {});

  const cell = freshController.getCell("Sheet1", "A1");
  assert.equal(cell.styleId, 0);

  workbook.dispose();
  freshWorkbook.dispose();
});
