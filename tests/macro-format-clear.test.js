import assert from "node:assert/strict";
import test from "node:test";

import { DocumentController } from "../apps/desktop/src/document/documentController.js";
import { MacroRecorder, generateTypeScriptMacro } from "../apps/desktop/src/macro-recorder/index.js";
import { DocumentControllerWorkbookAdapter } from "../apps/desktop/src/scripting/documentControllerWorkbookAdapter.js";
import { ScriptRuntime, Workbook } from "../packages/scripting/src/node.js";

test("macro recorder output can clear formatting (in-memory workbook)", async () => {
  const workbook = new Workbook();
  workbook.addSheet("Sheet1");
  workbook.setActiveSheet("Sheet1");

  const recorder = new MacroRecorder(workbook);
  recorder.start();

  const sheet = workbook.getActiveSheet();
  sheet.getRange("A1:B1").setFormat({ bold: true });
  sheet.getRange("A1:B1").setFormat(null);

  const script = generateTypeScriptMacro(recorder.stop());

  const freshWorkbook = new Workbook();
  freshWorkbook.addSheet("Sheet1");
  freshWorkbook.setActiveSheet("Sheet1");

  const runtime = new ScriptRuntime(freshWorkbook);
  const result = await runtime.run(script, { timeoutMs: 30_000 });
  assert.equal(result.error, undefined, result.error?.message);

  const freshSheet = freshWorkbook.getActiveSheet();
  assert.deepEqual(freshSheet.getRange("A1").getFormat(), {});
});

test("macro recorder output can replay unknown format keys (in-memory workbook)", async () => {
  const workbook = new Workbook();
  workbook.addSheet("Sheet1");
  workbook.setActiveSheet("Sheet1");

  const recorder = new MacroRecorder(workbook);
  recorder.start();

  const sheet = workbook.getActiveSheet();
  sheet.getRange("A1").setFormat({ foo: 123, bold: true });

  const script = generateTypeScriptMacro(recorder.stop());

  const freshWorkbook = new Workbook();
  freshWorkbook.addSheet("Sheet1");
  freshWorkbook.setActiveSheet("Sheet1");

  const runtime = new ScriptRuntime(freshWorkbook);
  const result = await runtime.run(script, { timeoutMs: 30_000 });
  assert.equal(result.error, undefined, result.error?.message);

  const freshSheet = freshWorkbook.getActiveSheet();
  assert.deepEqual(freshSheet.getRange("A1").getFormat(), { foo: 123, bold: true });
});

test("macro recorder output can clear formatting (DocumentController adapter)", async () => {
  const controller = new DocumentController();
  const workbook = new DocumentControllerWorkbookAdapter(controller, { activeSheetName: "Sheet1" });

  const recorder = new MacroRecorder(workbook);
  recorder.start();

  const sheet = workbook.getActiveSheet();
  sheet.getRange("A1:B1").setFormat({ bold: true });
  sheet.getRange("A1:B1").setFormat(null);

  const script = generateTypeScriptMacro(recorder.stop());

  const freshController = new DocumentController();
  const freshWorkbook = new DocumentControllerWorkbookAdapter(freshController, { activeSheetName: "Sheet1" });

  const runtime = new ScriptRuntime(freshWorkbook);
  const result = await runtime.run(script, { timeoutMs: 30_000 });
  assert.equal(result.error, undefined, result.error?.message);

  const freshSheet = freshWorkbook.getActiveSheet();
  assert.deepEqual(freshSheet.getRange("A1").getFormat(), {});

  workbook.dispose();
  freshWorkbook.dispose();
});
