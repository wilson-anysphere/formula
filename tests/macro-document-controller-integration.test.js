import assert from "node:assert/strict";
import test from "node:test";

import { DocumentController } from "../apps/desktop/src/document/documentController.js";
import { MacroRecorder, generateTypeScriptMacro } from "../apps/desktop/src/macro-recorder/index.js";
import { DocumentControllerWorkbookAdapter } from "../apps/desktop/src/scripting/documentControllerWorkbookAdapter.js";
import { ScriptRuntime } from "../packages/scripting/src/node.js";

test("macro recorder generates runnable TypeScript that replays edits against DocumentController", async () => {
  const controller = new DocumentController();
  const workbook = new DocumentControllerWorkbookAdapter(controller, { activeSheetName: "Sheet1" });

  const recorder = new MacroRecorder(workbook);
  recorder.start();

  workbook.setSelection("Sheet1", "A1");
  const sheet = workbook.getActiveSheet();
  sheet.getRange("A1:B1").setValues([[10, 32]]);
  sheet.getRange("A1:B1").setFormat({ bold: true, backgroundColor: "#FFFFFF00", numberFormat: "0%" });
  workbook.setSelection("Sheet1", "A2");

  const actions = recorder.stop();
  assert.deepEqual(actions, [
    { type: "setSelection", sheetName: "Sheet1", address: "A1" },
    { type: "setRangeValues", sheetName: "Sheet1", address: "A1:B1", values: [[10, 32]] },
    {
      type: "setFormat",
      sheetName: "Sheet1",
      address: "A1:B1",
      format: { bold: true, backgroundColor: "#FFFFFF00", numberFormat: "0%" },
    },
    { type: "setSelection", sheetName: "Sheet1", address: "A2" },
  ]);
  const script = generateTypeScriptMacro(actions);

  const freshController = new DocumentController();
  const freshWorkbook = new DocumentControllerWorkbookAdapter(freshController, { activeSheetName: "Sheet1" });

  const runtime = new ScriptRuntime(freshWorkbook);
  const result = await runtime.run(script, { timeoutMs: 30_000 });
  assert.equal(result.error, undefined, result.error?.message);

  const freshSheet = freshWorkbook.getActiveSheet();
  assert.deepEqual(freshSheet.getRange("A1:B1").getValues(), [[10, 32]]);
  assert.deepEqual(freshSheet.getRange("A1").getFormat(), {
    bold: true,
    backgroundColor: "#FFFFFF00",
    numberFormat: "0%",
  });
  assert.deepEqual(freshSheet.getRange("B1").getFormat(), {
    bold: true,
    backgroundColor: "#FFFFFF00",
    numberFormat: "0%",
  });
  assert.deepEqual(freshWorkbook.getSelection(), { sheetName: "Sheet1", address: "A2" });

  const a1 = freshController.getCell("Sheet1", "A1");
  const b1 = freshController.getCell("Sheet1", "B1");
  const a1Style = freshController.styleTable.get(a1.styleId);
  const b1Style = freshController.styleTable.get(b1.styleId);
  assert.equal(a1Style.font?.bold, true);
  assert.equal(b1Style.font?.bold, true);
  assert.equal(a1Style.bold, undefined); // adapter should map to DocumentController schema
  assert.equal(a1Style.fill?.fgColor, "#FFFFFF00");
  assert.equal(b1Style.fill?.fgColor, "#FFFFFF00");
  assert.equal(a1Style.numberFormat, "0%");

  workbook.dispose();
  freshWorkbook.dispose();
});
