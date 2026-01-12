import assert from "node:assert/strict";
import test from "node:test";

import { MacroRecorder, generateTypeScriptMacro } from "../apps/desktop/src/macro-recorder/index.js";
import { ScriptRuntime, Workbook } from "../packages/scripting/src/node.js";

let typescriptAvailable = true;
try {
  await import("typescript");
} catch {
  typescriptAvailable = false;
}

test(
  "macro recorder generates runnable TypeScript that replays formula edits (Workbook)",
  { skip: typescriptAvailable ? false : "typescript not installed" },
  async () => {
  const workbook = new Workbook();
  workbook.addSheet("Sheet1");
  workbook.setActiveSheet("Sheet1");

  const recorder = new MacroRecorder(workbook);
  recorder.start();

  workbook.setSelection("Sheet1", "A1");
  const sheet = workbook.getActiveSheet();
  sheet.setCellValue("A1", 10);
  sheet.setCellFormula("A2", "=A1*2");
  sheet.setCellValue("B1", 32);
  sheet.getRange("A1:B1").setFormat({ bold: true });
  sheet.setRangeValues("C1:D2", [
    [1, 2],
    [3, 4],
  ]);
  workbook.setSelection("Sheet1", "A2");

  const actions = recorder.stop();
  assert.ok(
    actions.some((a) => a.type === "setCellFormula" && a.sheetName === "Sheet1" && a.address === "A2" && a.formula === "=A1*2"),
    "expected recorder to capture setCellFormula(A2, =A1*2)",
  );
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
  assert.deepEqual(freshSheet.getRange("B1").getFormat(), { bold: true });
  assert.deepEqual(freshSheet.getRange("A2").getFormulas(), [["=A1*2"]]);
  assert.deepEqual(freshSheet.getRange("A2").getValues(), [[null]]);
  assert.deepEqual(
    freshSheet.getRange("C1:D2").getValues(),
    [
      [1, 2],
      [3, 4],
    ],
  );
  assert.deepEqual(freshWorkbook.getSelection(), { sheetName: "Sheet1", address: "A2" });
},
);
