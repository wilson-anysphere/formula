import assert from "node:assert/strict";
import test from "node:test";

import { analyzeVbaModule } from "../src/analyzer.js";

test("analyzeVbaModule detects object model usage, external references, and unsupported constructs", () => {
  const module = {
    name: "Module1",
    code: `
Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long

Sub Test()
    Worksheets("Sheet1").Range("A1").Value = 1
    Cells(1, 2).Value = 2
    On Error Resume Next
    GoTo Done
Done:
End Sub
`
  };

  const report = analyzeVbaModule(module);

  assert.equal(report.moduleName, "Module1");
  assert.equal(report.objectModelUsage.Range.length, 1);
  assert.equal(report.objectModelUsage.Cells.length, 1);
  assert.equal(report.objectModelUsage.Worksheets.length, 1);
  assert.equal(report.externalReferences.length, 1);
  assert.ok(report.warnings.some((w) => w.message.includes("error handling")));
  assert.ok(report.warnings.some((w) => w.message.includes("GoTo")));
  assert.ok(report.todos.length >= 1);
});

