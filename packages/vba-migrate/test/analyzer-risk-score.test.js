import assert from "node:assert/strict";
import test from "node:test";

import { analyzeVbaModule } from "../src/analyzer.js";

test("analyzeVbaModule assigns high risk scores to unsafe dynamic execution patterns", () => {
  const module = {
    name: "ModuleDanger",
    code: `
Sub Dangerous()
  Dim x
  x = Evaluate("1+1")
  Execute "MsgBox 1"
  CallByName Application, "Run", VbMethod, "Foo"
  Set fso = CreateObject("Scripting.FileSystemObject")
End Sub
`.trim(),
  };

  const report = analyzeVbaModule(module);
  assert.equal(report.moduleName, "ModuleDanger");
  assert.ok(report.unsafeConstructs.length >= 3);
  assert.ok(report.risk.score >= 70);
  assert.equal(report.risk.level, "high");

  const factorCodes = new Set(report.risk.factors.map((f) => f.code));
  assert.ok(factorCodes.has("unsafe_dynamic_execution"));
  assert.ok(factorCodes.has("external_dependency"));
});

