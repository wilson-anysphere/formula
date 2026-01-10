import assert from "node:assert/strict";
import test from "node:test";

import { Workbook } from "../src/workbook.js";
import { VbaMigrator } from "../src/converter.js";
import { validateMigration } from "../src/validator.js";

class MockLlmClient {
  async complete({ prompt }) {
    if (/to Python/i.test(prompt)) {
      return [
        "```python",
        "sheet = formula.active_sheet",
        'sheet.Range("A1").Value = 1',
        "sheet.Cells(1, 2).Value = 2",
        'sheet.Range("A3").Formula = "=A1+B1"',
        "```"
      ].join("\n");
    }

    return [
      "```typescript",
      "const sheet = ctx.activeSheet;",
      'sheet.Range("A1").Value = 1;',
      "sheet.Cells(1, 2).Value = 2;",
      'sheet.Range("A3").Formula = "=A1+B1";',
      "```"
    ].join("\n");
  }
}

test("end-to-end: convert + validate a simple macro fixture against resulting cell diffs", async () => {
  const workbook = new Workbook();
  workbook.addSheet("Sheet1", { makeActive: true });

  const module = {
    name: "Module1",
    code: `
Sub Main()
    Range("A1").Value = 1
    Cells(1, 2).Value = 2
    Range("A3").Formula = "=A1+B1"
End Sub
`
  };

  const migrator = new VbaMigrator({ llm: new MockLlmClient() });

  const python = await migrator.convertModule(module, { target: "python" });
  const pythonValidation = validateMigration({
    workbook,
    module,
    entryPoint: "Main",
    target: "python",
    code: python.code
  });

  assert.equal(pythonValidation.ok, true);
  assert.equal(pythonValidation.mismatches.length, 0);
  assert.equal(pythonValidation.vbaDiff.length, 3);
  assert.equal(pythonValidation.scriptDiff.length, 3);

  const ts = await migrator.convertModule(module, { target: "typescript" });
  const tsValidation = validateMigration({
    workbook,
    module,
    entryPoint: "Main",
    target: "typescript",
    code: ts.code
  });

  assert.equal(tsValidation.ok, true);
  assert.equal(tsValidation.mismatches.length, 0);
  assert.equal(tsValidation.vbaDiff.length, 3);
  assert.equal(tsValidation.scriptDiff.length, 3);
});
