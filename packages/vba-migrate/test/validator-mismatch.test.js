import assert from "node:assert/strict";
import test from "node:test";

import { Workbook } from "../src/workbook.js";
import { validateMigration } from "../src/validator.js";
import { MockOracle } from "../src/vba/oracle.js";

test("validator reports actionable mismatches when migrated script diverges from VBA oracle", async () => {
  const workbook = new Workbook();
  workbook.addSheet("Sheet1", { makeActive: true });

  const module = {
    name: "Module1",
    code: `
Sub Main()
  Range("A1").Value = 1
End Sub
`.trim(),
  };

  const pythonCode = `
import formula

def main():
    sheet = formula.active_sheet
    sheet["A1"] = 2
`;

  const result = await validateMigration({
    workbook,
    module,
    entryPoint: "Main",
    target: "python",
    code: pythonCode,
    oracle: new MockOracle(),
  });

  assert.equal(result.ok, false);
  assert.equal(result.mismatches.length, 1);
  assert.deepEqual(result.mismatches[0], {
    sheet: "Sheet1",
    address: "A1",
    expected: { value: 1, formula: null, format: null },
    actual: { value: 2, formula: null, format: null },
  });
});

