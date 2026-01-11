import assert from "node:assert/strict";
import test from "node:test";

import { VbaMigrator } from "../src/converter.js";

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

test("VbaMigrator converts + post-processes Python output into canonical Formula API calls", async () => {
  const migrator = new VbaMigrator({ llm: new MockLlmClient() });
  const module = {
    name: "Module1",
    code: `
Sub Main()
  Range("A1").Value = 1
End Sub
`
  };

  const result = await migrator.convertModule(module, { target: "python" });

  assert.match(result.code, /^import formula/m);
  assert.match(result.code, /\bdef main\(\):/);
  assert.match(result.code, /sheet = formula\.active_sheet/);
  assert.match(result.code, /sheet\["A1"\]\s*=\s*1/);
  assert.match(result.code, /sheet\["B1"\]\s*=\s*2/);
  assert.match(result.code, /sheet\["A3"\]\.formula\s*=\s*"=A1\+B1"/);
});

test("VbaMigrator converts + post-processes TypeScript output into canonical scripting API calls", async () => {
  const migrator = new VbaMigrator({ llm: new MockLlmClient() });
  const module = {
    name: "Module1",
    code: `
Sub Main()
  Range("A1").Value = 1
End Sub
`
  };

  const result = await migrator.convertModule(module, { target: "typescript" });

  assert.match(result.code, /\bexport default async function main\(ctx\)/);
  assert.match(result.code, /const sheet = ctx\.activeSheet/);
  assert.match(result.code, /await sheet\.getRange\("A1"\)\.setValue\(\s*1\s*\);/);
  assert.match(result.code, /await sheet\.getRange\("B1"\)\.setValue\(\s*2\s*\);/);
  assert.match(result.code, /await sheet\.getRange\("A3"\)\.setFormulas\(\s*\[\s*\[\s*"=A1\+B1"\s*\]\s*\]\s*\)\s*;/);
});
