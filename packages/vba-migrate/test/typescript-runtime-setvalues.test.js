import assert from "node:assert/strict";
import test from "node:test";

import { Workbook } from "../src/workbook.js";
import { executeTypeScriptMigrationScript } from "../src/runtime/typescript.js";

test("executeTypeScriptMigrationScript supports getRange().setValues + setFormulas matrices", () => {
  const workbook = new Workbook();
  workbook.addSheet("Sheet1", { makeActive: true });

  const code = `
export default async function main(ctx) {
  const sheet = ctx.activeSheet;
  await sheet.getRange("A1:B1").setValues([[1, 2]]);
  await sheet.getRange("C1").setFormulas([["=A1+B1"]]);
}
`;

  executeTypeScriptMigrationScript({ workbook, code });

  const sheet = workbook.getSheet("Sheet1");
  assert.equal(sheet.getCell("A1").value, 1);
  assert.equal(sheet.getCell("B1").value, 2);
  assert.equal(sheet.getCell("C1").formula, "=A1+B1");
});

test("executeTypeScriptMigrationScript resolves const bindings + Array.from fill helpers", () => {
  const workbook = new Workbook();
  workbook.addSheet("Sheet1", { makeActive: true });

  const code = `
export default async function main(ctx) {
  const sheet = ctx.activeSheet;
  const fill = 7;
  const values = Array.from({ length: 2 }, () => Array(2).fill(fill));
  await sheet.getRange("$A$1:$B$2").setValues(values);

  const formula = "=A1+B1";
  await sheet.getRange("C1:D1").setFormulas(Array.from({ length: 1 }, () => Array(2).fill(formula)));
}
`;

  executeTypeScriptMigrationScript({ workbook, code });

  const sheet = workbook.getSheet("Sheet1");
  assert.equal(sheet.getCell("$A$1").value, 7);
  assert.equal(sheet.getCell("B2").value, 7);
  assert.equal(sheet.getCell("C1").formula, "=A1+B1");
  assert.equal(sheet.getCell("D1").formula, "=A1+B1");
});
