import test from "node:test";
import assert from "node:assert/strict";

import { parseTsvToCellGrid, serializeCellGridToTsv } from "../tsv.js";

test("clipboard TSV parses numbers, formulas, and trailing newline", () => {
  const grid = parseTsvToCellGrid("1\t2\n3\t=SUM(A1:A2)\n");

  assert.equal(grid.length, 2);
  assert.equal(grid[0][0].value, 1);
  assert.equal(grid[0][1].value, 2);

  assert.equal(grid[1][0].value, 3);
  assert.equal(grid[1][1].formula, "=SUM(A1:A2)");
  assert.equal(grid[1][1].value, null);
});

test("clipboard TSV treats leading whitespace before '=' as a formula indicator", () => {
  const grid = parseTsvToCellGrid("  =SUM(A1:A2)");
  assert.equal(grid.length, 1);
  assert.equal(grid[0][0].formula, "=SUM(A1:A2)");
  assert.equal(grid[0][0].value, null);
});

test("clipboard TSV treats bare '=' as a formula input (cleared by downstream normalization)", () => {
  const grid = parseTsvToCellGrid("=");
  assert.equal(grid.length, 1);
  assert.equal(grid[0][0].formula, "=");
  assert.equal(grid[0][0].value, null);
});

test("clipboard TSV serializes to tab-separated lines", () => {
  const tsv = serializeCellGridToTsv([
    [{ value: "A" }, { value: 1 }],
    [{ value: true }, { value: null }],
  ]);

  assert.equal(tsv, "A\t1\ntrue\t");
});

test("clipboard TSV serializes formulas and escapes leading '='/' in strings", () => {
  const tsv = serializeCellGridToTsv([
    [{ value: null, formula: "=A1*2" }, { value: "=literal" }, { value: "'leading" }],
  ]);

  assert.equal(tsv, "=A1*2\t'=literal\t''leading");

  const grid = parseTsvToCellGrid(tsv);
  assert.equal(grid[0][0].formula, "=A1*2");
  assert.equal(grid[0][1].value, "=literal");
  assert.equal(grid[0][2].value, "'leading");
});

test("clipboard TSV escapes values that would otherwise look like formulas", () => {
  const tsv = serializeCellGridToTsv([[{ value: " =literal" }]]);
  assert.equal(tsv, "' =literal");

  const grid = parseTsvToCellGrid(tsv);
  assert.equal(grid[0][0].value, " =literal");
  assert.equal(grid[0][0].formula, null);
});

test("clipboard TSV serializes in-cell image values as alt text / placeholders (not [object Object])", () => {
  const tsvWithAlt = serializeCellGridToTsv([
    [{ value: { type: "image", value: { imageId: "img1", altText: " Alt " } } }],
  ]);
  assert.equal(tsvWithAlt, "Alt");

  const tsvWithoutAlt = serializeCellGridToTsv([
    [{ value: { type: "image", value: { imageId: "img1" } } }],
  ]);
  assert.equal(tsvWithoutAlt, "[Image]");
});
