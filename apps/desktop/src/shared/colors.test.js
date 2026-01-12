import assert from "node:assert/strict";
import test from "node:test";

import { normalizeExcelColorToCss } from "./colors.js";

test("normalizeExcelColorToCss converts Excel ARGB to canvas-safe CSS colors", () => {
  assert.equal(normalizeExcelColorToCss("#FF112233"), "#112233");

  // 0x80 / 255 = 0.50196..., rounded to 3 decimals for deterministic serialization.
  assert.equal(normalizeExcelColorToCss("#80112233"), "rgba(17,34,51,0.502)");

  assert.equal(normalizeExcelColorToCss("#fff"), "#ffffff");
  assert.equal(normalizeExcelColorToCss("112233"), "#112233");
  assert.equal(normalizeExcelColorToCss("rebeccapurple"), "rebeccapurple");

  assert.equal(normalizeExcelColorToCss("#GG112233"), undefined);
  assert.equal(normalizeExcelColorToCss("#12345"), undefined);
  assert.equal(normalizeExcelColorToCss(""), undefined);
});

test("normalizeExcelColorToCss resolves formula-model/XLSX color reference objects", () => {
  assert.equal(normalizeExcelColorToCss({ indexed: 2 }), "#ff0000");

  // Office 2013 default theme palette.
  assert.equal(normalizeExcelColorToCss({ theme: 4 }), "#5b9bd5"); // accent1
  assert.equal(normalizeExcelColorToCss({ theme: 0 }), "#ffffff"); // lt1

  // Tint values are thousandths, with negatives shading toward black.
  assert.equal(normalizeExcelColorToCss({ theme: 4, tint: -500 }), "#2e4e6b");

  assert.equal(normalizeExcelColorToCss({ auto: true }), undefined);
});
