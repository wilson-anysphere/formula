import assert from "node:assert/strict";
import test from "node:test";

import { valuesRangeToTsv } from "../src/tsv.js";

test("valuesRangeToTsv: formats rich text + in-cell images as plain text", () => {
  const values = [
    [
      { text: "Header", runs: [{ start: 0, end: 6, style: { bold: true } }] },
      { type: "image", value: { imageId: "img_1", altText: " Logo " } },
      { imageId: "img_2" },
      { t: "n", v: 42 },
      { t: "blank" },
      {},
    ],
  ];

  const out = valuesRangeToTsv(values, { startRow: 0, startCol: 0, endRow: 0, endCol: 5 }, { maxRows: 1 });
  assert.equal(out, "Header\tLogo\t[Image]\t42\t\t");
});
