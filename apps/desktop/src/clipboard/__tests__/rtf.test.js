import test from "node:test";
import assert from "node:assert/strict";

import { serializeCellGridToClipboardPayload } from "../clipboard.js";
import { serializeCellGridToRtf } from "../rtf.js";

test("clipboard RTF serializes a table with values", () => {
  const rtf = serializeCellGridToRtf([
    [{ value: "A1" }, { value: "B1" }],
    [{ value: null, formula: "=SUM(A1:B1)" }, { value: { text: "Rich", runs: [] } }],
  ]);

  assert.ok(rtf.startsWith("{\\rtf1"));
  assert.match(rtf, /A1/);
  assert.match(rtf, /B1/);
  assert.match(rtf, /=SUM\(A1:B1\)/);
  assert.match(rtf, /Rich/);
});

test("clipboard RTF includes basic formatting control words", () => {
  const rtf = serializeCellGridToRtf([
    [
      { value: "Bold", format: { bold: true } },
      { value: "Under", format: { underline: true } },
      { value: "Ital", format: { italic: true } },
    ],
  ]);

  assert.match(rtf, /\\b(?=\\|\s)/); // \b
  assert.match(rtf, /\\ul(?=\\|\s)/); // \ul
  assert.match(rtf, /\\i(?=\\|\s)/); // \i
});

test("serializeCellGridToClipboardPayload includes rtf", () => {
  const payload = serializeCellGridToClipboardPayload([[{ value: "X" }]]);
  assert.equal(typeof payload.rtf, "string");
  assert.ok(payload.rtf.startsWith("{\\rtf1"));
});
