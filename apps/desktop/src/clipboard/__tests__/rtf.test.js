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

test("clipboard RTF serializes named colors into a color table", () => {
  const rtf = serializeCellGridToRtf([
    [{ value: "X", format: { textColor: "red", backgroundColor: "yellow" } }],
  ]);

  assert.match(rtf, /\\colortbl;/);
  assert.match(rtf, /\\red255\\green0\\blue0;/);
  assert.match(rtf, /\\red255\\green255\\blue0;/);
  assert.match(rtf, /\\cf1\b/);
  assert.match(rtf, /\\clcbpat2\b/);
});

test("clipboard RTF supports OOXML-style ARGB hex colors", () => {
  const rtf = serializeCellGridToRtf([
    [{ value: "X", format: { textColor: "#FF112233", backgroundColor: "80112233" } }],
  ]);

  // #FF112233 -> rgb(17,34,51)
  assert.match(rtf, /\\red17\\green34\\blue51;/);

  // 0x80 alpha should be blended with white; result should still contain some color entry.
  assert.match(rtf, /\\colortbl;/);
});

test("clipboard RTF escapes unicode using \\\\uN? sequences", () => {
  // Astral-plane character (surrogate pair) + BMP character.
  const rtf = serializeCellGridToRtf([[{ value: "😀Ω" }]]);

  // 😀 is two UTF-16 code units, so it should emit *two* \\u escapes.
  assert.ok((rtf.match(/\\u-?\d+\?/g) ?? []).length >= 3);
});

test("clipboard RTF escapes special characters and newlines", () => {
  const rtf = serializeCellGridToRtf([[{ value: "a\\b{c}\nd\t" }]]);

  // Backslash, braces, newline->\line, tab->\tab
  assert.match(rtf, /a\\\\b\\\{c\\\}\\line d\\tab\b/);
});

test("serializeCellGridToClipboardPayload includes rtf", () => {
  const payload = serializeCellGridToClipboardPayload([[{ value: "X" }]]);
  assert.equal(typeof payload.rtf, "string");
  assert.ok(payload.rtf.startsWith("{\\rtf1"));
});
