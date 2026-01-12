import test from "node:test";
import assert from "node:assert/strict";

import { parseClipboardContentToCellGrid } from "../clipboard.js";
import { serializeCellGridToRtf } from "../rtf.js";

test("clipboard falls back to RTF when HTML/plain text are missing", () => {
  const rtf =
    "{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Arial;}}\\viewkind4\\uc1\\pard A\\tab B\\par C\\tab D\\par}";

  const grid = parseClipboardContentToCellGrid({ rtf });
  assert.ok(grid);

  assert.equal(grid.length, 2);
  assert.equal(grid[0].length, 2);
  assert.equal(grid[1].length, 2);

  assert.equal(grid[0][0].value, "A");
  assert.equal(grid[0][1].value, "B");
  assert.equal(grid[1][0].value, "C");
  assert.equal(grid[1][1].value, "D");
});

test("clipboard falls back to RTF tables (\\\\cell/\\\\row) when HTML/plain text are missing", () => {
  const rtf = serializeCellGridToRtf([
    [{ value: "A" }, { value: "B" }],
    [{ value: "C" }, { value: "D" }],
  ]);

  const grid = parseClipboardContentToCellGrid({ rtf });
  assert.ok(grid);

  assert.equal(grid.length, 2);
  assert.equal(grid[0].length, 2);
  assert.equal(grid[1].length, 2);

  assert.equal(grid[0][0].value, "A");
  assert.equal(grid[0][1].value, "B");
  assert.equal(grid[1][0].value, "C");
  assert.equal(grid[1][1].value, "D");
});

test("clipboard RTF table fallback treats \\tab/\\line inside cells as whitespace (no phantom columns/rows)", () => {
  const rtf = serializeCellGridToRtf([[{ value: "A\tB\nC" }, { value: "D" }]]);

  const grid = parseClipboardContentToCellGrid({ rtf });
  assert.ok(grid);
  assert.equal(grid.length, 1);
  assert.equal(grid[0].length, 2);

  // `\tab`/`\line` inside table cells are treated as whitespace during TSV fallback parsing.
  assert.equal(grid[0][0].value, "A B C");
  assert.equal(grid[0][1].value, "D");
});

test("RTF fallback preserves literal leading spaces after control word delimiters", () => {
  // Two spaces after \pard: first is the control-word delimiter, second is literal content.
  const rtf = "{\\rtf1\\ansi\\deff0\\uc1\\pard  A\\par}";
  const grid = parseClipboardContentToCellGrid({ rtf });
  assert.ok(grid);
  assert.equal(grid[0][0].value, " A");
});

test("RTF fallback preserves trailing tab separators for empty last cells", () => {
  // No trailing \par here: we want to ensure we don't strip the final tab, which represents
  // an empty last cell when parsing as TSV.
  const rtf = "{\\rtf1\\ansi\\deff0\\uc1\\pard A\\tab }";
  const grid = parseClipboardContentToCellGrid({ rtf });
  assert.ok(grid);
  assert.equal(grid.length, 1);
  assert.equal(grid[0].length, 2);
  assert.equal(grid[0][0].value, "A");
  assert.equal(grid[0][1].value, null);
});

test("clipboard RTF fallback still runs when text/plain is present but empty", () => {
  const rtf =
    "{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Arial;}}\\viewkind4\\uc1\\pard A\\tab B\\par C\\tab D\\par}";

  const grid = parseClipboardContentToCellGrid({ text: "", rtf });
  assert.ok(grid);

  assert.equal(grid.length, 2);
  assert.equal(grid[0].length, 2);
  assert.equal(grid[1].length, 2);

  assert.equal(grid[0][0].value, "A");
  assert.equal(grid[0][1].value, "B");
  assert.equal(grid[1][0].value, "C");
  assert.equal(grid[1][1].value, "D");
});
