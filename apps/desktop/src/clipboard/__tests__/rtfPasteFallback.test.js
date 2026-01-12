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
  assert.equal(grid[0][0].value, "A");
  assert.equal(grid[0][1].value, "B");
  assert.equal(grid[1][0].value, "C");
  assert.equal(grid[1][1].value, "D");
});
