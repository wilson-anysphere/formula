import test from "node:test";
import assert from "node:assert/strict";

import { parseClipboardContentToCellGrid } from "../clipboard.js";
import { parseHtmlToCellGrid } from "../html.js";
import { parseTsvToCellGrid } from "../tsv.js";

test("clipboard TSV parser enforces maxCells before materializing a huge grid", () => {
  assert.throws(
    () => parseTsvToCellGrid("1\t2\t3\t4\t5\t6", { maxCells: 5 }),
    (err) => err?.name === "ClipboardParseLimitError"
  );
});

test("clipboard HTML parser enforces maxCells before materializing a huge grid", () => {
  const html = `<!DOCTYPE html><html><body><table>
    <tr><td>1</td><td>2</td><td>3</td></tr>
    <tr><td>4</td><td>5</td><td>6</td></tr>
  </table></body></html>`;

  assert.throws(
    () => parseHtmlToCellGrid(html, { maxCells: 5 }),
    (err) => err?.name === "ClipboardParseLimitError"
  );
});

test("parseClipboardContentToCellGrid falls back to TSV when HTML exceeds maxChars", () => {
  const html = "<table><tr><td>1</td></tr></table>";
  const grid = parseClipboardContentToCellGrid({ html, text: "1\t2" }, { maxChars: 10 });
  assert.ok(grid);
  assert.equal(grid.length, 1);
  assert.equal(grid[0][0].value, 1);
  assert.equal(grid[0][1].value, 2);
});

test("parseClipboardContentToCellGrid does not treat empty text as a 1x1 clear when HTML is rejected for size", () => {
  const html = "<table><tr><td>1</td></tr></table>";
  assert.throws(
    () => parseClipboardContentToCellGrid({ html, text: "" }, { maxChars: 10 }),
    (err) => err?.name === "ClipboardParseLimitError"
  );
});

