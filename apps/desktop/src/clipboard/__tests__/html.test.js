import test from "node:test";
import assert from "node:assert/strict";

import { parseHtmlToCellGrid, serializeCellGridToHtml } from "../html.js";

test("clipboard HTML round-trips basic values and formatting", () => {
  const html = serializeCellGridToHtml([
    [{ value: 1, format: { bold: true, textColor: "red" } }, { value: "Hello" }],
  ]);

  const grid = parseHtmlToCellGrid(html);
  assert.ok(grid);

  assert.equal(grid[0][0].value, 1);
  assert.equal(grid[0][0].format.bold, true);
  assert.equal(grid[0][1].value, "Hello");
});

test("clipboard HTML parses Google Sheets-style data attributes", () => {
  const html = `<!DOCTYPE html><html><body><table>
    <tr>
      <td data-sheets-value='{"1":3,"3":42}'>42</td>
      <td data-sheets-value='{"1":2,"2":"hello"}'>hello</td>
    </tr>
    <tr>
      <td data-sheets-formula="=A1*2">84</td>
      <td style="font-weight:bold;background-color:yellow">X</td>
    </tr>
  </table></body></html>`;

  const grid = parseHtmlToCellGrid(html);
  assert.ok(grid);

  assert.equal(grid[0][0].value, 42);
  assert.equal(grid[0][1].value, "hello");
  assert.equal(grid[1][0].formula, "=A1*2");
  assert.equal(grid[1][1].format.bold, true);
  assert.equal(grid[1][1].format.backgroundColor, "yellow");
});
