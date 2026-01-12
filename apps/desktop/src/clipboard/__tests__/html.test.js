import test from "node:test";
import assert from "node:assert/strict";

import { parseHtmlToCellGrid, serializeCellGridToHtml } from "../html.js";

function buildCfHtmlPayload(innerHtml) {
  const markerStart = "<!--StartFragment-->";
  const markerEnd = "<!--EndFragment-->";
  const html = `<!DOCTYPE html><html><body>${markerStart}${innerHtml}${markerEnd}</body></html>`;

  const pad8 = (n) => String(n).padStart(8, "0");

  // Use fixed-width offset placeholders so the header length stays constant after substitution.
  const headerTemplate = [
    "Version:0.9",
    "StartHTML:00000000",
    "EndHTML:00000000",
    "StartFragment:00000000",
    "EndFragment:00000000",
    "SourceURL:https://example.com",
    "",
  ].join("\r\n");

  const startHTML = headerTemplate.length;
  const endHTML = startHTML + html.length;
  const startFragment = startHTML + html.indexOf(markerStart) + markerStart.length;
  const endFragment = startHTML + html.indexOf(markerEnd);

  const header = headerTemplate
    .replace("StartHTML:00000000", `StartHTML:${pad8(startHTML)}`)
    .replace("EndHTML:00000000", `EndHTML:${pad8(endHTML)}`)
    .replace("StartFragment:00000000", `StartFragment:${pad8(startFragment)}`)
    .replace("EndFragment:00000000", `EndFragment:${pad8(endFragment)}`);

  return header + html;
}

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

test("clipboard HTML parses Windows CF_HTML payloads", () => {
  const cfHtml = buildCfHtmlPayload("<table><tr><td>1</td><td>two</td></tr></table>");

  const grid = parseHtmlToCellGrid(cfHtml);
  assert.ok(grid);

  assert.equal(grid.length, 1);
  assert.equal(grid[0].length, 2);
  assert.equal(grid[0][0].value, 1);
  assert.equal(grid[0][1].value, "two");
});

test("clipboard HTML tolerates CF_HTML payloads with incorrect offsets", () => {
  const cfHtml = buildCfHtmlPayload("<table><tr><td>3</td><td>4</td></tr></table>")
    .replace(/StartHTML:\d{8}/, "StartHTML:00000010")
    .replace(/EndHTML:\d{8}/, "EndHTML:00000020")
    .replace(/StartFragment:\d{8}/, "StartFragment:00000010")
    .replace(/EndFragment:\d{8}/, "EndFragment:00000020");

  const grid = parseHtmlToCellGrid(cfHtml);
  assert.ok(grid);

  assert.equal(grid[0][0].value, 3);
  assert.equal(grid[0][1].value, 4);
});
