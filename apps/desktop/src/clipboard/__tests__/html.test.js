import test from "node:test";
import assert from "node:assert/strict";

import { parseClipboardContentToCellGrid } from "../clipboard.js";
import { parseHtmlToCellGrid, serializeCellGridToHtml } from "../html.js";

function buildCfHtmlPayload(
  innerHtml,
  {
    beforeFragmentHtml = "",
    includeFragmentMarkers = true,
    offsets = "codeUnits", // "codeUnits" | "bytes"
    afterHtml = "",
    fragmentEnd = "innerHtml", // "innerHtml" | "htmlEnd"
  } = {}
) {
  const markerStart = "<!--StartFragment-->";
  const markerEnd = "<!--EndFragment-->";
  const prefix = `<!DOCTYPE html><html><body>${beforeFragmentHtml}${includeFragmentMarkers ? markerStart : ""}`;
  const suffix = `${includeFragmentMarkers ? markerEnd : ""}</body></html>`;
  const html = `${prefix}${innerHtml}${suffix}${afterHtml}`;

  const pad8 = (n) => String(n).padStart(8, "0");

  const byteLength = (s) => {
    if (typeof TextEncoder !== "undefined") return new TextEncoder().encode(s).length;
    // eslint-disable-next-line no-undef
    return Buffer.from(s, "utf8").length;
  };

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

  const headerLen = offsets === "bytes" ? byteLength(headerTemplate) : headerTemplate.length;

  const startHTML = headerLen;
  const endHTML = startHTML + (offsets === "bytes" ? byteLength(html) : html.length);

  const startFragment =
    startHTML + (offsets === "bytes" ? byteLength(prefix) : prefix.length);
  const endFragment =
    fragmentEnd === "htmlEnd"
      ? endHTML
      : startFragment + (offsets === "bytes" ? byteLength(innerHtml) : innerHtml.length);

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

test("clipboard HTML serializes in-cell image values as alt text / placeholders (not [object Object])", () => {
  const htmlWithAlt = serializeCellGridToHtml([
    [{ value: { type: "image", value: { imageId: "img1", altText: " Alt " } } }],
  ]);
  assert.match(htmlWithAlt, /Alt/);

  const gridWithAlt = parseHtmlToCellGrid(htmlWithAlt);
  assert.ok(gridWithAlt);
  assert.equal(gridWithAlt[0][0].value, "Alt");

  const htmlWithoutAlt = serializeCellGridToHtml([
    [{ value: { type: "image", value: { imageId: "img1" } } }],
  ]);
  assert.match(htmlWithoutAlt, /\[Image\]/);

  const gridWithoutAlt = parseHtmlToCellGrid(htmlWithoutAlt);
  assert.ok(gridWithoutAlt);
  assert.equal(gridWithoutAlt[0][0].value, "[Image]");
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

test("clipboard HTML fallback parser does not double-count newlines after <br>", () => {
  const html = `<!DOCTYPE html><html><body><table><tr><td>Line1<br>
Line2</td></tr></table></body></html>`;

  const grid = parseHtmlToCellGrid(html);
  assert.ok(grid);
  assert.equal(grid[0][0].value, "Line1\nLine2");
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

test("clipboard HTML prefers CF_HTML fragment offsets when multiple tables exist", () => {
  const cfHtml = buildCfHtmlPayload("<table><tr><td>RIGHT</td></tr></table>", {
    beforeFragmentHtml: "<table><tr><td>WRONG</td></tr></table>",
  });

  const grid = parseHtmlToCellGrid(cfHtml);
  assert.ok(grid);

  assert.equal(grid[0][0].value, "RIGHT");
});

test("clipboard HTML tolerates incorrect CF_HTML offsets when multiple tables exist (uses fragment markers)", () => {
  const cfHtml = buildCfHtmlPayload("<table><tr><td>RIGHT</td></tr></table>", {
    beforeFragmentHtml: "<table><tr><td>WRONG</td></tr></table>",
  })
    .replace(/StartHTML:\d{8}/, "StartHTML:00000010")
    .replace(/EndHTML:\d{8}/, "EndHTML:00000020")
    .replace(/StartFragment:\d{8}/, "StartFragment:00000010")
    .replace(/EndFragment:\d{8}/, "EndFragment:00000020");

  const grid = parseHtmlToCellGrid(cfHtml);
  assert.ok(grid);

  assert.equal(grid[0][0].value, "RIGHT");
});

test("clipboard HTML uses fragment markers when fragment offsets are incorrect but StartHTML is valid", () => {
  // Some producers include both markers + StartHTML offsets but have incorrect StartFragment offsets.
  // Ensure we prefer the marker-delimited fragment over parsing the entire HTML (which might include
  // other tables before the intended fragment).
  const cfHtml = buildCfHtmlPayload("<table><tr><td>RIGHT</td></tr></table>", {
    beforeFragmentHtml: "<table><tr><td>WRONG</td></tr></table>",
  })
    .replace(/StartFragment:\d{8}/, "StartFragment:00000010")
    .replace(/EndFragment:\d{8}/, "EndFragment:00000020");

  const grid = parseHtmlToCellGrid(cfHtml);
  assert.ok(grid);
  assert.equal(grid[0][0].value, "RIGHT");
});

test("clipboard HTML uses StartFragment marker even when EndFragment marker is missing", () => {
  const cfHtml = buildCfHtmlPayload("<table><tr><td>RIGHT</td></tr></table>", {
    beforeFragmentHtml: "<table><tr><td>WRONG</td></tr></table>",
  })
    // Simulate malformed payloads that include StartFragment but not EndFragment.
    .replace(/<!--EndFragment-->/i, "")
    // Make fragment offsets unusable so we rely on marker fallback.
    .replace(/StartFragment:\d{8}/, "StartFragment:00000010")
    .replace(/EndFragment:\d{8}/, "EndFragment:00000020");

  const grid = parseHtmlToCellGrid(cfHtml);
  assert.ok(grid);
  assert.equal(grid[0][0].value, "RIGHT");
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

test("clipboard HTML tolerates CF_HTML payloads with truncated offsets (still containing '<table')", () => {
  let cfHtml = buildCfHtmlPayload("<table><tr><td>5</td><td>6</td></tr></table>");

  const getOffset = (name) => {
    const m = new RegExp(`${name}:(\\d{8})`).exec(cfHtml);
    assert.ok(m, `expected ${name} offset`);
    return Number.parseInt(m[1], 10);
  };

  const pad8 = (n) => String(n).padStart(8, "0");

  const startFragment = getOffset("StartFragment");
  const startHTML = getOffset("StartHTML");

  // Truncate the extracted slices so they include the opening <table> but not the closing tag.
  cfHtml = cfHtml
    .replace(/EndFragment:\d{8}/, `EndFragment:${pad8(startFragment + 20)}`)
    .replace(/EndHTML:\d{8}/, `EndHTML:${pad8(startHTML + 80)}`);

  const grid = parseHtmlToCellGrid(cfHtml);
  assert.ok(grid);

  assert.equal(grid[0][0].value, 5);
  assert.equal(grid[0][1].value, 6);
});

test("clipboard HTML handles CF_HTML byte offsets when non-ASCII content appears before the fragment", () => {
  // Include a bunch of multi-byte UTF-8 characters before the intended table, and omit fragment markers
  // so parsing must rely on StartFragment/EndFragment offsets.
  const cfHtml = buildCfHtmlPayload("<table><tr><td>RIGHT</td></tr></table>", {
    beforeFragmentHtml: `<table><tr><td>WRONG</td></tr></table>${"€".repeat(1000)}`,
    includeFragmentMarkers: false,
    offsets: "bytes",
  });

  const grid = parseHtmlToCellGrid(cfHtml);
  assert.ok(grid);

  assert.equal(grid[0][0].value, "RIGHT");
});

test("clipboard HTML strips leading/trailing NUL padding from clipboard payloads", () => {
  const cfHtml = buildCfHtmlPayload("<table><tr><td>Hello</td></tr></table>");
  const padded = `\u0000${cfHtml}\u0000\u0000`;

  const grid = parseHtmlToCellGrid(padded);
  assert.ok(grid);
  assert.equal(grid[0][0].value, "Hello");
});

test("clipboard HTML strips NUL padding from non-CF_HTML HTML payloads", () => {
  const html = `\u0000<!DOCTYPE html><html><body><table><tr><td>Hello</td></tr></table></body></html>\u0000`;
  const grid = parseHtmlToCellGrid(html);
  assert.ok(grid);
  assert.equal(grid[0][0].value, "Hello");
});

test("clipboard HTML honors CF_HTML offsets when payload has leading NUL padding", () => {
  // Omit fragment markers so parsing must rely on StartFragment/EndFragment offsets, and include a
  // preceding table so we can detect when offset parsing fails and we fall back to the wrong table.
  const base = buildCfHtmlPayload("<table><tr><td>RIGHT</td></tr></table>", {
    beforeFragmentHtml: "<table><tr><td>WRONG</td></tr></table>",
    includeFragmentMarkers: false,
    offsets: "bytes",
  });

  const pad8 = (n) => String(n).padStart(8, "0");
  const bump = (name, input) =>
    input.replace(new RegExp(`(${name}:)(\\d{8})`), (_, prefix, num) => {
      const next = Number.parseInt(num, 10) + 1;
      return `${prefix}${pad8(next)}`;
    });

  // Prefix with a NUL and update offsets to include it.
  let cfHtml = `\u0000${base}`;
  for (const name of ["StartHTML", "EndHTML", "StartFragment", "EndFragment"]) {
    cfHtml = bump(name, cfHtml);
  }

  const grid = parseHtmlToCellGrid(cfHtml);
  assert.ok(grid);
  assert.equal(grid[0][0].value, "RIGHT");
});

test("clipboard HTML tolerates leading NUL padding when CF_HTML offsets do not include it", () => {
  // Some clipboard bridges strip leading NUL bytes before computing CF_HTML offsets, while others
  // include them. Ensure we handle both by trying offset extraction against both the raw payload
  // and a version with leading NULs removed.
  const base = buildCfHtmlPayload("<table><tr><td>RIGHT</td></tr></table>", {
    beforeFragmentHtml: "<table><tr><td>WRONG</td></tr></table>",
    includeFragmentMarkers: false,
    offsets: "bytes",
  });

  const cfHtml = `\u0000${base}`;
  const grid = parseHtmlToCellGrid(cfHtml);
  assert.ok(grid);
  assert.equal(grid[0][0].value, "RIGHT");
});

test("parseClipboardContentToCellGrid does not trim CF_HTML payloads (offsets are relative to the original string)", () => {
  const cfHtml = buildCfHtmlPayload("<table><tr><td>RIGHT</td></tr></table>", {
    beforeFragmentHtml: "<table><tr><td>WRONG</td></tr></table>",
    includeFragmentMarkers: false,
    afterHtml: "\n\n", // would be stripped by .trim(), invalidating EndHTML/EndFragment.
    fragmentEnd: "htmlEnd",
  });

  const grid = parseClipboardContentToCellGrid({ html: cfHtml });
  assert.ok(grid);
  assert.equal(grid[0][0].value, "RIGHT");
});

test("clipboard HTML tolerates CF_HTML end offsets that include stripped trailing NUL padding", () => {
  const cfHtml = buildCfHtmlPayload("<table><tr><td>RIGHT</td></tr></table>", {
    beforeFragmentHtml: "<table><tr><td>WRONG</td></tr></table>",
    includeFragmentMarkers: false,
    offsets: "bytes",
    // Include trailing NUL padding *inside* the offset calculations (so EndHTML/EndFragment point past
    // the trimmed string). This simulates some native clipboard bridges.
    afterHtml: "\u0000\u0000\u0000",
    fragmentEnd: "htmlEnd",
  });

  const grid = parseHtmlToCellGrid(cfHtml);
  assert.ok(grid);
  assert.equal(grid[0][0].value, "RIGHT");
});
