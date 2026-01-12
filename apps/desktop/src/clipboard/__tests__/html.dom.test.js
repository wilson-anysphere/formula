import test from "node:test";
import assert from "node:assert/strict";

import { JSDOM } from "jsdom";

import { parseHtmlToCellGrid, serializeCellGridToHtml } from "../html.js";

function buildCfHtmlPayload(innerHtml) {
  const markerStart = "<!--StartFragment-->";
  const markerEnd = "<!--EndFragment-->";
  const html = `<!DOCTYPE html><html><body>${markerStart}${innerHtml}${markerEnd}</body></html>`;

  const pad8 = (n) => String(n).padStart(8, "0");

  const headerTemplate = [
    "Version:0.9",
    "StartHTML:00000000",
    "EndHTML:00000000",
    "StartFragment:00000000",
    "EndFragment:00000000",
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

test("clipboard HTML DOM parser preserves <br> line breaks", () => {
  const dom = new JSDOM("", { url: "http://localhost" });

  const prevDomParser = globalThis.DOMParser;
  globalThis.DOMParser = dom.window.DOMParser;

  try {
    const html = `<!DOCTYPE html><html><body><table><tr><td>Line1<br>Line2</td></tr></table></body></html>`;
    const grid = parseHtmlToCellGrid(html);
    assert.ok(grid);
    assert.equal(grid[0][0].value, "Line1\nLine2");
  } finally {
    if (prevDomParser === undefined) {
      // `delete` is required to make `typeof DOMParser === "undefined"` checks pass.
      delete globalThis.DOMParser;
    } else {
      globalThis.DOMParser = prevDomParser;
    }
  }
});

test("clipboard HTML DOM parser does not double-count whitespace newlines after <br>", () => {
  const dom = new JSDOM("", { url: "http://localhost" });

  const prevDomParser = globalThis.DOMParser;
  globalThis.DOMParser = dom.window.DOMParser;

  try {
    const html =
      "<!DOCTYPE html><html><body><table><tr><td>Line1<br>\nLine2</td></tr></table></body></html>";
    const grid = parseHtmlToCellGrid(html);
    assert.ok(grid);
    assert.equal(grid[0][0].value, "Line1\nLine2");
  } finally {
    if (prevDomParser === undefined) {
      delete globalThis.DOMParser;
    } else {
      globalThis.DOMParser = prevDomParser;
    }
  }
});

test("clipboard HTML DOM parser handles Windows CF_HTML payloads", () => {
  const dom = new JSDOM("", { url: "http://localhost" });

  const prevDomParser = globalThis.DOMParser;
  globalThis.DOMParser = dom.window.DOMParser;

  try {
    const cfHtml = buildCfHtmlPayload("<table><tr><td>1</td><td>two</td></tr></table>");
    const grid = parseHtmlToCellGrid(cfHtml);
    assert.ok(grid);
    assert.equal(grid[0][0].value, 1);
    assert.equal(grid[0][1].value, "two");
  } finally {
    if (prevDomParser === undefined) {
      delete globalThis.DOMParser;
    } else {
      globalThis.DOMParser = prevDomParser;
    }
  }
});

test("clipboard HTML DOM parser normalizes NBSP to spaces", () => {
  const dom = new JSDOM("", { url: "http://localhost" });

  const prevDomParser = globalThis.DOMParser;
  globalThis.DOMParser = dom.window.DOMParser;

  try {
    const html = `<!DOCTYPE html><html><body><table><tr><td>Hello&nbsp;world</td></tr></table></body></html>`;
    const grid = parseHtmlToCellGrid(html);
    assert.ok(grid);
    assert.equal(grid[0][0].value, "Hello world");
  } finally {
    if (prevDomParser === undefined) {
      delete globalThis.DOMParser;
    } else {
      globalThis.DOMParser = prevDomParser;
    }
  }
});

test("clipboard HTML DOM parser round-trips multiline content", () => {
  const dom = new JSDOM("", { url: "http://localhost" });

  const prevDomParser = globalThis.DOMParser;
  globalThis.DOMParser = dom.window.DOMParser;

  try {
    const html = serializeCellGridToHtml([[{ value: "Line1\nLine2" }]]);
    const grid = parseHtmlToCellGrid(html);
    assert.ok(grid);
    assert.equal(grid[0][0].value, "Line1\nLine2");
  } finally {
    if (prevDomParser === undefined) {
      delete globalThis.DOMParser;
    } else {
      globalThis.DOMParser = prevDomParser;
    }
  }
});
