import test from "node:test";
import assert from "node:assert/strict";

import { JSDOM } from "jsdom";

import { parseHtmlToCellGrid } from "../html.js";

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

