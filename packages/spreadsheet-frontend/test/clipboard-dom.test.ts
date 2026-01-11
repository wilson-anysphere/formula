// @vitest-environment jsdom

import { describe, expect, it } from "vitest";
import { parseHtmlTableToGrid } from "../src/index.js";

describe("clipboard helpers (DOM parser)", () => {
  it("parses HTML tables via DOMParser when available", () => {
    const grid = parseHtmlTableToGrid("<table><tr><td>1</td><td>2</td></tr><tr><td>3</td><td>4</td></tr></table>");
    expect(grid).toEqual([
      ["1", "2"],
      ["3", "4"]
    ]);
  });
});

