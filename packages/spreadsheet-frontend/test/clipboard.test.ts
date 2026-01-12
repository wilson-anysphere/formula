import { describe, expect, it } from "vitest";
import { parseHtmlTableToGrid, parseTsvToGrid, serializeGridToHtmlTable, serializeGridToTsv } from "../src/index.js";

describe("clipboard helpers", () => {
  it("serializes a grid to TSV without extra delimiters", () => {
    expect(
      serializeGridToTsv([
        ["1", "2"],
        ["3", "4"]
      ])
    ).toBe("1\t2\n3\t4");
  });

  it("parses TSV payloads with CRLF and trailing newline", () => {
    expect(parseTsvToGrid("1\t2\r\n3\t4\r\n")).toEqual([
      ["1", "2"],
      ["3", "4"]
    ]);
  });

  it("caps TSV parsing when maxCells is exceeded", () => {
    const tsv = ["a", "b", "c", "d", "e", "f"].join("\t");
    expect(parseTsvToGrid(tsv, { maxCells: 5 })).toBe(null);
  });

  it("serializes a grid to an HTML table fragment", () => {
    const html = serializeGridToHtmlTable([
      ["1", "2"],
      ["3", "4"]
    ]);

    expect(html).toContain("<!--StartFragment-->");
    expect(html).toContain("<table>");
    expect(html).toContain("<td>1</td>");
    expect(html).toContain("<td>4</td>");
  });

  it("parses HTML table payloads (including <br> newlines + entities)", () => {
    const grid = parseHtmlTableToGrid(
      "<table><tr><td>hello<br>world</td><td>&lt;tag&gt;</td></tr><tr><td>&nbsp;</td><td>2</td></tr></table>"
    );
    expect(grid).toEqual([
      ["hello\nworld", "<tag>"],
      [" ", "2"]
    ]);
  });

  it("round-trips HTML table serialization back into a grid", () => {
    const original = [
      ["hello\nworld", "<tag>"],
      ["", "2"]
    ];
    const html = serializeGridToHtmlTable(original);
    expect(parseHtmlTableToGrid(html)).toEqual(original);
  });
});
