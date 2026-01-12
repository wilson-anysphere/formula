import { describe, expect, it } from "vitest";

import { mergeFormattingIntoSnapshot } from "../mergeFormattingIntoSnapshot.js";

describe("mergeFormattingIntoSnapshot", () => {
  it("merges formatting into an existing value cell", () => {
    const cells = [{ row: 0, col: 0, value: 123, formula: null, format: null }];
    const result = mergeFormattingIntoSnapshot({
      cells,
      formatting: {
        cellFormats: [{ row: 0, col: 0, format: { font: { bold: true } } }],
      },
    });

    expect(result.cells).toEqual([{ row: 0, col: 0, value: 123, formula: null, format: { font: { bold: true } } }]);
  });

  it("adds a format-only cell when it does not exist in the value snapshot", () => {
    const cells = [{ row: 0, col: 0, value: 123, formula: null, format: null }];
    const result = mergeFormattingIntoSnapshot({
      cells,
      formatting: {
        cellFormats: [{ row: 5, col: 7, format: { fill: { color: "#ff0000" } } }],
      },
    });

    expect(result.cells).toEqual([
      { row: 0, col: 0, value: 123, formula: null, format: null },
      { row: 5, col: 7, value: null, formula: null, format: { fill: { color: "#ff0000" } } },
    ]);
  });

  it("clamps cellFormats to bounds when requested", () => {
    const result = mergeFormattingIntoSnapshot({
      cells: [],
      formatting: {
        cellFormats: [
          { row: 0, col: 0, format: { font: { bold: true } } },
          { row: 1, col: 0, format: { font: { italic: true } } },
        ],
      },
      clampCellFormatsTo: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 },
    });

    expect(result.cells).toEqual([{ row: 0, col: 0, value: null, formula: null, format: { font: { bold: true } } }]);
  });

  it("preserves layered formatting fields", () => {
    const formatting = {
      defaultFormat: { font: { name: "Arial" } },
      rowFormats: [{ row: 1, format: { font: { bold: true } } }],
      colFormats: [{ col: 2, format: { wrap: true } }],
      formatRunsByCol: [{ col: 0, runs: [{ startRow: 0, endRowExclusive: 2, format: { fill: { color: "#fff" } } }] }],
      cellFormats: [],
    };

    const result = mergeFormattingIntoSnapshot({ cells: [], formatting });

    expect(result.defaultFormat).toEqual(formatting.defaultFormat);
    expect(result.rowFormats).toEqual(formatting.rowFormats);
    expect(result.colFormats).toEqual(formatting.colFormats);
    expect(result.formatRunsByCol).toEqual(formatting.formatRunsByCol);
  });
});

