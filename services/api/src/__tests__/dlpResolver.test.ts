import { describe, expect, it } from "vitest";
import { resolveClassification } from "../dlp/dlp";

describe("DLP classification resolver", () => {
  it("resolves max across cell/range/column/sheet/document and unions labels", () => {
    const documentId = "doc-1";
    const sheetId = "Sheet1";

    const records = [
      { selector: { scope: "document", documentId }, classification: { level: "Internal", labels: ["doc"] } },
      {
        selector: { scope: "sheet", documentId, sheetId },
        classification: { level: "Confidential", labels: ["sheet"] }
      },
      {
        selector: { scope: "column", documentId, sheetId, columnIndex: 1 },
        classification: { level: "Internal", labels: ["col"] }
      },
      {
        selector: {
          scope: "range",
          documentId,
          sheetId,
          range: { start: { row: 0, col: 0 }, end: { row: 2, col: 2 } }
        },
        classification: { level: "Restricted", labels: ["range", "pii"] }
      },
      {
        selector: { scope: "cell", documentId, sheetId, row: 1, col: 1 },
        classification: { level: "Confidential", labels: ["cell", "pii"] }
      }
    ];

    const resolved = resolveClassification({
      querySelector: { scope: "cell", documentId, sheetId, row: 1, col: 1 },
      records,
      options: { includeMatchedSelectors: true }
    });

    expect(resolved.effectiveClassification.level).toBe("Restricted");
    expect(resolved.effectiveClassification.labels).toEqual(["cell", "col", "doc", "pii", "range", "sheet"]);
    expect(resolved.matchedCount).toBe(5);
    expect(resolved.matchedSelectors?.map((m) => m.selector.scope)).toEqual(["cell", "range", "column", "sheet", "document"]);
  });

  it("treats overlapping ranges as intersecting for range queries", () => {
    const documentId = "doc-2";
    const sheetId = "Sheet1";

    const records = [
      {
        selector: { scope: "range", documentId, sheetId, range: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } } },
        classification: { level: "Confidential", labels: ["a"] }
      },
      {
        selector: { scope: "range", documentId, sheetId, range: { start: { row: 1, col: 1 }, end: { row: 2, col: 2 } } },
        classification: { level: "Restricted", labels: ["b"] }
      }
    ];

    const resolved = resolveClassification({
      querySelector: { scope: "range", documentId, sheetId, range: { start: { row: 0, col: 0 }, end: { row: 2, col: 2 } } },
      records
    });

    expect(resolved.effectiveClassification.level).toBe("Restricted");
    expect(resolved.effectiveClassification.labels).toEqual(["a", "b"]);
    expect(resolved.matchedCount).toBe(2);
  });

  it("unions labels even when a more restrictive selector dominates the level", () => {
    const documentId = "doc-3";
    const sheetId = "Sheet1";

    const records = [
      { selector: { scope: "document", documentId }, classification: { level: "Internal", labels: ["doc"] } },
      {
        selector: { scope: "range", documentId, sheetId, range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } } },
        classification: { level: "Restricted", labels: ["pii"] }
      },
      {
        selector: { scope: "cell", documentId, sheetId, row: 0, col: 0 },
        classification: { level: "Confidential", labels: ["finance"] }
      }
    ];

    const resolved = resolveClassification({
      querySelector: { scope: "cell", documentId, sheetId, row: 0, col: 0 },
      records
    });

    expect(resolved.effectiveClassification.level).toBe("Restricted");
    expect(resolved.effectiveClassification.labels).toEqual(["doc", "finance", "pii"]);
    expect(resolved.matchedCount).toBe(3);
  });

  it("rejects very large range queries when matched selectors are requested", () => {
    const documentId = "doc-4";
    const sheetId = "Sheet1";
    const records = [
      { selector: { scope: "document", documentId }, classification: { level: "Internal", labels: [] } }
    ];

    expect(() =>
      resolveClassification({
        querySelector: {
          scope: "range",
          documentId,
          sheetId,
          range: { start: { row: 0, col: 0 }, end: { row: 999, col: 1000 } } // 1,001,000 cells
        },
        records,
        options: { includeMatchedSelectors: true }
      })
    ).toThrow(/Range too large to include matched selectors/);
  });
});

