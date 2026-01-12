import { describe, expect, it, vi } from "vitest";

// IMPORTANT: This test file lives next to ToolExecutor so we can mock the exact module specifier
// used by ToolExecutor ("../../../security/dlp/src/selectors.js").
vi.mock("../../../security/dlp/src/selectors.js", async () => {
  const actual = await vi.importActual<any>("../../../security/dlp/src/selectors.js");
  return {
    ...actual,
    effectiveCellClassification: vi.fn(actual.effectiveCellClassification)
  };
});

import { ToolExecutor } from "./tool-executor.ts";
import { parseA1Cell } from "../spreadsheet/a1.ts";
import { InMemoryWorkbook } from "../spreadsheet/in-memory-workbook.ts";

import { DLP_ACTION } from "../../../security/dlp/src/actions.js";
import * as selectors from "../../../security/dlp/src/selectors.js";

describe("ToolExecutor DLP indexing", () => {
  it("read_range does not scan all classification records per cell when indexed selectors are used", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: "ok" });
    workbook.setCell(parseA1Cell("Sheet1!B1"), { value: "secret" });

    // 100x100 = 10k cells/records.
    const classification_records = [];
    for (let row = 0; row < 100; row++) {
      for (let col = 0; col < 100; col++) {
        classification_records.push({
          selector: { scope: "cell", documentId: "doc-1", sheetId: "Sheet1", row, col },
          classification: { level: (row + col) % 2 === 0 ? "Public" : "Restricted", labels: [] }
        });
      }
    }

    const executor = new ToolExecutor(workbook, {
      // Allow reading the 10k-cell range for this test.
      max_read_range_cells: 20_000,
      dlp: {
        document_id: "doc-1",
        policy: {
          version: 1,
          allowDocumentOverrides: true,
          rules: {
            [DLP_ACTION.AI_CLOUD_PROCESSING]: {
              maxAllowed: "Internal",
              allowRestrictedContent: false,
              redactDisallowed: true
            }
          }
        },
        classification_records
      }
    });

    const effectiveCellClassification = selectors.effectiveCellClassification as unknown as ReturnType<typeof vi.fn>;
    effectiveCellClassification.mockClear();

    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "Sheet1!A1:CV100" }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");

    // Correctness (spot check).
    expect(result.data?.values[0]?.[0]).toBe("ok"); // A1 -> Public
    expect(result.data?.values[0]?.[1]).toBe("[REDACTED]"); // B1 -> Restricted

    // Perf proxy: with the DLP index, we should not call effectiveCellClassification per cell.
    expect(effectiveCellClassification).toHaveBeenCalledTimes(0);
  });

  it("read_range results match max-over-scopes semantics (document + sheet + column + range + cell)", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);

    // Populate A1:C3
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: "a1" });
    workbook.setCell(parseA1Cell("Sheet1!B1"), { value: "b1" });
    workbook.setCell(parseA1Cell("Sheet1!C1"), { value: "c1" });
    workbook.setCell(parseA1Cell("Sheet1!A2"), { value: "a2" });
    workbook.setCell(parseA1Cell("Sheet1!B2"), { value: "b2" });
    workbook.setCell(parseA1Cell("Sheet1!C2"), { value: "c2" });
    workbook.setCell(parseA1Cell("Sheet1!A3"), { value: "a3" });
    workbook.setCell(parseA1Cell("Sheet1!B3"), { value: "b3" });
    workbook.setCell(parseA1Cell("Sheet1!C3"), { value: "c3" });

    const executor = new ToolExecutor(workbook, {
      dlp: {
        document_id: "doc-1",
        policy: {
          version: 1,
          allowDocumentOverrides: true,
          rules: {
            [DLP_ACTION.AI_CLOUD_PROCESSING]: {
              maxAllowed: "Internal",
              allowRestrictedContent: false,
              redactDisallowed: true
            }
          }
        },
        classification_records: [
          {
            selector: { scope: "document", documentId: "doc-1" },
            classification: { level: "Internal", labels: ["doc"] }
          },
          {
            selector: { scope: "sheet", documentId: "doc-1", sheetId: "Sheet1" },
            classification: { level: "Internal", labels: ["sheet"] }
          },
          {
            selector: { scope: "column", documentId: "doc-1", sheetId: "Sheet1", columnIndex: 1 }, // B
            classification: { level: "Confidential", labels: ["col"] }
          },
          {
            selector: {
              scope: "range",
              documentId: "doc-1",
              sheetId: "Sheet1",
              range: { start: { row: 1, col: 0 }, end: { row: 1, col: 2 } } // A2:C2
            },
            classification: { level: "Restricted", labels: ["range"] }
          },
          {
            selector: { scope: "cell", documentId: "doc-1", sheetId: "Sheet1", row: 0, col: 0 }, // A1
            classification: { level: "Confidential", labels: ["cellA1"] }
          },
          {
            selector: { scope: "cell", documentId: "doc-1", sheetId: "Sheet1", row: 2, col: 2 }, // C3
            classification: { level: "Restricted", labels: ["cellC3"] }
          }
        ]
      }
    });

    const effectiveCellClassification = selectors.effectiveCellClassification as unknown as ReturnType<typeof vi.fn>;
    effectiveCellClassification.mockClear();

    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "Sheet1!A1:C3" }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");

    expect(result.data?.values).toEqual([
      ["[REDACTED]", "[REDACTED]", "c1"],
      ["[REDACTED]", "[REDACTED]", "[REDACTED]"],
      ["a3", "[REDACTED]", "[REDACTED]"]
    ]);

    // No fallback selectors -> should stay on the indexed path.
    expect(effectiveCellClassification).toHaveBeenCalledTimes(0);
  });
});
