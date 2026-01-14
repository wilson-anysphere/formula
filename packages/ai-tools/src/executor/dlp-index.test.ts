import { describe, expect, it } from "vitest";

import { ToolExecutor } from "./tool-executor.ts";
import { parseA1Cell } from "../spreadsheet/a1.ts";
import { InMemoryWorkbook } from "../spreadsheet/in-memory-workbook.ts";

import { DLP_ACTION } from "../../../security/dlp/src/actions.js";

function instrumentRecordList(records: any[]) {
  let passes = 0;
  let elementGets = 0;
  const proxy = new Proxy(records, {
    get(target, prop, receiver) {
      if (prop === Symbol.iterator) {
        return function () {
          passes += 1;
          // Bind iterator to proxy so numeric index access is observable.
          return Array.prototype[Symbol.iterator].call(receiver);
        };
      }
      if (typeof prop === "string" && /^[0-9]+$/.test(prop)) {
        elementGets += 1;
      }
      return Reflect.get(target, prop, receiver);
    }
  });
  return { proxy, getPasses: () => passes, getElementGets: () => elementGets };
}

describe("ToolExecutor DLP indexing", () => {
  it("read_range does not scan all classification records per cell when indexed selectors are used", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: "ok" });
    workbook.setCell(parseA1Cell("Sheet1!B1"), { value: "secret" });

    const documentId = "doc-1";
    const sheetId = "Sheet1";

    // A single Restricted cell selector is sufficient to trigger a REDACT decision
    // for the selection. If ToolExecutor regresses to scanning `classification_records`
    // per cell, we'd see thousands of record passes over this list.
    const { proxy: classification_records, getPasses, getElementGets } = instrumentRecordList([
      {
        selector: { scope: "cell", documentId, sheetId, row: 0, col: 1 }, // B1
        classification: { level: "Restricted", labels: [] }
      }
    ]);

    const executor = new ToolExecutor(workbook, {
      // Allow reading the 10k-cell range for this test.
      max_read_range_cells: 20_000,
      dlp: {
        document_id: documentId,
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

    // Perf proxy: expect only a small number of linear scans (selection classification + index build).
    // Any per-cell scan regression would exceed this by orders of magnitude.
    expect(getPasses()).toBeLessThan(50);
    expect(getElementGets()).toBeLessThan(200);
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
  });
});
