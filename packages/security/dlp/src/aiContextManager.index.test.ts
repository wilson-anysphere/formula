import { describe, expect, it, vi } from "vitest";

// Mock the exact module specifier used by AiContextManager so we can spy on the slow path.
vi.mock("./selectors.js", async () => {
  const actual = await vi.importActual<any>("./selectors.js");
  return {
    ...actual,
    effectiveCellClassification: vi.fn(actual.effectiveCellClassification),
  };
});

import { AiContextManager } from "./aiContextManager.js";
import * as selectors from "./selectors.js";
import { DLP_ACTION } from "./actions.js";

describe("AiContextManager DLP indexing", () => {
  it("buildCloudContext does not call effectiveCellClassification per cell for indexed selectors", () => {
    // 100x100 = 10k cells/records.
    const classificationRecords = [];
    for (let row = 0; row < 100; row++) {
      for (let col = 0; col < 100; col++) {
        classificationRecords.push({
          selector: { scope: "cell", documentId: "doc-1", sheetId: "Sheet1", row, col },
          classification: { level: (row + col) % 2 === 0 ? "Public" : "Restricted", labels: [] },
        });
      }
    }

    const classificationStore = {
      list: () => classificationRecords,
    };
    const auditLogger = { log: vi.fn() };
    const manager = new AiContextManager({ classificationStore, auditLogger });

    const effectiveCellClassification = selectors.effectiveCellClassification as unknown as ReturnType<typeof vi.fn>;
    effectiveCellClassification.mockClear();

    const out = manager.buildCloudContext({
      documentId: "doc-1",
      sheetId: "Sheet1",
      range: { start: { row: 0, col: 0 }, end: { row: 99, col: 99 } },
      // Only a couple of cells have values; the rest are undefined, which is fine for redaction logic.
      cells: [
        { row: 0, col: 0, value: "ok" },
        { row: 0, col: 1, value: "secret" },
      ],
      policy: {
        version: 1,
        allowDocumentOverrides: true,
        rules: {
          [DLP_ACTION.AI_CLOUD_PROCESSING]: {
            maxAllowed: "Internal",
            allowRestrictedContent: false,
            redactDisallowed: true,
          },
        },
      },
    });

    expect(out.redactions.length).toBeGreaterThan(0);
    expect(out.context).toContain("[REDACTED]");

    // Perf proxy: we should not do an O(records) scan per cell.
    expect(effectiveCellClassification).toHaveBeenCalledTimes(0);

    expect(auditLogger.log).toHaveBeenCalledTimes(1);
    expect(auditLogger.log.mock.calls[0]?.[0]).toMatchObject({
      type: "ai.request",
      documentId: "doc-1",
      sheetId: "Sheet1",
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      redactedCellCount: out.redactions.length,
    });
  });
});

