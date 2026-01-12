import { describe, expect, it, vi } from "vitest";

// Mock the exact module specifier used by ContextManager so we can spy on the slow path.
vi.mock("../../security/dlp/src/selectors.js", async () => {
  const actual = await vi.importActual<any>("../../security/dlp/src/selectors.js");
  return {
    ...actual,
    effectiveCellClassification: vi.fn(actual.effectiveCellClassification),
  };
});

import { ContextManager } from "./contextManager.js";
import * as selectors from "../../security/dlp/src/selectors.js";
import { DLP_ACTION } from "../../security/dlp/src/actions.js";

describe("ContextManager DLP indexing", () => {
  it("buildContext does not call effectiveCellClassification per cell for indexed selectors", async () => {
    const values: unknown[][] = Array.from({ length: 100 }, (_, r) =>
      Array.from({ length: 100 }, (_v, c) => (r === 0 && c === 0 ? "ok" : r === 0 && c === 1 ? "secret" : null)),
    );

    // 100x100 = 10k cell-level records.
    const classificationRecords = [];
    for (let row = 0; row < 100; row++) {
      for (let col = 0; col < 100; col++) {
        classificationRecords.push({
          selector: { scope: "cell", documentId: "doc-1", sheetId: "Sheet1", row, col },
          classification: { level: (row + col) % 2 === 0 ? "Public" : "Restricted", labels: [] },
        });
      }
    }

    const cm = new ContextManager({ tokenBudgetTokens: 1000 });
    const effectiveCellClassification = selectors.effectiveCellClassification as unknown as ReturnType<typeof vi.fn>;
    effectiveCellClassification.mockClear();

    const auditLogger = { log: vi.fn() };

    const out = await cm.buildContext({
      sheet: { name: "Sheet1", values },
      query: "test",
      dlp: {
        documentId: "doc-1",
        sheetId: "Sheet1",
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
        classificationRecords,
        auditLogger,
      },
    });

    expect(out.promptContext).toContain("[REDACTED]");

    // Perf proxy: ensure we used the indexed path rather than scanning all records per cell.
    expect(effectiveCellClassification).toHaveBeenCalledTimes(0);

    expect(auditLogger.log).toHaveBeenCalledTimes(1);
    expect(auditLogger.log.mock.calls[0]?.[0]).toMatchObject({
      type: "ai.context",
      documentId: "doc-1",
      sheetId: "Sheet1",
    });
  });

  it("buildContext resolves sheet display names to stable sheet ids for structured DLP enforcement", async () => {
    const documentId = "doc-1";
    const displayName = "Budget";
    const stableSheetId = "Sheet2";

    const cm = new ContextManager({ tokenBudgetTokens: 1000 });

    const auditLogger = { log: vi.fn() };

    const out = await cm.buildContext({
      sheet: { name: displayName, values: [["Public"], ["TOP SECRET"]] },
      query: "secret",
      dlp: {
        document_id: documentId,
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
        classification_records: [
          {
            selector: { scope: "cell", documentId, sheetId: stableSheetId, row: 1, col: 0 },
            classification: { level: "Restricted", labels: [] },
          },
        ],
        sheet_name_resolver: {
          getSheetIdByName: (name: string) => (name.trim().toLowerCase() === displayName.toLowerCase() ? stableSheetId : null),
        },
        auditLogger,
      },
    });

    expect(out.schema.name).toBe(displayName);
    expect(out.retrieved[0]?.range).toBe(`${displayName}!A1:A2`);
    expect(out.promptContext).toContain("[REDACTED]");
    expect(out.promptContext).not.toContain("TOP SECRET");

    expect(auditLogger.log).toHaveBeenCalledTimes(1);
    expect(auditLogger.log.mock.calls[0]?.[0]).toMatchObject({
      type: "ai.context",
      documentId,
      sheetId: stableSheetId,
      sheetName: displayName,
    });
  });
});
