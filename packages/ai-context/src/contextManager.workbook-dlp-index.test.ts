import { describe, expect, it, vi } from "vitest";

vi.mock("../../security/dlp/src/selectors.js", async () => {
  const actual = await vi.importActual<any>("../../security/dlp/src/selectors.js");
  return {
    ...actual,
    effectiveRangeClassification: vi.fn(actual.effectiveRangeClassification),
  };
});

import { ContextManager } from "./contextManager.js";
import { HashEmbedder, InMemoryVectorStore } from "../../ai-rag/src/index.js";
import { DLP_ACTION } from "../../security/dlp/src/actions.js";
import * as selectors from "../../security/dlp/src/selectors.js";

describe("ContextManager buildWorkbookContext DLP indexing", () => {
  it("avoids per-chunk effectiveRangeClassification scans for large structured record sets", async () => {
    const workbook = {
      id: "wb-dlp-index",
      sheets: [
        {
          name: "Sheet1",
          cells: [
            [{ v: "Header1" }, { v: "Header2" }],
            [{ v: "ok" }, { v: "secret" }],
          ],
        },
      ],
      tables: [],
    };

    // 100x100 = 10k cell-level records (structured).
    const classificationRecords = [];
    for (let row = 0; row < 100; row++) {
      for (let col = 0; col < 100; col++) {
        classificationRecords.push({
          selector: { scope: "cell", documentId: workbook.id, sheetId: "Sheet1", row, col },
          classification: { level: row === 1 && col === 1 ? "Restricted" : "Public", labels: [] },
        });
      }
    }

    const embedder = new HashEmbedder({ dimension: 32 });
    const vectorStore = new InMemoryVectorStore({ dimension: 32 });
    const cm = new ContextManager({ tokenBudgetTokens: 500, workbookRag: { vectorStore, embedder, topK: 3 } });

    const effectiveRangeClassification = selectors.effectiveRangeClassification as unknown as ReturnType<typeof vi.fn>;
    effectiveRangeClassification.mockClear();

    const out = await cm.buildWorkbookContext({
      workbook: workbook as any,
      query: "secret",
      dlp: {
        documentId: workbook.id,
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
        auditLogger: { log: vi.fn() },
      },
    });

    expect(out).toBeTruthy();

    // Perf proxy: we should use the document-level selector index rather than scanning
    // all classification records for each chunk/hit range.
    expect(effectiveRangeClassification).toHaveBeenCalledTimes(0);
  });

  it("resolves RAG sheet display names back to stable sheet ids for structured DLP enforcement", async () => {
    const workbookId = "wb-dlp-sheet-resolver";
    const displayName = "Budget";
    const stableSheetId = "Sheet2";

    // SpreadsheetApi surface returns display names; internal DLP records use stable ids.
    const spreadsheet: any = {
      listSheets: () => [displayName],
      listNonEmptyCells: (_sheet?: string) => [
        {
          address: { sheet: displayName, row: 1, col: 1 },
          cell: { value: "Header" },
        },
        {
          address: { sheet: displayName, row: 2, col: 1 },
          cell: { value: "hello" },
        },
      ],
      sheetNameResolver: {
        getSheetIdByName: (name: string) => (name.trim().toLowerCase() === displayName.toLowerCase() ? stableSheetId : null),
        getSheetNameById: (id: string) => (id === stableSheetId ? displayName : null),
      },
    };

    const classificationRecords = [
      {
        selector: { scope: "sheet", documentId: workbookId, sheetId: stableSheetId },
        classification: { level: "Restricted", labels: [] },
      },
    ];

    const embedder = new HashEmbedder({ dimension: 32 });
    const vectorStore = new InMemoryVectorStore({ dimension: 32 });
    const cm = new ContextManager({ tokenBudgetTokens: 500, workbookRag: { vectorStore, embedder, topK: 3 } });

    await cm.buildWorkbookContextFromSpreadsheetApi({
      spreadsheet,
      workbookId,
      query: "hello",
      includePromptContext: false,
      dlp: {
        documentId: workbookId,
        policy: {
          version: 1,
          allowDocumentOverrides: true,
          rules: {
            [DLP_ACTION.AI_CLOUD_PROCESSING]: {
              maxAllowed: "Public",
              allowRestrictedContent: false,
              redactDisallowed: true,
            },
          },
        },
        classificationRecords,
        auditLogger: { log: vi.fn() },
      },
    });

    const stored = await vectorStore.list({ workbookId, includeVector: false });
    expect(stored.length).toBeGreaterThan(0);
    for (const rec of stored) {
      expect(rec.metadata?.text).toContain("[REDACTED]");
      expect(rec.metadata?.text).not.toContain("hello");
    }
  });
});
