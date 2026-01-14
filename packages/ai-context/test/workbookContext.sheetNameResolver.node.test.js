import assert from "node:assert/strict";
import test from "node:test";

import { ContextManager } from "../src/contextManager.js";
import { HashEmbedder } from "../../ai-rag/src/embedding/hashEmbedder.js";
import { InMemoryVectorStore } from "../../ai-rag/src/store/inMemoryVectorStore.js";
import { DLP_ACTION } from "../../security/dlp/src/actions.js";

test("buildWorkbookContextFromSpreadsheetApi: resolves sheet display names back to stable sheet ids for structured DLP enforcement", async () => {
  const workbookId = "wb-dlp-sheet-resolver";
  const displayName = "Budget";
  const stableSheetId = "Sheet2";

  // SpreadsheetApi surface returns display names; internal DLP records use stable ids.
  const spreadsheet = {
    listSheets: () => [displayName],
    listNonEmptyCells: (_sheet) => [
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
      getSheetIdByName: (name) => (name.trim().toLowerCase() === displayName.toLowerCase() ? stableSheetId : null),
      getSheetNameById: (id) => (id === stableSheetId ? displayName : null),
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
      auditLogger: { log: () => {} },
    },
  });

  const stored = await vectorStore.list({ workbookId, includeVector: false });
  assert.ok(stored.length > 0);
  for (const rec of stored) {
    assert.ok(String(rec.metadata?.text ?? "").includes("[REDACTED]"));
    assert.ok(!String(rec.metadata?.text ?? "").includes("hello"));
  }
});

