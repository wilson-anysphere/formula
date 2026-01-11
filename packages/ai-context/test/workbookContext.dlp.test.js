import assert from "node:assert/strict";
import test from "node:test";

import { ContextManager } from "../src/contextManager.js";
import { HashEmbedder, InMemoryVectorStore } from "../../ai-rag/src/index.js";
import { DLP_ACTION } from "../../security/dlp/src/actions.js";
import { DlpViolationError } from "../../security/dlp/src/errors.js";

function makeSensitiveWorkbook() {
  return {
    id: "wb-dlp",
    sheets: [
      {
        name: "Contacts",
        cells: [
          [{ v: "Name" }, { v: "Email" }, { v: "SSN" }],
          [{ v: "Alice" }, { v: "alice@example.com" }, { v: "123-45-6789" }],
        ],
      },
    ],
    tables: [
      {
        name: "ContactsTable",
        sheetName: "Contacts",
        rect: { r0: 0, c0: 0, r1: 1, c1: 2 },
      },
    ],
  };
}

function makePolicy({ maxAllowed = "Public", redactDisallowed }) {
  return {
    version: 1,
    allowDocumentOverrides: false,
    rules: {
      [DLP_ACTION.AI_CLOUD_PROCESSING]: {
        maxAllowed,
        allowRestrictedContent: false,
        redactDisallowed,
      },
    },
  };
}

test("buildWorkbookContext: blocks sensitive workbook chunks when policy blocks", async () => {
  const workbook = makeSensitiveWorkbook();
  const embedder = new HashEmbedder({ dimension: 128 });
  const vectorStore = new InMemoryVectorStore({ dimension: 128 });

  const cm = new ContextManager({
    tokenBudgetTokens: 500,
    workbookRag: { vectorStore, embedder, topK: 3 },
  });

  const auditEvents = [];

  await assert.rejects(
    () =>
      cm.buildWorkbookContext({
        workbook,
        query: "alice@example.com",
        dlp: {
          documentId: workbook.id,
          policy: makePolicy({ redactDisallowed: false }),
          auditLogger: { log: (e) => auditEvents.push(e) },
        },
      }),
    (err) => {
      assert.ok(err instanceof DlpViolationError);
      return true;
    }
  );

  assert.equal(auditEvents.length, 1);
  assert.equal(auditEvents[0].type, "ai.workbook_context");
  assert.equal(auditEvents[0].documentId, workbook.id);
  assert.equal(auditEvents[0].decision.decision, "block");
});

test("buildWorkbookContext: redacts sensitive workbook chunks when policy allows with redaction", async () => {
  const workbook = makeSensitiveWorkbook();
  const embedder = new HashEmbedder({ dimension: 128 });
  const vectorStore = new InMemoryVectorStore({ dimension: 128 });

  const cm = new ContextManager({
    tokenBudgetTokens: 800,
    workbookRag: { vectorStore, embedder, topK: 3 },
  });

  const auditEvents = [];

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "123-45-6789",
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ redactDisallowed: true }),
      auditLogger: { log: (e) => auditEvents.push(e) },
    },
  });

  assert.match(out.promptContext, /## dlp/i);
  assert.match(out.promptContext, /\[REDACTED_(EMAIL|SSN)\]/);
  assert.ok(out.retrieved.length > 0);
  assert.equal(out.retrieved[0].metadata.text, undefined);
  assert.equal(auditEvents.length, 1);
  assert.equal(auditEvents[0].type, "ai.workbook_context");
  assert.equal(auditEvents[0].documentId, workbook.id);
  assert.equal(auditEvents[0].decision.decision, "redact");
  assert.equal(auditEvents[0].redactedChunkCount, 1);
});

test("buildWorkbookContext: structured Restricted classifications fully redact retrieved chunks", async () => {
  const workbook = makeSensitiveWorkbook();
  // Add a value that isn't handled by the regex redactor but should still be suppressed by
  // explicit DLP classification.
  workbook.sheets[0].cells[1][0].v = "TopSecret";

  const embedder = new HashEmbedder({ dimension: 128 });
  const vectorStore = new InMemoryVectorStore({ dimension: 128 });

  const cm = new ContextManager({
    tokenBudgetTokens: 800,
    workbookRag: { vectorStore, embedder, topK: 3 },
  });

  const auditEvents = [];

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "TopSecret",
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Confidential", redactDisallowed: true }),
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: workbook.id,
            sheetId: "Contacts",
            range: { start: { row: 0, col: 0 }, end: { row: 1, col: 2 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
      auditLogger: { log: (e) => auditEvents.push(e) },
    },
  });

  assert.ok(out.retrieved.length > 0);
  assert.match(out.retrieved[0].text, /\[REDACTED\]/);
  assert.doesNotMatch(out.promptContext, /TopSecret/);
  assert.equal(auditEvents.length, 1);
  assert.equal(auditEvents[0].decision.decision, "redact");
});
