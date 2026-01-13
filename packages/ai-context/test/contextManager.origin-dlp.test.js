import assert from "node:assert/strict";
import test from "node:test";

import { ContextManager } from "../src/contextManager.js";
import { DLP_ACTION } from "../../security/dlp/src/actions.js";

function makePolicy() {
  return {
    version: 1,
    allowDocumentOverrides: true,
    rules: {
      [DLP_ACTION.AI_CLOUD_PROCESSING]: {
        maxAllowed: "Internal",
        allowRestrictedContent: false,
        redactDisallowed: true,
      },
    },
  };
}

test("buildContext: origin-offset structured DLP selectors redact the correct absolute cells", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (text) => text,
  });

  const auditEvents = [];

  const out = await cm.buildContext({
    sheet: {
      name: "Sheet1",
      origin: { row: 10, col: 10 }, // K11
      values: [
        ["a", "b"],
        ["c", "TOP SECRET"],
      ],
    },
    query: "secret",
    dlp: {
      documentId: "doc-1",
      sheetId: "Sheet1",
      policy: makePolicy(),
      classificationRecords: [
        {
          selector: { scope: "cell", documentId: "doc-1", sheetId: "Sheet1", row: 11, col: 11 },
          classification: { level: "Restricted", labels: [] },
        },
      ],
      auditLogger: { log: (e) => auditEvents.push(e) },
    },
  });

  assert.equal(out.schema.dataRegions[0]?.range, "Sheet1!K11:L12");
  assert.equal(out.retrieved[0]?.range, "Sheet1!K11:L12");

  assert.equal(out.sampledRows.length, 2);
  assert.equal(out.sampledRows[1][1], "[REDACTED]");

  assert.match(out.promptContext, /\[REDACTED\]/);
  assert.doesNotMatch(out.promptContext, /TOP SECRET/);

  assert.equal(auditEvents.length, 1);
  assert.equal(auditEvents[0]?.type, "ai.context");
});

test("buildContext: structured DLP selectors outside the origin window do not affect redaction", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (text) => text,
  });

  const out = await cm.buildContext({
    sheet: {
      name: "Sheet1",
      origin: { row: 10, col: 10 },
      values: [
        ["a", "b"],
        ["c", "d"],
      ],
    },
    query: "a",
    dlp: {
      documentId: "doc-1",
      sheetId: "Sheet1",
      policy: makePolicy(),
      // Classified cell is outside the provided (origin-offset) window.
      classificationRecords: [
        {
          selector: { scope: "cell", documentId: "doc-1", sheetId: "Sheet1", row: 0, col: 1 },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.equal(out.sampledRows[0][1], "b");
  assert.equal(out.sampledRows[1][1], "d");
  assert.doesNotMatch(out.promptContext, /\[REDACTED\]/);
});

test("buildContext: DLP REDACT also prevents attachments from leaking heuristic-sensitive strings (even with a no-op redactor)", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (text) => text,
  });

  const out = await cm.buildContext({
    sheet: {
      name: "Sheet1",
      values: [["Email"], ["alice@example.com"]],
    },
    query: "anything",
    attachments: [
      {
        type: "chart",
        reference: "Chart1",
        data: { note: "123-45-6789" },
      },
    ],
    dlp: {
      documentId: "doc-1",
      sheetId: "Sheet1",
      policy: makePolicy(),
      classificationRecords: [],
    },
  });

  assert.match(out.promptContext, /## attachments/i);
  assert.doesNotMatch(out.promptContext, /123-45-6789/);
  assert.match(out.promptContext, /\[REDACTED\]/);
});

test("buildContext: DLP REDACT also prevents attachment_data range previews from leaking heuristic-sensitive strings (even with a no-op redactor)", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (text) => text,
  });

  const out = await cm.buildContext({
    sheet: {
      name: "Sheet1",
      values: [["Email"], ["alice@example.com"]],
    },
    query: "anything",
    // Force the range preview section to include the sensitive cell.
    attachments: [{ type: "range", reference: "A1:A2" }],
    dlp: {
      documentId: "doc-1",
      sheetId: "Sheet1",
      policy: makePolicy(),
      classificationRecords: [],
    },
  });

  assert.match(out.promptContext, /## attachment_data/i);
  assert.doesNotMatch(out.promptContext, /alice@example\.com/);
  assert.match(out.promptContext, /\[REDACTED\]/);
});

test("buildContext: heuristic DLP redaction also redacts numeric values that match sensitive patterns", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (text) => text,
  });

  const cardNumber = 4111111111111111;

  const out = await cm.buildContext({
    sheet: {
      name: "Sheet1",
      values: [["Card"], [cardNumber]],
    },
    query: "anything",
    dlp: {
      documentId: "doc-1",
      sheetId: "Sheet1",
      policy: makePolicy(),
      classificationRecords: [],
    },
  });

  assert.doesNotMatch(out.promptContext, /4111111111111111/);
  assert.match(out.promptContext, /\[REDACTED\]/);
  assert.doesNotMatch(JSON.stringify(out.sampledRows), /4111111111111111/);
  assert.doesNotMatch(JSON.stringify(out.retrieved), /4111111111111111/);
});

test("buildContext: DLP REDACT deep-redacts Map/Set values inside attachments", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (text) => text,
  });

  const out = await cm.buildContext({
    sheet: {
      name: "Sheet1",
      values: [["Email"], ["alice@example.com"]],
    },
    query: "anything",
    attachments: [
      {
        type: "chart",
        reference: "MapAttachment",
        data: new Map([
          ["ssn", "123-45-6789"],
          ["email", "alice@example.com"],
        ]),
      },
      {
        type: "chart",
        reference: "SetAttachment",
        data: new Set(["alice@example.com"]),
      },
    ],
    dlp: {
      documentId: "doc-1",
      sheetId: "Sheet1",
      policy: makePolicy(),
      classificationRecords: [],
    },
  });

  assert.doesNotMatch(out.promptContext, /123-45-6789/);
  assert.doesNotMatch(out.promptContext, /alice@example\.com/);
  assert.match(out.promptContext, /\[REDACTED\]/);
});
