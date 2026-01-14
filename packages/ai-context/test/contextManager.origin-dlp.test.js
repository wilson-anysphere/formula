import assert from "node:assert/strict";
import test from "node:test";

import { ContextManager } from "../src/contextManager.js";
import { DLP_ACTION } from "../../security/dlp/src/actions.js";
import { DlpViolationError } from "../../security/dlp/src/errors.js";

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

function makeBlockingPolicy() {
  return {
    version: 1,
    allowDocumentOverrides: true,
    rules: {
      [DLP_ACTION.AI_CLOUD_PROCESSING]: {
        maxAllowed: "Internal",
        allowRestrictedContent: false,
        redactDisallowed: false,
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

test("buildContext: structured DLP REDACT also drops attachment payload fields for non-heuristic strings (no-op redactor)", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (text) => text,
  });

  const out = await cm.buildContext({
    sheet: {
      name: "Sheet1",
      values: [["TopSecret"]],
    },
    query: "ignore",
    attachments: [
      // `note` is a top-level field (not under `.data`) to ensure structured DLP enforcement
      // doesn't rely on heuristic redaction or specific attachment shapes.
      { type: "chart", reference: "Chart1", note: "TopSecret" },
    ],
    dlp: {
      documentId: "doc-1",
      sheetId: "Sheet1",
      policy: makePolicy(),
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: "doc-1",
            sheetId: "Sheet1",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.match(out.promptContext, /## attachments/i);
  assert.doesNotMatch(out.promptContext, /TopSecret/);
});

test("buildContext: structured DLP REDACT also redacts non-heuristic attachment references (no-op redactor)", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (text) => text,
  });

  const out = await cm.buildContext({
    sheet: {
      name: "Sheet1",
      values: [["Hello"]],
    },
    query: "ignore",
    attachments: [{ type: "chart", reference: "TopSecret" }],
    dlp: {
      documentId: "doc-1",
      sheetId: "Sheet1",
      policy: makePolicy(),
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: "doc-1",
            sheetId: "Sheet1",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.match(out.promptContext, /## attachments/i);
  assert.doesNotMatch(out.promptContext, /TopSecret/);
  assert.match(out.promptContext, /\[REDACTED\]/);
});

test("buildContext: structured DLP REDACT also redacts non-string attachment references (no-op redactor)", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (text) => text,
  });

  const out = await cm.buildContext({
    sheet: {
      name: "Sheet1",
      values: [["Hello"]],
    },
    query: "ignore",
    // `reference` should be a string, but callers may accidentally pass objects. Under structured
    // DLP redaction, treat it as disallowed metadata and redact it deterministically.
    attachments: [{ type: "chart", reference: { token: "TopSecret" } }],
    dlp: {
      documentId: "doc-1",
      sheetId: "Sheet1",
      policy: makePolicy(),
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: "doc-1",
            sheetId: "Sheet1",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.match(out.promptContext, /## attachments/i);
  assert.doesNotMatch(out.promptContext, /TopSecret/);
  assert.match(out.promptContext, /\[REDACTED\]/);
});

test("buildContext: structured DLP REDACT also redacts non-object attachment entries (no-op redactor)", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (text) => text,
  });

  const out = await cm.buildContext({
    sheet: {
      name: "Sheet1",
      values: [["Hello"]],
    },
    query: "ignore",
    attachments: ["TopSecret"],
    dlp: {
      documentId: "doc-1",
      sheetId: "Sheet1",
      policy: makePolicy(),
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: "doc-1",
            sheetId: "Sheet1",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.match(out.promptContext, /## attachments/i);
  assert.doesNotMatch(out.promptContext, /TopSecret/);
  assert.match(out.promptContext, /\[REDACTED\]/);
});

test("buildContext: structured DLP REDACT also redacts non-array attachments payloads (no-op redactor)", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (text) => text,
  });

  const out = await cm.buildContext({
    sheet: {
      name: "Sheet1",
      values: [["Hello"]],
    },
    query: "ignore",
    // `attachments` should be an array, but callers may accidentally pass a scalar. Under structured
    // DLP redaction, treat it as prompt-unsafe to avoid leaking non-heuristic strings.
    attachments: "TopSecret",
    dlp: {
      documentId: "doc-1",
      sheetId: "Sheet1",
      policy: makePolicy(),
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: "doc-1",
            sheetId: "Sheet1",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.match(out.promptContext, /## attachments/i);
  assert.doesNotMatch(out.promptContext, /TopSecret/);
  assert.match(out.promptContext, /\[REDACTED\]/);
});

test("buildContext: structured DLP REDACT also redacts non-heuristic table/namedRange names (no-op redactor)", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (text) => text,
  });

  const out = await cm.buildContext({
    sheet: {
      name: "Sheet1",
      values: [["Hello"]],
      tables: [{ name: "TopSecret", range: "Sheet1!A1:A1" }],
      namedRanges: [{ name: "TopSecretRange", range: "Sheet1!A1:A1" }],
    },
    query: "hello",
    dlp: {
      documentId: "doc-1",
      sheetId: "Sheet1",
      policy: makePolicy(),
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: "doc-1",
            sheetId: "Sheet1",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.match(out.promptContext, /\[REDACTED\]/);
  assert.doesNotMatch(out.promptContext, /TopSecret/);
  assert.doesNotMatch(JSON.stringify(out.schema), /TopSecret/);
});

test("buildContext: structured DLP REDACT treats table/namedRange names as disallowed metadata even when their ranges are allowed (no-op redactor)", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (text) => text,
  });

  const out = await cm.buildContext({
    sheet: {
      name: "Sheet1",
      values: [
        ["a", "b"],
        ["c", "d"],
      ],
      // Explicit schema metadata may contain non-heuristic sensitive identifiers.
      tables: [{ name: "TopSecretTable", range: "Sheet1!A1:A1" }],
      namedRanges: [{ name: "TopSecretRange", range: "Sheet1!A1:A1" }],
    },
    query: "hello",
    dlp: {
      documentId: "doc-1",
      sheetId: "Sheet1",
      policy: makePolicy(),
      classificationRecords: [
        {
          selector: {
            scope: "cell",
            documentId: "doc-1",
            sheetId: "Sheet1",
            row: 1,
            col: 1, // B2 (does not intersect A1)
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.match(out.promptContext, /\[REDACTED\]/);
  assert.doesNotMatch(out.promptContext, /TopSecretTable/);
  assert.doesNotMatch(out.promptContext, /TopSecretRange/);
  assert.doesNotMatch(JSON.stringify(out.schema), /TopSecretTable/);
  assert.doesNotMatch(JSON.stringify(out.schema), /TopSecretRange/);
});

test("buildContext: structured DLP REDACT also redacts non-heuristic sheet names when the sheet is classified (no-op redactor)", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (text) => text,
  });

  const out = await cm.buildContext({
    sheet: {
      name: "TopSecret",
      values: [["Hello"]],
    },
    query: "hello",
    attachments: [{ type: "range", reference: "TopSecret!A1:A1" }],
    dlp: {
      documentId: "doc-1",
      sheetId: "TopSecret",
      policy: makePolicy(),
      classificationRecords: [
        {
          selector: { scope: "sheet", documentId: "doc-1", sheetId: "TopSecret" },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.match(out.promptContext, /\[REDACTED\]/);
  assert.doesNotMatch(out.promptContext, /TopSecret/);
  assert.doesNotMatch(JSON.stringify(out.schema), /TopSecret/);
  assert.doesNotMatch(JSON.stringify(out.retrieved), /TopSecret/);
});

test("buildContext: structured DLP REDACT also redacts sheet names in attachment_data range previews (no-op redactor)", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (text) => text,
  });

  const out = await cm.buildContext({
    sheet: {
      name: "TopSecretSheet",
      values: [["Hello"], ["World"]],
    },
    query: "ignore",
    attachments: [{ type: "range", reference: "A1:A2" }],
    dlp: {
      documentId: "doc-1",
      sheetId: "TopSecretSheet",
      policy: makePolicy(),
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: "doc-1",
            sheetId: "TopSecretSheet",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  const attachmentDataSection =
    out.promptContext.match(/## attachment_data\n([\s\S]*?)(?:\n\n## [^\n]+\n|$)/i)?.[1] ?? "";

  assert.match(out.promptContext, /## attachment_data/i);
  assert.doesNotMatch(attachmentDataSection, /TopSecretSheet/);
  assert.match(attachmentDataSection, /\[REDACTED\]/);
});

test("buildContext: attachment-only sensitive patterns can trigger DLP REDACT (even when sheet window is public)", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (text) => text,
  });

  const auditEvents = [];

  const out = await cm.buildContext({
    sheet: {
      name: "Sheet1",
      values: [["Hello"], ["World"]],
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
      auditLogger: { log: (e) => auditEvents.push(e) },
    },
  });

  assert.match(out.promptContext, /## dlp/i);
  assert.match(out.promptContext, /\[REDACTED\]/);
  assert.doesNotMatch(out.promptContext, /123-45-6789/);
  assert.equal(auditEvents.length, 1);
  assert.equal(auditEvents[0]?.decision?.decision, "redact");
});

test("buildContext: attachment-only sensitive patterns inside a cyclic Map still trigger DLP REDACT", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (text) => text,
  });

  const auditEvents = [];

  const map = new Map();
  map.set("self", map);
  map.set("ssn", "123-45-6789");

  const out = await cm.buildContext({
    sheet: {
      name: "Sheet1",
      values: [["Hello"], ["World"]],
    },
    query: "anything",
    attachments: [{ type: "chart", reference: "CycleMap", data: map }],
    dlp: {
      documentId: "doc-1",
      sheetId: "Sheet1",
      policy: makePolicy(),
      classificationRecords: [],
      auditLogger: { log: (e) => auditEvents.push(e) },
    },
  });

  assert.match(out.promptContext, /## dlp/i);
  assert.match(out.promptContext, /\[REDACTED\]/);
  assert.doesNotMatch(out.promptContext, /123-45-6789/);
  assert.equal(auditEvents.length, 1);
  assert.equal(auditEvents[0]?.decision?.decision, "redact");
});

test("buildContext: heuristic DLP also considers sheet name and redacts it when policy requires (no-op redactor)", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (text) => text,
  });

  const auditEvents = [];

  const out = await cm.buildContext({
    sheet: {
      name: "alice@example.com",
      values: [["Header"], ["Value"]],
    },
    query: "value",
    dlp: {
      documentId: "doc-1",
      sheetId: "sheet-1",
      policy: makePolicy(),
      classificationRecords: [],
      auditLogger: { log: (e) => auditEvents.push(e) },
    },
  });

  assert.match(out.promptContext, /## dlp/i);
  assert.doesNotMatch(out.promptContext, /alice@example\.com/);
  assert.match(out.promptContext, /\[REDACTED\]/);
  assert.doesNotMatch(JSON.stringify(out.schema), /alice@example\.com/);
  assert.doesNotMatch(JSON.stringify(out.retrieved), /alice@example\.com/);
  assert.equal(auditEvents.length, 1);
  assert.equal(auditEvents[0]?.decision?.decision, "redact");
});

test("buildContext: heuristic DLP also considers sheet metadata (namedRanges/tables) when policy requires", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (text) => text,
  });

  const auditEvents = [];

  const out = await cm.buildContext({
    sheet: {
      name: "Sheet1",
      values: [["Header"], ["Value"]],
      namedRanges: [{ name: "alice@example.com", range: "Sheet1!A1:A2" }],
    },
    query: "value",
    dlp: {
      documentId: "doc-1",
      sheetId: "Sheet1",
      policy: makePolicy(),
      classificationRecords: [],
      auditLogger: { log: (e) => auditEvents.push(e) },
    },
  });

  assert.match(out.promptContext, /## dlp/i);
  assert.doesNotMatch(out.promptContext, /alice@example\.com/);
  assert.match(out.promptContext, /\[REDACTED\]/);
  assert.doesNotMatch(JSON.stringify(out.schema), /alice@example\.com/);
  assert.equal(auditEvents.length, 1);
  assert.equal(auditEvents[0]?.decision?.decision, "redact");
});

test("buildContext: heuristic DLP detects percent-encoded sensitive tokens (e.g. alice%40example.com) even with a no-op redactor", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (text) => text,
  });

  const auditEvents = [];

  const out = await cm.buildContext({
    sheet: {
      name: "Sheet1",
      values: [["Email"], ["alice%40example.com"]],
    },
    query: "email",
    dlp: {
      documentId: "doc-1",
      sheetId: "Sheet1",
      policy: makePolicy(),
      classificationRecords: [],
      auditLogger: { log: (e) => auditEvents.push(e) },
    },
  });

  assert.match(out.promptContext, /## dlp/i);
  assert.match(out.promptContext, /\[REDACTED\]/);
  assert.doesNotMatch(out.promptContext, /alice%40example\.com/);
  assert.doesNotMatch(out.promptContext, /alice@example\.com/);
  assert.doesNotMatch(JSON.stringify(out.sampledRows), /alice%40example\.com/);
  assert.doesNotMatch(JSON.stringify(out.sampledRows), /alice@example\.com/);
  assert.doesNotMatch(JSON.stringify(out.retrieved), /alice%40example\.com/);
  assert.doesNotMatch(JSON.stringify(out.retrieved), /alice@example\.com/);

  assert.equal(auditEvents.length, 1);
  assert.equal(auditEvents[0]?.decision?.decision, "redact");
});

test("buildContext: heuristic DLP sheet-name findings can block when policy requires", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (text) => text,
  });

  /** @type {any[]} */
  const auditEvents = [];

  await assert.rejects(
    cm.buildContext({
      sheet: {
        name: "alice@example.com",
        values: [["Header"], ["Value"]],
      },
      query: "value",
      dlp: {
        documentId: "doc-1",
        sheetId: "sheet-1",
        policy: makeBlockingPolicy(),
        classificationRecords: [],
        auditLogger: { log: (e) => auditEvents.push(e) },
      },
    }),
    (err) => err instanceof DlpViolationError,
  );

  assert.equal(auditEvents.length, 1);
  assert.equal(auditEvents[0]?.decision?.decision, "block");
  assert.equal(cm.ragIndex.store.size, 0);
});

test("buildContext: attachment-only sensitive patterns can trigger DLP BLOCK when policy requires", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (text) => text,
  });

  /** @type {any[]} */
  const auditEvents = [];

  await assert.rejects(
    cm.buildContext({
      sheet: {
        name: "Sheet1",
        values: [["Hello"], ["World"]],
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
        policy: makeBlockingPolicy(),
        classificationRecords: [],
        auditLogger: { log: (e) => auditEvents.push(e) },
      },
    }),
    (err) => err instanceof DlpViolationError,
  );

  assert.equal(auditEvents.length, 1);
  assert.equal(auditEvents[0]?.decision?.decision, "block");
  assert.equal(cm.ragIndex.store.size, 0);
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

test("buildContext: DLP REDACT deep-redacts class instance attachment data", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (text) => text,
  });

  class SecretPayload {
    constructor() {
      this.ssn = "123-45-6789";
    }
  }

  const out = await cm.buildContext({
    sheet: {
      name: "Sheet1",
      values: [["Email"], ["alice@example.com"]],
    },
    query: "anything",
    attachments: [
      {
        type: "chart",
        reference: "ClassAttachment",
        data: new SecretPayload(),
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
  assert.match(out.promptContext, /\[REDACTED\]/);
});

test("buildContext: DLP REDACT handles cyclic attachment data without crashing", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (text) => text,
  });

  const cyclic = { note: "123-45-6789" };
  cyclic.self = cyclic;

  const out = await cm.buildContext({
    sheet: {
      name: "Sheet1",
      values: [["Email"], ["alice@example.com"]],
    },
    query: "anything",
    attachments: [
      {
        type: "chart",
        reference: "CycleAttachment",
        data: cyclic,
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
  assert.match(out.promptContext, /\[REDACTED\]/);
});

test("buildContext: DLP REDACT redacts heuristic-sensitive strings in object keys", async () => {
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
        reference: "KeyLeak",
        // Sensitive string used as an object key (could leak through JSON serialization).
        data: { "alice@example.com": "ok" },
      },
    ],
    dlp: {
      documentId: "doc-1",
      sheetId: "Sheet1",
      policy: makePolicy(),
      classificationRecords: [],
    },
  });

  assert.doesNotMatch(out.promptContext, /alice@example\.com/);
  assert.match(out.promptContext, /\[REDACTED_KEY/);
});
