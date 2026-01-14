import assert from "node:assert/strict";
import test from "node:test";

import { ContextManager } from "../src/contextManager.js";
import { DLP_ACTION } from "../../security/dlp/src/actions.js";

test("ContextManager.buildContext: resolves sheet display names to stable sheet ids for structured DLP enforcement", async () => {
  const documentId = "doc-1";
  const displayName = "Budget";
  const stableSheetId = "Sheet2";

  const cm = new ContextManager({ tokenBudgetTokens: 1_000 });

  const auditEvents = [];

  const out = await cm.buildContext({
    sheet: { name: displayName, values: [["Public"], ["TOP SECRET"]] },
    query: "secret",
    dlp: {
      // Exercise snake_case inputs (common in JSON hosts).
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
        getSheetIdByName: (name) => (name.trim().toLowerCase() === displayName.toLowerCase() ? stableSheetId : null),
      },
      auditLogger: { log: (e) => auditEvents.push(e) },
    },
  });

  assert.equal(out.schema.name, displayName);
  assert.equal(out.retrieved[0]?.range, `${displayName}!A1:A2`);
  assert.ok(out.promptContext.includes("[REDACTED]"));
  assert.ok(!out.promptContext.includes("TOP SECRET"));

  assert.equal(auditEvents.length, 1);
  assert.equal(auditEvents[0]?.type, "ai.context");
  assert.equal(auditEvents[0]?.documentId, documentId);
  assert.equal(auditEvents[0]?.sheetId, stableSheetId);
  assert.equal(auditEvents[0]?.sheetName, displayName);
});

