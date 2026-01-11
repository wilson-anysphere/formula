import test from "node:test";
import assert from "node:assert/strict";

import { LocalClassificationStore, createMemoryStorage } from "../../packages/security/dlp/src/classificationStore.js";
import { CLASSIFICATION_LEVEL } from "../../packages/security/dlp/src/classification.js";
import { CLASSIFICATION_SCOPE } from "../../packages/security/dlp/src/selectors.js";
import { createDefaultOrgPolicy } from "../../packages/security/dlp/src/policy.js";
import { DocumentController } from "../../apps/desktop/src/document/documentController.js";
import { copyRangeToClipboardPayload } from "../../apps/desktop/src/clipboard/clipboard.js";
import { exportDocumentRangeToCsv } from "../../apps/desktop/src/import-export/csv/export.js";
import { InMemoryAuditLogger } from "../../packages/security/dlp/src/audit.js";
import { AiContextManager } from "../../packages/security/dlp/src/aiContextManager.js";
import { DlpViolationError } from "../../packages/security/dlp/src/errors.js";

test("integration: Restricted range blocks clipboard + export and redacts AI context", () => {
  const storage = createMemoryStorage();
  const classificationStore = new LocalClassificationStore({ storage });
  const policy = createDefaultOrgPolicy();

  const documentId = "doc_1";
  const sheetId = "sheet_1";
  const doc = new DocumentController();
  const restrictedRangeSelector = {
    scope: CLASSIFICATION_SCOPE.RANGE,
    documentId,
    sheetId,
    range: {
      start: { row: 0, col: 0 }, // A1
      end: { row: 1, col: 1 }, // B2
    },
  };

  classificationStore.upsert(documentId, restrictedRangeSelector, {
    level: CLASSIFICATION_LEVEL.RESTRICTED,
    labels: ["PII"],
  });

  const range = restrictedRangeSelector.range;

  doc.setCellValue(sheetId, { row: 0, col: 0 }, "Alice");
  doc.setCellValue(sheetId, { row: 0, col: 1 }, "111-22-3333");
  doc.setCellValue(sheetId, { row: 1, col: 0 }, "Bob");
  doc.setCellValue(sheetId, { row: 1, col: 1 }, "444-55-6666");

  assert.throws(
    () => {
      copyRangeToClipboardPayload(doc, sheetId, range, {
        dlp: { documentId, classificationStore, policy },
      });
    },
    (err) => err instanceof DlpViolationError && /Clipboard copy is blocked/.test(err.message),
  );

  assert.throws(
    () => {
      exportDocumentRangeToCsv(doc, sheetId, range, {
        dlp: { documentId, classificationStore, policy },
      });
    },
    (err) => err instanceof DlpViolationError && /Export is blocked/.test(err.message),
  );

  const auditLogger = new InMemoryAuditLogger();
  const aiContextManager = new AiContextManager({ classificationStore, auditLogger });

  const cells = [
    { row: 0, col: 0, value: "Alice" },
    { row: 0, col: 1, value: "111-22-3333" },
    { row: 1, col: 0, value: "Bob" },
    { row: 1, col: 1, value: "444-55-6666" },
  ];

  const { context, redactions } = aiContextManager.buildCloudContext({
    documentId,
    sheetId,
    range,
    cells,
    policy,
  });

  assert.equal(redactions.length, 4);
  assert.ok(context.includes("[REDACTED]"));

  const events = auditLogger.list();
  assert.equal(events.length, 1);
  assert.equal(events[0].details.redactedCellCount, 4);
});
