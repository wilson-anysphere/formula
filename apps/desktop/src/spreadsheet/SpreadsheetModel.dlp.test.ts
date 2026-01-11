import { beforeEach, describe, expect, it } from "vitest";

import { SpreadsheetModel } from "./SpreadsheetModel.js";

import { createDefaultOrgPolicy } from "../../../../packages/security/dlp/src/policy.js";
import { DLP_ACTION } from "../../../../packages/security/dlp/src/actions.js";
import { LocalPolicyStore } from "../../../../packages/security/dlp/src/policyStore.js";
import { LocalClassificationStore } from "../../../../packages/security/dlp/src/classificationStore.js";
import { CLASSIFICATION_SCOPE } from "../../../../packages/security/dlp/src/selectors.js";
import { CLASSIFICATION_LEVEL } from "../../../../packages/security/dlp/src/classification.js";

describe("SpreadsheetModel AI cell functions (DLP wiring)", () => {
  beforeEach(() => {
    globalThis.localStorage?.clear();
  });

  it("blocks restricted cell references using document policy + classification records", () => {
    const workbookId = "local-workbook";
    const storage = globalThis.localStorage as any;

    const policyStore = new LocalPolicyStore({ storage });
    const policy = createDefaultOrgPolicy();
    policy.rules[DLP_ACTION.AI_CLOUD_PROCESSING] = {
      ...policy.rules[DLP_ACTION.AI_CLOUD_PROCESSING],
      redactDisallowed: false,
    };
    policyStore.setDocumentPolicy(workbookId, policy);

    const classificationStore = new LocalClassificationStore({ storage });
    classificationStore.upsert(
      workbookId,
      { scope: CLASSIFICATION_SCOPE.CELL, documentId: workbookId, sheetId: "Sheet1", row: 0, col: 0 },
      { level: CLASSIFICATION_LEVEL.RESTRICTED, labels: [] },
    );

    const sheet = new SpreadsheetModel({ A1: "top secret" });
    sheet.setCellInput("B1", '=AI("summarize", A1)');
    expect(sheet.getCellValue("B1")).toBe("#DLP!");
  });
});

