import { describe, expect, it } from "vitest";

import { DocumentController } from "../../../document/documentController.js";

import { DLP_ACTION } from "../../../../../../packages/security/dlp/src/actions.js";
import { createDefaultOrgPolicy } from "../../../../../../packages/security/dlp/src/policy.js";
import { CLASSIFICATION_LEVEL } from "../../../../../../packages/security/dlp/src/classification.js";
import { LocalClassificationStore, createMemoryStorage } from "../../../../../../packages/security/dlp/src/classificationStore.js";

import { DocumentControllerSpreadsheetApi } from "../../tools/documentControllerSpreadsheetApi.js";
import { WorkbookContextBuilder } from "../WorkbookContextBuilder.js";

describe("WorkbookContextBuilder", () => {
  it("extracts a schema-first summary from a sheet with headers", async () => {
    const documentController = new DocumentController();
    documentController.setRangeValues("Sheet1", "A1", [
      ["Name", "Age"],
      ["Alice", 30],
      ["Bob", 40],
    ]);

    const spreadsheet = new DocumentControllerSpreadsheetApi(documentController);
    const builder = new WorkbookContextBuilder({
      workbookId: "wb_schema",
      documentController,
      spreadsheet,
      ragService: null,
      mode: "chat",
      model: "unit-test-model",
    });

    const ctx = await builder.build({ activeSheetId: "Sheet1" });
    const sheet = ctx.payload.sheets.find((s) => s.sheetId === "Sheet1");
    expect(sheet).toBeTruthy();

    expect(sheet!.schema.dataRegions).toHaveLength(1);
    expect(sheet!.schema.dataRegions[0]!.hasHeader).toBe(true);
    expect(sheet!.schema.dataRegions[0]!.range).toBe("Sheet1!A1:B3");

    expect(sheet!.schema.tables).toHaveLength(1);
    const table = sheet!.schema.tables[0]!;
    expect(table.columns.map((c) => c.name)).toEqual(["Name", "Age"]);
    expect(table.columns.map((c) => c.type)).toEqual(["string", "number"]);
  });

  it("includes explicit named ranges and tables when provided by a schemaProvider", async () => {
    const documentController = new DocumentController();
    documentController.setRangeValues("Sheet1", "A1", [
      ["Region", "Revenue"],
      ["North", 1000],
      ["South", 2000],
    ]);

    const spreadsheet = new DocumentControllerSpreadsheetApi(documentController);
    const builder = new WorkbookContextBuilder({
      workbookId: "wb_meta",
      documentController,
      spreadsheet,
      ragService: null,
      mode: "chat",
      model: "unit-test-model",
      schemaProvider: {
        getNamedRanges: () => [
          { name: "SalesData", sheetId: "Sheet1", range: { startRow: 0, endRow: 2, startCol: 0, endCol: 1 } },
        ],
        getTables: () => [
          { name: "SalesTable", sheetId: "Sheet1", range: { startRow: 0, endRow: 2, startCol: 0, endCol: 1 } },
        ],
      },
    });

    const ctx = await builder.build({ activeSheetId: "Sheet1" });
    expect(ctx.payload.namedRanges).toEqual([{ name: "SalesData", range: "Sheet1!A1:B3" }]);

    const sheet = ctx.payload.sheets.find((s) => s.sheetId === "Sheet1");
    expect(sheet).toBeTruthy();
    expect(sheet!.schema.namedRanges).toEqual([{ name: "SalesData", range: "Sheet1!A1:B3" }]);
    expect(sheet!.schema.tables[0]!.name).toBe("SalesTable");

    expect(ctx.payload.tables).toEqual([{ sheetId: "Sheet1", name: "SalesTable", range: "Sheet1!A1:B3" }]);
  });

  it("redacts restricted cells in sampled blocks (DLP)", async () => {
    const workbookId = "wb_dlp";
    const documentController = new DocumentController();
    documentController.setRangeValues("Sheet1", "A1", [
      ["Public"],
      ["TOP SECRET"],
    ]);

    const storage = createMemoryStorage();
    const classificationStore = new LocalClassificationStore({ storage });
    classificationStore.upsert(
      workbookId,
      { scope: "cell", documentId: workbookId, sheetId: "Sheet1", row: 1, col: 0 },
      { level: CLASSIFICATION_LEVEL.RESTRICTED, labels: ["test"] },
    );

    const classificationRecords = classificationStore.list(workbookId);
    const policy = createDefaultOrgPolicy();

    const auditLogger = { log: (_event: any) => {} };

    const dlp = {
      // ContextManager style
      documentId: workbookId,
      sheetId: "Sheet1",
      policy,
      classificationRecords,
      classificationStore,
      includeRestrictedContent: false,
      auditLogger,
      // ToolExecutor style
      document_id: workbookId,
      sheet_id: "Sheet1",
      classification_records: classificationRecords,
      classification_store: classificationStore,
      include_restricted_content: false,
      audit_logger: auditLogger,
    };

    const spreadsheet = new DocumentControllerSpreadsheetApi(documentController);
    const builder = new WorkbookContextBuilder({
      workbookId,
      documentController,
      spreadsheet,
      ragService: null,
      dlp,
      mode: "inline_edit",
      model: "unit-test-model",
    });

    const ctx = await builder.build({
      activeSheetId: "Sheet1",
      selectedRange: { sheetId: "Sheet1", range: { startRow: 0, endRow: 1, startCol: 0, endCol: 0 } },
    });

    const selection = ctx.payload.blocks.find((b) => b.kind === "selection");
    expect(selection).toBeTruthy();

    expect(selection!.values[0]![0]).toBe("Public");
    expect(selection!.values[1]![0]).toBe("[REDACTED]");
    expect(ctx.promptContext).toContain("[REDACTED]");
    expect(ctx.promptContext).not.toContain("TOP SECRET");
  });

  it("omits schema extraction when sheet sampling is denied by policy (no placeholder-derived tables)", async () => {
    const workbookId = "wb_policy_denied";
    const documentController = new DocumentController();
    documentController.setRangeValues("Sheet1", "A1", [["TOP SECRET"]]);

    const storage = createMemoryStorage();
    const classificationStore = new LocalClassificationStore({ storage });
    classificationStore.upsert(
      workbookId,
      { scope: "cell", documentId: workbookId, sheetId: "Sheet1", row: 0, col: 0 },
      { level: CLASSIFICATION_LEVEL.RESTRICTED, labels: ["test"] },
    );

    const classificationRecords = classificationStore.list(workbookId);
    const policy = createDefaultOrgPolicy();
    // Force a hard block instead of redaction.
    (policy.rules as any)[DLP_ACTION.AI_CLOUD_PROCESSING].redactDisallowed = false;

    const auditLogger = { log: (_event: any) => {} };
    const dlp = {
      documentId: workbookId,
      sheetId: "Sheet1",
      policy,
      classificationRecords,
      classificationStore,
      includeRestrictedContent: false,
      auditLogger,
      document_id: workbookId,
      sheet_id: "Sheet1",
      classification_records: classificationRecords,
      classification_store: classificationStore,
      include_restricted_content: false,
      audit_logger: auditLogger,
    };

    const spreadsheet = new DocumentControllerSpreadsheetApi(documentController);
    const builder = new WorkbookContextBuilder({
      workbookId,
      documentController,
      spreadsheet,
      ragService: null,
      dlp,
      mode: "chat",
      model: "unit-test-model",
    });

    const ctx = await builder.build({
      activeSheetId: "Sheet1",
      selectedRange: { sheetId: "Sheet1", range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 } },
    });

    const sheet = ctx.payload.sheets.find((s) => s.sheetId === "Sheet1");
    expect(sheet).toBeTruthy();
    expect(sheet!.schema.tables).toEqual([]);
    expect(sheet!.schema.dataRegions).toEqual([]);

    const selection = ctx.payload.blocks.find((b) => b.kind === "selection");
    expect(selection).toBeTruthy();
    expect(selection!.values[0]![0]).toBe("[POLICY_DENIED]");
    expect(ctx.promptContext).toContain("[POLICY_DENIED]");
    expect(ctx.promptContext).not.toContain("TOP SECRET");
  });

  it("serializes a stable payload snapshot", async () => {
    const documentController = new DocumentController();
    documentController.setRangeValues("Sheet1", "A1", [
      ["Region", "Revenue"],
      ["North", 1000],
      ["South", 2000],
    ]);
    documentController.setRangeValues("Sheet2", "A1", [["Note"], ["Hello"]]);

    const spreadsheet = new DocumentControllerSpreadsheetApi(documentController);
    const builder = new WorkbookContextBuilder({
      workbookId: "wb_snapshot",
      documentController,
      spreadsheet,
      ragService: null,
      mode: "chat",
      model: "unit-test-model",
      maxSheets: 5,
    });

    const ctx = await builder.build({
      activeSheetId: "Sheet1",
      selectedRange: { sheetId: "Sheet1", range: { startRow: 0, endRow: 2, startCol: 0, endCol: 1 } },
      focusQuestion: "revenue by region",
    });

    // Snapshot a sanitized payload (strip noisy counters; keep ordering stable).
    const snapshotPayload = { ...ctx.payload, budget: { ...ctx.payload.budget, usedPromptContextTokens: 0 } };
    expect(WorkbookContextBuilder.serializePayload(snapshotPayload as any)).toMatchSnapshot();
  });
});
