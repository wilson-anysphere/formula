import { beforeEach, describe, expect, it, vi } from "vitest";

let toolExecutorConstructorCalls = 0;

vi.mock("../../../../../../packages/ai-tools/src/executor/tool-executor.js", async () => {
  const actual = await vi.importActual<any>("../../../../../../packages/ai-tools/src/executor/tool-executor.js");

  return {
    ...actual,
    ToolExecutor: class ToolExecutor extends actual.ToolExecutor {
      constructor(...args: any[]) {
        super(...args);
        toolExecutorConstructorCalls += 1;
      }
    }
  };
});

import { DocumentController } from "../../../document/documentController.js";

import { ContextManager } from "../../../../../../packages/ai-context/src/contextManager.js";
import { HashEmbedder, InMemoryVectorStore } from "../../../../../../packages/ai-rag/src/index.js";

import { DLP_ACTION } from "../../../../../../packages/security/dlp/src/actions.js";
import { createDefaultOrgPolicy } from "../../../../../../packages/security/dlp/src/policy.js";
import { CLASSIFICATION_LEVEL } from "../../../../../../packages/security/dlp/src/classification.js";
import { LocalClassificationStore, createMemoryStorage } from "../../../../../../packages/security/dlp/src/classificationStore.js";

import { DocumentControllerSpreadsheetApi } from "../../tools/documentControllerSpreadsheetApi.js";
import { WorkbookContextBuilder } from "../WorkbookContextBuilder.js";

describe("WorkbookContextBuilder", () => {
  beforeEach(() => {
    toolExecutorConstructorCalls = 0;
  });

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

  it("includes retrieved chunk text in promptContext without depending on ragResult.promptContext", async () => {
    const documentController = new DocumentController();
    documentController.setRangeValues("Sheet1", "A1", [
      ["Region", "Revenue"],
      ["North", 1000],
      ["South", 2000],
    ]);

    const spreadsheet = new DocumentControllerSpreadsheetApi(documentController);

    const embedder = new HashEmbedder({ dimension: 64 });
    const vectorStore = new InMemoryVectorStore({ dimension: 64 });
    const contextManager = new ContextManager({
      tokenBudgetTokens: 800,
      workbookRag: { vectorStore, embedder, topK: 3 },
    });

    let lastRagResult: any = null;
    const ragService = {
      async buildWorkbookContextFromSpreadsheetApi(params: any) {
        lastRagResult = await contextManager.buildWorkbookContextFromSpreadsheetApi(params);
        return lastRagResult;
      },
    };

    const builder = new WorkbookContextBuilder({
      workbookId: "wb_rag_builder",
      documentController,
      spreadsheet,
      ragService: ragService as any,
      mode: "chat",
      model: "unit-test-model",
      maxPromptContextTokens: 4000,
    });

    const ctx = await builder.build({ activeSheetId: "Sheet1", focusQuestion: "revenue by region" });

    expect(lastRagResult).toBeTruthy();
    expect(lastRagResult.promptContext).toBe("");
    expect(ctx.retrieved.length).toBeGreaterThan(0);

    // WorkbookContextBuilder formats the retrieved chunks into the final packed prompt context.
    expect(ctx.promptContext).toContain("## retrieved");
    expect(ctx.promptContext).toMatch(/score=/);
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

  it("caches schemaProvider named ranges/tables by schemaVersion", async () => {
    const documentController = new DocumentController();
    documentController.setRangeValues("Sheet1", "A1", [["A"]]);

    const spreadsheet = new DocumentControllerSpreadsheetApi(documentController);

    let schemaVersion = 1;
    const getNamedRanges = vi.fn(() => [
      { name: "MyRange", sheetId: "Sheet1", range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 } },
    ]);
    const getTables = vi.fn(() => [
      { name: "MyTable", sheetId: "Sheet1", range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 } },
    ]);

    const builder = new WorkbookContextBuilder({
      workbookId: "wb_schema_cache",
      documentController,
      spreadsheet,
      ragService: null,
      mode: "chat",
      model: "unit-test-model",
      schemaProvider: {
        getSchemaVersion: () => schemaVersion,
        getNamedRanges,
        getTables,
      },
    });

    await builder.build({ activeSheetId: "Sheet1" });
    expect(getNamedRanges).toHaveBeenCalledTimes(1);
    expect(getTables).toHaveBeenCalledTimes(1);

    // Same schema version -> cached metadata -> no extra provider calls.
    await builder.build({ activeSheetId: "Sheet1" });
    expect(getNamedRanges).toHaveBeenCalledTimes(1);
    expect(getTables).toHaveBeenCalledTimes(1);

    // Version bump should invalidate the cache.
    schemaVersion += 1;
    await builder.build({ activeSheetId: "Sheet1" });
    expect(getNamedRanges).toHaveBeenCalledTimes(2);
    expect(getTables).toHaveBeenCalledTimes(2);
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

  it("reuses a single ToolExecutor instance for all read_range calls in a build", async () => {
    const documentController = new DocumentController();
    documentController.setRangeValues("Sheet1", "A1", [
      ["Region", "Revenue"],
      ["North", 1000],
      ["South", 2000],
    ]);
    documentController.setRangeValues("Sheet2", "A1", [["Note"], ["Hello"]]);

    const spreadsheet = new DocumentControllerSpreadsheetApi(documentController);
    const builder = new WorkbookContextBuilder({
      workbookId: "wb_tool_executor_reuse",
      documentController,
      spreadsheet,
      ragService: null,
      mode: "chat",
      model: "unit-test-model",
      maxSheets: 5,
    });

    await builder.build({
      activeSheetId: "Sheet1",
      selectedRange: { sheetId: "Sheet1", range: { startRow: 0, endRow: 2, startCol: 0, endCol: 1 } },
    });

    expect(toolExecutorConstructorCalls).toBe(1);
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

  it("builds a deterministic, human-readable promptContext", async () => {
    const documentController = new DocumentController();
    const header = Array.from({ length: 8 }, (_v, idx) => `Col${idx + 1}`);
    const rows = Array.from({ length: 14 }, (_v, rIdx) =>
      Array.from({ length: 8 }, (_v2, cIdx) => `R${rIdx + 1}C${cIdx + 1}`),
    );
    documentController.setRangeValues("Sheet1", "A1", [header, ...rows]);

    const spreadsheet = new DocumentControllerSpreadsheetApi(documentController);
    const builder = new WorkbookContextBuilder({
      workbookId: "wb_prompt_ctx",
      documentController,
      spreadsheet,
      ragService: null,
      mode: "chat",
      model: "unit-test-model",
      // Avoid trimming so we can compare size differences deterministically.
      maxPromptContextTokens: 1_000_000,
    });

    const input = {
      activeSheetId: "Sheet1",
      selectedRange: { sheetId: "Sheet1", range: { startRow: 0, endRow: 14, startCol: 0, endCol: 7 } },
    };

    const ctx1 = await builder.build(input);
    const ctx2 = await builder.build(input);

    // Deterministic output is critical for caching and for stable prompt packing decisions.
    expect(ctx1.promptContext).toEqual(ctx2.promptContext);

    // Prompt context should contain stable, pretty-printed JSON (human-readable).
    expect(ctx1.promptContext).toContain('\n  "');
    expect(ctx1.promptContext).toContain('"kind": "selection"');
    // Ensure we don't regress to minified JSON for core fields.
    expect(ctx1.promptContext).not.toContain('"kind":"selection"');
  });

  it("reuses cached sheet samples when only sheet view changes (no extra read_range calls)", async () => {
    const documentController = new DocumentController();
    documentController.setRangeValues("Sheet1", "A1", [["Header"], ["Value"]]);

    const spreadsheet = new DocumentControllerSpreadsheetApi(documentController);
    const readSpy = vi.spyOn(spreadsheet, "readRange");

    const builder = new WorkbookContextBuilder({
      workbookId: "wb_cache",
      documentController,
      spreadsheet,
      ragService: null,
      mode: "chat",
      model: "unit-test-model",
    });

    await builder.build({ activeSheetId: "Sheet1" });
    const readsAfterFirst = readSpy.mock.calls.length;

    // Sheet view only: should not bump DocumentController sheet content versions, so context cache stays valid.
    documentController.setFrozen("Sheet1", 1, 0);

    await builder.build({ activeSheetId: "Sheet1" });
    expect(readSpy.mock.calls.length).toBe(readsAfterFirst);

    readSpy.mockRestore();
  });

  it("keeps other sheets cached when only one sheet changes (per-sheet content versioning)", async () => {
    const documentController = new DocumentController();
    documentController.setRangeValues("Sheet1", "A1", [
      ["Name", "Age"],
      ["Alice", 30],
    ]);
    documentController.setRangeValues("Sheet2", "A1", [
      ["Product", "Price"],
      ["Widget", 10],
    ]);

    const spreadsheet = new DocumentControllerSpreadsheetApi(documentController);

    const readRangeCalls: Record<string, number> = {};
    const originalReadRange = spreadsheet.readRange.bind(spreadsheet);
    const spy = vi.spyOn(spreadsheet, "readRange").mockImplementation((range: any) => {
      const sheet = String(range?.sheet ?? "");
      readRangeCalls[sheet] = (readRangeCalls[sheet] ?? 0) + 1;
      return originalReadRange(range);
    });

    const builder = new WorkbookContextBuilder({
      workbookId: "wb_cache_multi_sheet",
      documentController,
      spreadsheet,
      ragService: null,
      mode: "chat",
      model: "unit-test-model",
      maxSheets: 2,
    });

    await builder.build({ activeSheetId: "Sheet1" });
    const firstSheet2Reads = readRangeCalls["Sheet2"] ?? 0;
    expect(firstSheet2Reads).toBeGreaterThan(0);

    // Mutate only Sheet1.
    documentController.setCellValue("Sheet1", "A2", "Alicia");

    await builder.build({ activeSheetId: "Sheet1" });

    // Sheet2 should be a cache hit: no additional read_range calls.
    expect(readRangeCalls["Sheet2"] ?? 0).toBe(firstSheet2Reads);

    spy.mockRestore();
  });
});

// Note: Prompt context now includes pretty-printed, stable JSON for readability.
