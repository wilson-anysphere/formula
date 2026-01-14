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
import type { RangeAddress } from "../../../../../../packages/ai-tools/src/spreadsheet/a1.js";
import type { CellData } from "../../../../../../packages/ai-tools/src/spreadsheet/types.js";

import { DLP_ACTION } from "../../../../../../packages/security/dlp/src/actions.js";
import { createDefaultOrgPolicy } from "../../../../../../packages/security/dlp/src/policy.js";
import { CLASSIFICATION_LEVEL } from "../../../../../../packages/security/dlp/src/classification.js";
import { LocalClassificationStore, createMemoryStorage } from "../../../../../../packages/security/dlp/src/classificationStore.js";

import { DocumentControllerSpreadsheetApi } from "../../tools/documentControllerSpreadsheetApi.js";
import { WorkbookContextBuilder } from "../WorkbookContextBuilder.js";

class CountingSpreadsheetApi extends DocumentControllerSpreadsheetApi {
  readonly readRangeCalls: RangeAddress[] = [];

  override readRange(range: RangeAddress): CellData[][] {
    this.readRangeCalls.push(range);
    this.onReadRange?.(this.readRangeCalls.length, range);
    return super.readRange(range);
  }

  constructor(
    controller: DocumentController,
    private readonly onReadRange?: (count: number, range: RangeAddress) => void,
  ) {
    super(controller);
  }
}

describe("WorkbookContextBuilder", () => {
  beforeEach(() => {
    toolExecutorConstructorCalls = 0;
  });

  it("can include computed formula values in data blocks when includeFormulaValues is enabled", async () => {
    const documentController = new DocumentController();
    // Simulate an imported workbook cell that contains both a formula and a cached/computed value.
    documentController.model.setCell("Sheet1", 0, 0, { value: 2, formula: "=1+1", styleId: 0 });

    const spreadsheet = new DocumentControllerSpreadsheetApi(documentController);
    const builder = new WorkbookContextBuilder({
      workbookId: "wb_formula_values",
      documentController,
      spreadsheet,
      ragService: null,
      mode: "chat",
      model: "unit-test-model",
      includeFormulaValues: true,
    });

    const ctx = await builder.build({
      activeSheetId: "Sheet1",
      selectedRange: { sheetId: "Sheet1", range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 } },
    });

    const selectionBlock = ctx.payload.blocks.find((b) => b.kind === "selection");
    expect(selectionBlock).toBeTruthy();
    expect(selectionBlock?.values).toEqual([[2]]);
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

    const onBuildStats = vi.fn();
    const builder = new WorkbookContextBuilder({
      workbookId: "wb_rag_builder",
      documentController,
      spreadsheet,
      ragService: ragService as any,
      mode: "chat",
      model: "unit-test-model",
      maxPromptContextTokens: 4000,
      onBuildStats,
    });

    const ctx = await builder.build({ activeSheetId: "Sheet1", focusQuestion: "revenue by region" });

    expect(lastRagResult).toBeTruthy();
    expect(lastRagResult.promptContext).toBe("");
    expect(ctx.retrieved.length).toBeGreaterThan(0);

    expect(onBuildStats).toHaveBeenCalledTimes(1);
    const stats = onBuildStats.mock.calls[0]![0];
    expect(stats.rag.enabled).toBe(true);
    expect(stats.rag.retrievedCount).toBe(ctx.retrieved.length);
    expect(stats.rag.retrievedBlockCount).toBeGreaterThan(0);

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

  it("shares schemaProvider metadata cache across builders when schemaVersion is stable", async () => {
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

    const schemaProvider = {
      getSchemaVersion: () => schemaVersion,
      getNamedRanges,
      getTables,
    };

    const builder1 = new WorkbookContextBuilder({
      workbookId: "wb_schema_cache_shared",
      documentController,
      spreadsheet,
      ragService: null,
      mode: "chat",
      model: "unit-test-model",
      schemaProvider,
    });
    await builder1.build({ activeSheetId: "Sheet1" });
    expect(getNamedRanges).toHaveBeenCalledTimes(1);
    expect(getTables).toHaveBeenCalledTimes(1);

    // New builder instance, same schemaVersion -> shared cache should avoid provider reads.
    const builder2 = new WorkbookContextBuilder({
      workbookId: "wb_schema_cache_shared",
      documentController,
      spreadsheet,
      ragService: null,
      mode: "chat",
      model: "unit-test-model",
      schemaProvider,
    });
    await builder2.build({ activeSheetId: "Sheet1" });
    expect(getNamedRanges).toHaveBeenCalledTimes(1);
    expect(getTables).toHaveBeenCalledTimes(1);

    // Version bump should invalidate the shared cache.
    schemaVersion += 1;
    const builder3 = new WorkbookContextBuilder({
      workbookId: "wb_schema_cache_shared",
      documentController,
      spreadsheet,
      ragService: null,
      mode: "chat",
      model: "unit-test-model",
      schemaProvider,
    });
    await builder3.build({ activeSheetId: "Sheet1" });
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

  it("does not reuse cached unredacted blocks when DLP settings tighten (cache is keyed by DLP state)", async () => {
    const workbookId = "wb_dlp_cache_key";
    const documentController = new DocumentController();
    documentController.setRangeValues("Sheet1", "A1", [
      ["Public"],
      ["TOP SECRET"],
    ]);

    const spreadsheet = new DocumentControllerSpreadsheetApi(documentController);
    const builder = new WorkbookContextBuilder({
      workbookId,
      documentController,
      spreadsheet,
      ragService: null,
      mode: "chat",
      model: "unit-test-model",
      maxSheets: 1,
    });

    // Build once with no DLP: caches will contain unredacted values.
    const ctx1 = await builder.build({ activeSheetId: "Sheet1" });
    expect(ctx1.promptContext).toContain("TOP SECRET");

    // Now apply DLP that redacts the restricted cell.
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

    const ctx2 = await builder.build({ activeSheetId: "Sheet1", dlp });
    expect(ctx2.promptContext).toContain("[REDACTED]");
    expect(ctx2.promptContext).not.toContain("TOP SECRET");
  });

  it("does not reuse cached blocks when classification records change without updatedAt fields (cache key includes selector/classification)", async () => {
    const workbookId = "wb_dlp_cache_key_records";
    const documentController = new DocumentController();
    documentController.setRangeValues("Sheet1", "A1", [
      ["Public"],
      ["TOP SECRET"],
    ]);

    const spreadsheet = new DocumentControllerSpreadsheetApi(documentController);
    const builder = new WorkbookContextBuilder({
      workbookId,
      documentController,
      spreadsheet,
      ragService: null,
      mode: "inline_edit",
      model: "unit-test-model",
      maxSheets: 1,
    });

    const policy = createDefaultOrgPolicy();
    const auditLogger = { log: (_event: any) => {} };

    const selector = { scope: "cell", documentId: workbookId, sheetId: "Sheet1", row: 1, col: 0 };

    const publicRecord = { selector, classification: { level: CLASSIFICATION_LEVEL.PUBLIC, labels: [] } };
    const restrictedRecord = { selector, classification: { level: CLASSIFICATION_LEVEL.RESTRICTED, labels: ["test"] } };

    const dlpBase = {
      // ContextManager style
      documentId: workbookId,
      sheetId: "Sheet1",
      policy,
      includeRestrictedContent: false,
      auditLogger,
      // ToolExecutor style
      document_id: workbookId,
      sheet_id: "Sheet1",
      include_restricted_content: false,
      audit_logger: auditLogger,
    };

    const dlpPublic = {
      ...dlpBase,
      classificationRecords: [publicRecord],
      classification_records: [publicRecord],
    };

    const ctx1 = await builder.build({
      activeSheetId: "Sheet1",
      selectedRange: { sheetId: "Sheet1", range: { startRow: 0, endRow: 1, startCol: 0, endCol: 0 } },
      dlp: dlpPublic,
    });
    const selection1 = ctx1.payload.blocks.find((b) => b.kind === "selection");
    expect(selection1).toBeTruthy();
    expect(selection1!.values[1]![0]).toBe("TOP SECRET");

    // Now tighten classification to Restricted (no updatedAt field; cache key must still change).
    const dlpRestricted = {
      ...dlpBase,
      classificationRecords: [restrictedRecord],
      classification_records: [restrictedRecord],
    };

    const ctx2 = await builder.build({
      activeSheetId: "Sheet1",
      selectedRange: { sheetId: "Sheet1", range: { startRow: 0, endRow: 1, startCol: 0, endCol: 0 } },
      dlp: dlpRestricted,
    });
    const selection2 = ctx2.payload.blocks.find((b) => b.kind === "selection");
    expect(selection2).toBeTruthy();
    expect(selection2!.values[1]![0]).toBe("[REDACTED]");
    expect(ctx2.promptContext).not.toContain("TOP SECRET");
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

  it("reads the active sheet only once on a cache miss by reusing the schema sample for the prompt sample", async () => {
    const documentController = new DocumentController();
    // Exceed the default maxBlockRows (20) so the schema window (maxSchemaRows) differs
    // from the prompt window, which used to trigger a second read_range call.
    const rows = Array.from({ length: 25 }, (_v, rIdx) =>
      Array.from({ length: 5 }, (_v2, cIdx) => `R${rIdx + 1}C${cIdx + 1}`),
    );
    documentController.setRangeValues("Sheet1", "A1", rows);

    const spreadsheet = new DocumentControllerSpreadsheetApi(documentController);
    const originalReadRange = spreadsheet.readRange.bind(spreadsheet);
    let readRangeCalls = 0;
    const spy = vi.spyOn(spreadsheet, "readRange").mockImplementation((range: any) => {
      readRangeCalls += 1;
      return originalReadRange(range);
    });

    const builder = new WorkbookContextBuilder({
      workbookId: "wb_single_read",
      documentController,
      spreadsheet,
      ragService: null,
      mode: "chat",
      model: "unit-test-model",
    });

    const ctx = await builder.build({ activeSheetId: "Sheet1" });
    expect(readRangeCalls).toBe(1);

    const sample = ctx.payload.blocks.find((b) => b.kind === "sheet_sample" && b.sheetId === "Sheet1");
    expect(sample).toBeTruthy();
    expect(sample!.range).toBe("Sheet1!A1:E20");

    spy.mockRestore();
  });

  it("cancels promptly when aborted during a read_range call (no extra reads after abort)", async () => {
    const documentController = new DocumentController();
    documentController.setRangeValues("Sheet1", "A1", [
      ["A"],
      ["B"],
    ]);
    documentController.setRangeValues("Sheet2", "A1", [
      ["C"],
      ["D"],
    ]);

    const spreadsheet = new DocumentControllerSpreadsheetApi(documentController);
    const abortController = new AbortController();
    const signal = abortController.signal;

    const originalReadRange = spreadsheet.readRange.bind(spreadsheet);
    let readRangeCalls = 0;
    const readRangeSpy = vi.spyOn(spreadsheet, "readRange").mockImplementation((range: any) => {
      readRangeCalls += 1;
      if (readRangeCalls === 1) abortController.abort();
      return originalReadRange(range);
    });

    const builder = new WorkbookContextBuilder({
      workbookId: "wb_abort",
      documentController,
      spreadsheet,
      ragService: null,
      mode: "chat",
      model: "unit-test-model",
      maxSheets: 2,
    });

    await expect(builder.build({ activeSheetId: "Sheet1", signal })).rejects.toMatchObject({ name: "AbortError" });
    expect(readRangeCalls).toBe(1);

    // Give any stray background work a chance to schedule, then assert we didn't keep reading.
    await new Promise((resolve) => setTimeout(resolve, 10));
    expect(readRangeCalls).toBe(1);

    readRangeSpy.mockRestore();
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

  it("invokes onBuildStats with sane counters", async () => {
    const documentController = new DocumentController();
    documentController.setRangeValues("Sheet1", "A1", [["hello"]]);

    const spreadsheet = new DocumentControllerSpreadsheetApi(documentController);
    const onBuildStats = vi.fn();
    const builder = new WorkbookContextBuilder({
      workbookId: "wb_stats",
      documentController,
      spreadsheet,
      ragService: null,
      mode: "chat",
      model: "unit-test-model",
      maxSheets: 1,
      onBuildStats,
    });

    const ctx1 = await builder.build({ activeSheetId: "Sheet1" });
    const ctx2 = await builder.build({ activeSheetId: "Sheet1" });

    expect(onBuildStats).toHaveBeenCalledTimes(2);
    const first = onBuildStats.mock.calls[0]![0];
    const second = onBuildStats.mock.calls[1]![0];

    expect(first.mode).toBe("chat");
    expect(first.model).toBe("unit-test-model");
    expect(first.durationMs).toBeGreaterThanOrEqual(0);
    expect(first.ok).toBe(true);
    expect(first.sheetCountSummarized).toBe(1);
    expect(first.blockCount).toBe(1);
    expect(first.blockCountByKind).toEqual({ selection: 0, sheet_sample: 1, retrieved: 0 });
    expect(first.blockCellCount).toBeGreaterThan(0);
    expect(first.blockCellCountByKind.sheet_sample).toBeGreaterThan(0);
    expect(first.promptContextChars).toBe(ctx1.promptContext.length);
    expect(first.promptContextTokens).toBe(ctx1.payload.budget.usedPromptContextTokens);
    expect(first.promptContextBudgetTokens).toBe(ctx1.payload.budget.maxPromptContextTokens);
    expect(first.promptContextTrimmedSectionCount).toBeGreaterThanOrEqual(0);
    expect(first.cache.schema.misses).toBeGreaterThanOrEqual(1);
    expect(first.cache.block.misses).toBeGreaterThanOrEqual(1);
    expect(first.cache.schema.entries).toBeGreaterThanOrEqual(1);
    expect(first.cache.block.entries).toBeGreaterThanOrEqual(1);
    expect(first.cache.block.entriesByKind.sheet_sample).toBeGreaterThanOrEqual(1);
    expect(first.readBlockCellCount).toBeGreaterThanOrEqual(1);
    expect(first.readBlockCellCountByKind.sheet_sample).toBeGreaterThanOrEqual(1);
    expect(first.rag.enabled).toBe(false);
    expect(first.rag.retrievedCount).toBe(0);

    // Second build should reuse cached schema + sampled blocks.
    expect(second.cache.schema.hits).toBeGreaterThanOrEqual(1);
    expect(second.cache.block.hits).toBeGreaterThanOrEqual(1);
    expect(second.ok).toBe(true);
    expect(second.cache.schema.entries).toBeGreaterThanOrEqual(1);
    expect(second.cache.block.entries).toBeGreaterThanOrEqual(1);
    expect(second.cache.block.entriesByKind.sheet_sample).toBeGreaterThanOrEqual(1);
    expect(second.readBlockCellCount).toBe(0);
    expect(second.readBlockCellCountByKind.sheet_sample).toBe(0);
    expect(second.blockCountByKind).toEqual({ selection: 0, sheet_sample: 1, retrieved: 0 });
    expect(second.blockCellCount).toBeGreaterThan(0);
    expect(second.blockCellCountByKind.sheet_sample).toBeGreaterThan(0);
    expect(second.promptContextChars).toBe(ctx2.promptContext.length);
    expect(second.promptContextTokens).toBe(ctx2.payload.budget.usedPromptContextTokens);
    expect(second.promptContextBudgetTokens).toBe(ctx2.payload.budget.maxPromptContextTokens);
    expect(second.promptContextTrimmedSectionCount).toBeGreaterThanOrEqual(0);
  });

  it("invokes onBuildStats when build throws (rag error)", async () => {
    const documentController = new DocumentController();
    documentController.setRangeValues("Sheet1", "A1", [["A"]]);

    const spreadsheet = new DocumentControllerSpreadsheetApi(documentController);
    const onBuildStats = vi.fn();
    const ragService = {
      async buildWorkbookContextFromSpreadsheetApi() {
        throw new Error("boom");
      },
    };

    const builder = new WorkbookContextBuilder({
      workbookId: "wb_stats_error",
      documentController,
      spreadsheet,
      ragService: ragService as any,
      mode: "chat",
      model: "unit-test-model",
      maxSheets: 1,
      onBuildStats,
    });

    await expect(builder.build({ activeSheetId: "Sheet1", focusQuestion: "test query" })).rejects.toThrow(/boom/);

    expect(onBuildStats).toHaveBeenCalledTimes(1);
    const stats = onBuildStats.mock.calls[0]![0];
    expect(stats.ok).toBe(false);
    expect(stats.error?.message).toMatch(/boom/);
    expect(stats.rag.enabled).toBe(true);
  });

  it("builds a deterministic promptContext with compact stable JSON", async () => {
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

    // Prompt context should contain stable, compact JSON (token-efficient + machine-readable).
    expect(ctx1.promptContext).not.toContain('\n  "');
    expect(ctx1.promptContext).toContain('"kind":"selection"');
    // Ensure we don't regress to pretty JSON for core fields.
    expect(ctx1.promptContext).not.toContain('"kind": "selection"');
  });

  it("respects a custom tokenEstimator when packing promptContext sections", async () => {
    const documentController = new DocumentController();
    documentController.setRangeValues("Sheet1", "A1", [
      ["Name", "Age"],
      ["Alice", 30],
      ["Bob", 40],
    ]);

    const spreadsheet = new DocumentControllerSpreadsheetApi(documentController);
    const input = {
      activeSheetId: "Sheet1",
      selectedRange: { sheetId: "Sheet1", range: { startRow: 0, endRow: 2, startCol: 0, endCol: 1 } },
    };

    const onBuildStatsDefault = vi.fn();
    // With the default heuristic estimator, this should fit without trimming.
    const builderDefault = new WorkbookContextBuilder({
      workbookId: "wb_estimator_default",
      documentController,
      spreadsheet,
      ragService: null,
      mode: "chat",
      model: "unit-test-model",
      // Large enough to be resilient to future schema/payload expansions.
      maxPromptContextTokens: 2_000,
      onBuildStats: onBuildStatsDefault,
    });
    const ctxDefault = await builderDefault.build(input);
    expect(ctxDefault.promptContext).not.toContain("trimmed to fit token budget");
    expect(onBuildStatsDefault).toHaveBeenCalledTimes(1);
    expect(onBuildStatsDefault.mock.calls[0]![0].promptContextTrimmedSectionCount).toBe(0);

    // With a much stricter estimator, the exact same context should be trimmed.
    const onBuildStatsStrict = vi.fn();
    const strictEstimator = {
      estimateTextTokens: (text: string) => text.length * 10,
      estimateMessageTokens: (_message: any) => 0,
      estimateMessagesTokens: (_messages: any[]) => 0,
    };
    const builderStrict = new WorkbookContextBuilder({
      workbookId: "wb_estimator_strict",
      documentController,
      spreadsheet,
      ragService: null,
      mode: "chat",
      model: "unit-test-model",
      maxPromptContextTokens: 2_000,
      tokenEstimator: strictEstimator as any,
      onBuildStats: onBuildStatsStrict,
    });
    const ctxStrict = await builderStrict.build(input);
    expect(ctxStrict.promptContext).toContain("trimmed to fit token budget");
    expect(onBuildStatsStrict).toHaveBeenCalledTimes(1);
    expect(onBuildStatsStrict.mock.calls[0]![0].promptContextTrimmedSectionCount).toBeGreaterThan(0);
  });

  it("reuses cached sheet summaries + blocks when the workbook hasn't changed, and invalidates on content edits", async () => {
    const documentController = new DocumentController();
    documentController.setRangeValues("Sheet1", "A1", [
      ["Name", "Age"],
      ["Alice", 30],
    ]);

    const spreadsheet = new DocumentControllerSpreadsheetApi(documentController);
    const readSpy = vi.spyOn(spreadsheet, "readRange");

    const builder = new WorkbookContextBuilder({
      workbookId: "wb_cache_basic",
      documentController,
      spreadsheet,
      ragService: null,
      mode: "chat",
      model: "unit-test-model",
      maxSheets: 1,
    });

    await builder.build({ activeSheetId: "Sheet1" });
    const readsAfterFirst = readSpy.mock.calls.length;

    await builder.build({ activeSheetId: "Sheet1" });
    expect(readSpy.mock.calls.length).toBe(readsAfterFirst);

    // Content edit should bump sheet content version -> cache miss -> another read_range.
    documentController.setCellValue("Sheet1", "A2", "Alicia");

    await builder.build({ activeSheetId: "Sheet1" });
    expect(readSpy.mock.calls.length).toBeGreaterThan(readsAfterFirst);

    readSpy.mockRestore();
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
      maxSheets: 1,
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

  it("includes retrieved sheets even when they would be excluded by maxSheets capping", async () => {
    const documentController = new DocumentController();

    const sheetIds = Array.from({ length: 12 }, (_v, idx) => `Sheet${String(idx + 1).padStart(2, "0")}`);
    for (const sheetId of sheetIds) {
      documentController.setRangeValues(sheetId, "A1", [["Value"], [sheetId]]);
    }

    const spreadsheet = new DocumentControllerSpreadsheetApi(documentController);
    const builder = new WorkbookContextBuilder({
      workbookId: "wb_retrieval_sheet_selection",
      documentController,
      spreadsheet,
      mode: "chat",
      model: "unit-test-model",
      maxSheets: 3,
      ragService: {
        buildWorkbookContextFromSpreadsheetApi: async () => ({
          retrieved: [
            {
              id: "chunk_sheet12",
              text: "Sheet12 has the value.",
              metadata: { sheetName: "Sheet12", rect: { r0: 0, c0: 0, r1: 1, c1: 0 } },
            },
          ],
        }),
      },
    });

    // Without retrieval, maxSheets=3 would only include Sheet01..Sheet03 (plus active).
    // With retrieval, the referenced sheet should be summarized even if it sorts outside the cap.
    const ctx = await builder.build({ activeSheetId: "Sheet01", focusQuestion: "what is on Sheet12?" });

    const summarizedSheetIds = ctx.payload.sheets.map((s) => s.sheetId);
    expect(summarizedSheetIds).toContain("Sheet01");
    expect(summarizedSheetIds).toContain("Sheet12");
  });

  it("does not read schema samples for unrelated sheets when retrieval narrows to specific sheets", async () => {
    const documentController = new DocumentController();

    const sheetIds = Array.from({ length: 15 }, (_v, idx) => `Sheet${String(idx + 1).padStart(2, "0")}`);
    for (const sheetId of sheetIds) {
      documentController.setRangeValues(sheetId, "A1", [["Value"], [sheetId]]);
    }

    const spreadsheet = new CountingSpreadsheetApi(documentController);
    const builder = new WorkbookContextBuilder({
      workbookId: "wb_retrieval_perf",
      documentController,
      spreadsheet,
      mode: "chat",
      model: "unit-test-model",
      maxSheets: 10,
      // Keep the test focused on schema sampling rather than extra retrieved block reads.
      maxRetrievedBlocks: 0,
      ragService: {
        buildWorkbookContextFromSpreadsheetApi: async () => ({
          retrieved: [
            { id: "chunk_sheet10", text: "Relevant", metadata: { sheetName: "Sheet10", rect: { r0: 0, c0: 0, r1: 1, c1: 0 } } },
            { id: "chunk_sheet11", text: "Relevant", metadata: { sheetName: "Sheet11", rect: { r0: 0, c0: 0, r1: 1, c1: 0 } } },
          ],
        }),
      },
    });

    const ctx = await builder.build({ activeSheetId: "Sheet01", focusQuestion: "compare Sheet10 and Sheet11" });
    expect(ctx.payload.sheets.map((s) => s.sheetId).sort()).toEqual(["Sheet01", "Sheet10", "Sheet11"].sort());

    const sheetsRead = new Set(spreadsheet.readRangeCalls.map((call) => call.sheet));
    expect([...sheetsRead].sort()).toEqual(["Sheet01", "Sheet10", "Sheet11"].sort());
  });

  it("aborts mid-build and stops reading additional ranges", async () => {
    const documentController = new DocumentController();

    // Populate a bunch of sheets so a non-aborted build would attempt many reads.
    for (let i = 1; i <= 50; i++) {
      documentController.setRangeValues(`Sheet${i}`, "A1", [[`v${i}`]]);
    }

    const abortController = new AbortController();
    const spreadsheet = new CountingSpreadsheetApi(documentController, (count) => {
      // Abort during the second range read; the builder should stop promptly and
      // avoid reading the rest of the sheets.
      if (count === 2) abortController.abort();
    });

    const builder = new WorkbookContextBuilder({
      workbookId: "wb_abort_mid_build",
      documentController,
      spreadsheet,
      ragService: null,
      mode: "chat",
      model: "unit-test-model",
      maxSheets: 50,
    });

    const promise = builder.build({ activeSheetId: "Sheet1", signal: abortController.signal });
    await expect(promise).rejects.toMatchObject({ name: "AbortError" });
    // We should stop promptly after aborting (and not scan all 50 sheets).
    expect(spreadsheet.readRangeCalls.length).toBeLessThan(10);
  });

  it("passes AbortSignal to RAG and aborts while awaiting retrieval", async () => {
    const documentController = new DocumentController();
    documentController.setRangeValues("Sheet1", "A1", [["Hello"]]);

    const spreadsheet = new DocumentControllerSpreadsheetApi(documentController);

    const abortController = new AbortController();
    let receivedSignal: AbortSignal | undefined;

    const ragService = {
      buildWorkbookContextFromSpreadsheetApi: (params: any) => {
        receivedSignal = params.signal;
        // Never resolve unless aborted.
        return new Promise((_resolve, reject) => {
          const signal: AbortSignal | undefined = params.signal;
          if (!signal) return;
          const onAbort = () => {
            const error = new Error("Aborted");
            error.name = "AbortError";
            reject(error);
          };
          signal.addEventListener("abort", onAbort, { once: true });
          if (signal.aborted) onAbort();
        });
      },
    };

    const builder = new WorkbookContextBuilder({
      workbookId: "wb_rag_abort",
      documentController,
      spreadsheet,
      ragService: ragService as any,
      mode: "chat",
      model: "unit-test-model",
    });

    const promise = builder.build({
      activeSheetId: "Sheet1",
      focusQuestion: "test query",
      signal: abortController.signal,
    });

    expect(receivedSignal).toBe(abortController.signal);
    abortController.abort();
    await expect(promise).rejects.toMatchObject({ name: "AbortError" });
  });
});
