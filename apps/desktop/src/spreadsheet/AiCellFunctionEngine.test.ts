import { describe, expect, it, vi } from "vitest";

import { evaluateFormula } from "./evaluateFormula.js";
import { AI_CELL_DLP_ERROR, AI_CELL_PLACEHOLDER, AiCellFunctionEngine } from "./AiCellFunctionEngine.js";

import { MemoryAIAuditStore } from "../../../../packages/ai-audit/src/memory-store.js";

import { InMemoryAuditLogger } from "../../../../packages/security/dlp/src/audit.js";
import { DLP_ACTION } from "../../../../packages/security/dlp/src/actions.js";
import { CLASSIFICATION_LEVEL } from "../../../../packages/security/dlp/src/classification.js";
import { createDefaultOrgPolicy } from "../../../../packages/security/dlp/src/policy.js";
import { LocalClassificationStore, createMemoryStorage } from "../../../../packages/security/dlp/src/classificationStore.js";
import { CLASSIFICATION_SCOPE } from "../../../../packages/security/dlp/src/selectors.js";

type Deferred<T> = {
  promise: Promise<T>;
  resolve: (value: T) => void;
  reject: (error: unknown) => void;
};

function defer<T>(): Deferred<T> {
  let resolve!: (value: T) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((res, rej) => {
    resolve = res;
    reject = rej;
  });
  return { promise, resolve, reject };
}

describe("AiCellFunctionEngine", () => {
  it("returns #GETTING_DATA while pending and resolves via cache", async () => {
    const deferred = defer<any>();
    const llmClient = {
      chat: vi.fn(() => deferred.promise),
    };

    const auditStore = new MemoryAIAuditStore();
    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      model: "test-model",
      sessionId: "test-session",
      auditStore,
    });

    const pending = evaluateFormula('=AI("summarize", "hello")', () => null, { ai: engine, cellAddress: "Sheet1!A1" });
    expect(pending).toBe(AI_CELL_PLACEHOLDER);
    expect(llmClient.chat).toHaveBeenCalledTimes(1);

    deferred.resolve({
      message: { role: "assistant", content: "ok" },
      usage: { promptTokens: 3, completionTokens: 7 },
    });
    await engine.waitForIdle();

    const resolved = evaluateFormula('=AI("summarize", "hello")', () => null, { ai: engine, cellAddress: "Sheet1!A1" });
    expect(resolved).toBe("ok");
    expect(llmClient.chat).toHaveBeenCalledTimes(1);

    const entries = await auditStore.listEntries({ session_id: "test-session" });
    expect(entries).toHaveLength(1);
    expect(entries[0]?.mode).toBe("cell_function");
    expect((entries[0]?.input as any)?.prompt).toBe("summarize");
    expect((entries[0]?.input as any)?.inputs_hash).toMatch(/^[0-9a-f]{8}$/);
    expect(entries[0]?.token_usage).toEqual({
      prompt_tokens: 3,
      completion_tokens: 7,
      total_tokens: 10,
    });
  });

  it("cache hit avoids re-calling the LLM", async () => {
    const llmClient = {
      chat: vi.fn(async () => ({
        message: { role: "assistant", content: "cached" },
        usage: { promptTokens: 1, completionTokens: 1 },
      })),
    };

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
    });

    const first = evaluateFormula('=AI("summarize", "hello")', () => null, { ai: engine, cellAddress: "Sheet1!A1" });
    expect(first).toBe(AI_CELL_PLACEHOLDER);
    await engine.waitForIdle();

    const second = evaluateFormula('=AI("summarize", "hello")', () => null, { ai: engine, cellAddress: "Sheet1!A1" });
    expect(second).toBe("cached");
    expect(llmClient.chat).toHaveBeenCalledTimes(1);
  });

  it("changing referenced cells invalidates the cache key", async () => {
    const deferred1 = defer<any>();
    const deferred2 = defer<any>();

    let callCount = 0;
    const llmClient = {
      chat: vi.fn(() => {
        callCount += 1;
        return callCount === 1 ? deferred1.promise : deferred2.promise;
      }),
    };

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
    });

    let a1: any = "hello";
    const getCellValue = (addr: string) => (addr === "A1" ? a1 : null);

    const pending1 = evaluateFormula('=AI("summarize", A1)', getCellValue, { ai: engine, cellAddress: "Sheet1!B1" });
    expect(pending1).toBe(AI_CELL_PLACEHOLDER);

    deferred1.resolve({ message: { role: "assistant", content: "first" }, usage: { promptTokens: 1, completionTokens: 1 } });
    await engine.waitForIdle();
    expect(evaluateFormula('=AI("summarize", A1)', getCellValue, { ai: engine, cellAddress: "Sheet1!B1" })).toBe("first");

    // Change the referenced cell value -> new inputs hash -> new request.
    a1 = "goodbye";
    const pending2 = evaluateFormula('=AI("summarize", A1)', getCellValue, { ai: engine, cellAddress: "Sheet1!B1" });
    expect(pending2).toBe(AI_CELL_PLACEHOLDER);
    expect(llmClient.chat).toHaveBeenCalledTimes(2);

    deferred2.resolve({ message: { role: "assistant", content: "second" }, usage: { promptTokens: 1, completionTokens: 1 } });
    await engine.waitForIdle();
    expect(evaluateFormula('=AI("summarize", A1)', getCellValue, { ai: engine, cellAddress: "Sheet1!B1" })).toBe("second");
  });

  it("DLP blocks disallowed inputs and emits a deterministic cell error", () => {
    const llmClient = { chat: vi.fn() };
    const dlpAudit = new InMemoryAuditLogger();

    const policy = createDefaultOrgPolicy();
    // Force strict behavior: block instead of redacting.
    policy.rules[DLP_ACTION.AI_CLOUD_PROCESSING] = {
      ...policy.rules[DLP_ACTION.AI_CLOUD_PROCESSING],
      redactDisallowed: false,
    };

    const documentId = "unit-test-doc";
    const storage = createMemoryStorage();
    const classificationStore = new LocalClassificationStore({ storage });
    classificationStore.upsert(
      documentId,
      { scope: CLASSIFICATION_SCOPE.CELL, documentId, sheetId: "Sheet1", row: 0, col: 0 },
      { level: CLASSIFICATION_LEVEL.RESTRICTED, labels: [] },
    );

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      dlp: {
        policy,
        auditLogger: dlpAudit,
        documentId,
        classificationStore,
        classify: () => ({ level: CLASSIFICATION_LEVEL.PUBLIC, labels: [] }),
      },
    });

    const value = evaluateFormula('=AI("summarize", A1)', (ref) => (ref === "A1" ? "top secret" : null), {
      ai: engine,
      cellAddress: "Sheet1!B1",
    });
    expect(value).toBe(AI_CELL_DLP_ERROR);
    expect(llmClient.chat).not.toHaveBeenCalled();

    const events = dlpAudit.list();
    expect(events.some((e: any) => e.details?.type === "ai.cell_function")).toBe(true);
  });

  it("DLP does not allow restricted cells to be smuggled via nested formulas (e.g. IF)", () => {
    const llmClient = { chat: vi.fn() };

    const policy = createDefaultOrgPolicy();
    policy.rules[DLP_ACTION.AI_CLOUD_PROCESSING] = {
      ...policy.rules[DLP_ACTION.AI_CLOUD_PROCESSING],
      redactDisallowed: false,
    };

    const documentId = "unit-test-doc";
    const storage = createMemoryStorage();
    const classificationStore = new LocalClassificationStore({ storage });
    classificationStore.upsert(
      documentId,
      { scope: CLASSIFICATION_SCOPE.CELL, documentId, sheetId: "Sheet1", row: 0, col: 0 },
      { level: CLASSIFICATION_LEVEL.RESTRICTED, labels: [] },
    );

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      dlp: {
        policy,
        documentId,
        classificationStore,
        classify: () => ({ level: CLASSIFICATION_LEVEL.PUBLIC, labels: [] }),
      },
    });

    const value = evaluateFormula('=AI("summarize", IF(TRUE, A1, "x"))', (ref) => (ref === "A1" ? "top secret" : null), {
      ai: engine,
      cellAddress: "Sheet1!B1",
    });
    expect(value).toBe(AI_CELL_DLP_ERROR);
    expect(llmClient.chat).not.toHaveBeenCalled();
  });

  it("DLP does not allow restricted cells to be smuggled via arithmetic coercion (e.g. A1+0)", () => {
    const llmClient = { chat: vi.fn() };

    const policy = createDefaultOrgPolicy();
    policy.rules[DLP_ACTION.AI_CLOUD_PROCESSING] = {
      ...policy.rules[DLP_ACTION.AI_CLOUD_PROCESSING],
      redactDisallowed: false,
    };

    const documentId = "unit-test-doc";
    const storage = createMemoryStorage();
    const classificationStore = new LocalClassificationStore({ storage });
    classificationStore.upsert(
      documentId,
      { scope: CLASSIFICATION_SCOPE.CELL, documentId, sheetId: "Sheet1", row: 0, col: 0 },
      { level: CLASSIFICATION_LEVEL.RESTRICTED, labels: [] },
    );

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      dlp: {
        policy,
        documentId,
        classificationStore,
        classify: () => ({ level: CLASSIFICATION_LEVEL.PUBLIC, labels: [] }),
      },
    });

    const value = evaluateFormula('=AI("summarize", A1+0)', (ref) => (ref === "A1" ? 123 : null), {
      ai: engine,
      cellAddress: "Sheet1!B1",
    });
    expect(value).toBe(AI_CELL_DLP_ERROR);
    expect(llmClient.chat).not.toHaveBeenCalled();
  });

  it("DLP does not allow restricted cells to be smuggled via derived aggregations (e.g. SUM(range))", () => {
    const llmClient = { chat: vi.fn() };

    const policy = createDefaultOrgPolicy();
    policy.rules[DLP_ACTION.AI_CLOUD_PROCESSING] = {
      ...policy.rules[DLP_ACTION.AI_CLOUD_PROCESSING],
      redactDisallowed: false,
    };

    const documentId = "unit-test-doc";
    const storage = createMemoryStorage();
    const classificationStore = new LocalClassificationStore({ storage });
    classificationStore.upsert(
      documentId,
      { scope: CLASSIFICATION_SCOPE.CELL, documentId, sheetId: "Sheet1", row: 0, col: 0 },
      { level: CLASSIFICATION_LEVEL.RESTRICTED, labels: [] },
    );

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      dlp: {
        policy,
        documentId,
        classificationStore,
        classify: () => ({ level: CLASSIFICATION_LEVEL.PUBLIC, labels: [] }),
      },
    });

    const value = evaluateFormula(
      '=AI("summarize", SUM(A1:A2))',
      (ref) => (ref === "A1" ? 1 : ref === "A2" ? 2 : null),
      {
        ai: engine,
        cellAddress: "Sheet1!B1",
      },
    );
    expect(value).toBe(AI_CELL_DLP_ERROR);
    expect(llmClient.chat).not.toHaveBeenCalled();
  });

  it("DLP redacts inputs before sending to the LLM", async () => {
    const llmClient = {
      chat: vi.fn(async () => ({
        message: { role: "assistant", content: "ok" },
        usage: { promptTokens: 1, completionTokens: 1 },
      })),
    };

    const policy = createDefaultOrgPolicy();
    const documentId = "unit-test-doc";
    const storage = createMemoryStorage();
    const classificationStore = new LocalClassificationStore({ storage });
    classificationStore.upsert(
      documentId,
      { scope: CLASSIFICATION_SCOPE.CELL, documentId, sheetId: "Sheet1", row: 0, col: 0 },
      { level: CLASSIFICATION_LEVEL.RESTRICTED, labels: ["test"] },
    );

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      dlp: {
        policy,
        documentId,
        classificationStore,
        classify: () => ({ level: CLASSIFICATION_LEVEL.PUBLIC, labels: [] }),
      },
    });

    const pending = evaluateFormula('=AI("summarize", A1)', (ref) => (ref === "A1" ? "secret payload" : null), {
      ai: engine,
      cellAddress: "Sheet1!B1",
    });
    expect(pending).toBe(AI_CELL_PLACEHOLDER);
    await engine.waitForIdle();

    const call = llmClient.chat.mock.calls[0]?.[0];
    const userMessage = call?.messages?.find((m: any) => m.role === "user")?.content ?? "";
    expect(userMessage).toContain("[REDACTED]");
    expect(userMessage).not.toContain("secret payload");
  });

  it("propagates spreadsheet error codes from referenced cells", () => {
    const llmClient = { chat: vi.fn() };
    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
    });

    const value = evaluateFormula('=AI("summarize", A1)', (ref) => (ref === "A1" ? "#DIV/0!" : null), {
      ai: engine,
      cellAddress: "Sheet1!B1",
    });
    expect(value).toBe("#DIV/0!");
    expect(llmClient.chat).not.toHaveBeenCalled();
  });

  it("truncates large range inputs before sending to the LLM", async () => {
    const llmClient = {
      chat: vi.fn(async () => ({
        message: { role: "assistant", content: "ok" },
        usage: { promptTokens: 1, completionTokens: 1 },
      })),
    };

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      limits: { maxInputCells: 50 },
    });

    const getCellValue = (addr: string) => {
      const idx = Number(addr.slice(1));
      return `CELL_${String(idx).padStart(4, "0")}`;
    };

    const pending = evaluateFormula('=AI("summarize", A1:A200)', getCellValue, { ai: engine, cellAddress: "Sheet1!B1" });
    expect(pending).toBe(AI_CELL_PLACEHOLDER);
    await engine.waitForIdle();

    const call = llmClient.chat.mock.calls[0]?.[0];
    const userMessage = call?.messages?.find((m: any) => m.role === "user")?.content ?? "";
    expect(userMessage).toContain('"truncated":true');
    expect(userMessage).toContain('"total_cells":200');
    expect(userMessage).toContain('"sampled_cells":50');
    expect(userMessage).toContain("CELL_0001");
    expect(userMessage).not.toContain("CELL_0200");
  });
});
