import { describe, expect, it, vi } from "vitest";

import { evaluateFormula } from "./evaluateFormula.js";
import { AI_CELL_DLP_ERROR, AI_CELL_PLACEHOLDER, AiCellFunctionEngine } from "./AiCellFunctionEngine.js";

import { MemoryAIAuditStore } from "../../../../packages/ai-audit/src/memory-store.js";

import { InMemoryAuditLogger } from "../../../../packages/security/dlp/src/audit.js";
import { DLP_ACTION } from "../../../../packages/security/dlp/src/actions.js";
import { CLASSIFICATION_LEVEL } from "../../../../packages/security/dlp/src/classification.js";
import { createDefaultOrgPolicy } from "../../../../packages/security/dlp/src/policy.js";

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

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      dlp: {
        policy,
        auditLogger: dlpAudit,
        classify: () => ({ level: CLASSIFICATION_LEVEL.RESTRICTED, labels: [] }),
      },
    });

    const value = evaluateFormula('=AI("summarize", "top secret")', () => null, { ai: engine, cellAddress: "Sheet1!A1" });
    expect(value).toBe(AI_CELL_DLP_ERROR);
    expect(llmClient.chat).not.toHaveBeenCalled();

    const events = dlpAudit.list();
    expect(events.some((e: any) => e.type === "ai.cell_function")).toBe(true);
  });

  it("DLP redacts inputs before sending to the LLM", async () => {
    const llmClient = {
      chat: vi.fn(async () => ({
        message: { role: "assistant", content: "ok" },
        usage: { promptTokens: 1, completionTokens: 1 },
      })),
    };

    const policy = createDefaultOrgPolicy();

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      dlp: {
        policy,
        classify: (value) =>
          String(value).includes("secret")
            ? { level: CLASSIFICATION_LEVEL.RESTRICTED, labels: ["test"] }
            : { level: CLASSIFICATION_LEVEL.PUBLIC, labels: [] },
      },
    });

    const pending = evaluateFormula('=AI("summarize", "secret payload")', () => null, { ai: engine, cellAddress: "Sheet1!A1" });
    expect(pending).toBe(AI_CELL_PLACEHOLDER);
    await engine.waitForIdle();

    const call = llmClient.chat.mock.calls[0]?.[0];
    const userMessage = call?.messages?.find((m: any) => m.role === "user")?.content ?? "";
    expect(userMessage).toContain("[REDACTED]");
    expect(userMessage).not.toContain("secret payload");
  });
});

