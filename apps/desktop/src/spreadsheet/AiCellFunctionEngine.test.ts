import { beforeEach, describe, expect, it, vi } from "vitest";

import { evaluateFormula } from "./evaluateFormula.js";
import { AI_CELL_DLP_ERROR, AI_CELL_ERROR, AI_CELL_PLACEHOLDER, AiCellFunctionEngine } from "./AiCellFunctionEngine.js";

import { MemoryAIAuditStore } from "../../../../packages/ai-audit/src/memory-store.js";

import { DLP_ACTION } from "../../../../packages/security/dlp/src/actions.js";
import { CLASSIFICATION_LEVEL } from "../../../../packages/security/dlp/src/classification.js";
import { LocalClassificationStore } from "../../../../packages/security/dlp/src/classificationStore.js";
import { createDefaultOrgPolicy } from "../../../../packages/security/dlp/src/policy.js";
import { LocalPolicyStore } from "../../../../packages/security/dlp/src/policyStore.js";
import { getAiDlpAuditLogger, resetAiDlpAuditLoggerForTests } from "../ai/dlp/aiDlp.js";
import { createSheetNameResolverFromIdToNameMap } from "../sheet/sheetNameResolver.js";

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

function markCellRestricted(params: { workbookId: string; sheetId: string; row: number; col: number }): void {
  const storage = globalThis.localStorage as any;
  const classificationStore = new LocalClassificationStore({ storage });
  classificationStore.upsert(
    params.workbookId,
    { scope: "cell", documentId: params.workbookId, sheetId: params.sheetId, row: params.row, col: params.col },
    { level: CLASSIFICATION_LEVEL.RESTRICTED, labels: [] },
  );
}

function setBlockPolicy(workbookId: string): void {
  const storage = globalThis.localStorage as any;
  const policy = createDefaultOrgPolicy();
  policy.rules[DLP_ACTION.AI_CLOUD_PROCESSING] = {
    ...(policy.rules[DLP_ACTION.AI_CLOUD_PROCESSING] as any),
    // Force strict behavior: block instead of redacting.
    redactDisallowed: false,
  };
  const policyStore = new LocalPolicyStore({ storage });
  policyStore.setDocumentPolicy(workbookId, policy);
}

describe("AiCellFunctionEngine", () => {
  beforeEach(() => {
    globalThis.localStorage?.clear();
    resetAiDlpAuditLoggerForTests();
  });

  it("returns #GETTING_DATA while pending and resolves via cache", async () => {
    const deferred = defer<any>();
    const llmClient = {
      chat: vi.fn((_request: any) => deferred.promise),
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
    expect((entries[0]?.input as any)?.inputs_hash).toMatch(/^[0-9a-f]{16}$/);
    expect(entries[0]?.token_usage).toEqual({
      prompt_tokens: 3,
      completion_tokens: 7,
      total_tokens: 10,
    });
  });

  it("aborts LLM requests that exceed the timeout and caches #AI!", async () => {
    vi.useFakeTimers();
    try {
      let observedSignal: AbortSignal | undefined;
      let abortEvents = 0;

      const llmClient = {
        chat: vi.fn((request: any) => {
          observedSignal = request?.signal;
          if (observedSignal && typeof (observedSignal as any).addEventListener === "function") {
          observedSignal.addEventListener("abort", () => {
            abortEvents += 1;
          });
        }
        return new Promise(() => {
          // Never resolve/reject: we rely on the engine timeout + abort.
          });
        }),
      };

      const auditStore = new MemoryAIAuditStore();
      const engine = new AiCellFunctionEngine({
        llmClient: llmClient as any,
        model: "test-model",
        sessionId: "timeout-session",
        auditStore,
        limits: { requestTimeoutMs: 50 },
      });

      const pending = evaluateFormula('=AI("summarize", "hello")', () => null, { ai: engine, cellAddress: "Sheet1!A1" });
      expect(pending).toBe(AI_CELL_PLACEHOLDER);
      expect(llmClient.chat).toHaveBeenCalledTimes(1);
      expect(observedSignal).toBeDefined();
      expect(observedSignal?.aborted).toBe(false);

      vi.advanceTimersByTime(60);
      await engine.waitForIdle();

      expect(observedSignal?.aborted).toBe(true);
      expect(abortEvents).toBeGreaterThan(0);

      const resolved = evaluateFormula('=AI("summarize", "hello")', () => null, { ai: engine, cellAddress: "Sheet1!A1" });
      expect(resolved).toBe(AI_CELL_ERROR);
      expect(llmClient.chat).toHaveBeenCalledTimes(1);

      const entries = await auditStore.listEntries({ session_id: "timeout-session" });
      expect(entries).toHaveLength(1);
      expect((entries[0]?.input as any)?.error).toContain("timed out");
      expect(entries[0]?.user_feedback).toBe("rejected");
    } finally {
      vi.useRealTimers();
    }
  });

  it("times out hung LLM calls that reject only on abort and caches #AI!", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date(0));
    try {
      const requestTimeoutMs = 1_000;
      const llmClient = {
        chat: vi.fn((request: any) => {
          // Simulate a provider call that never resolves unless it is explicitly aborted.
          return new Promise((_resolve, reject) => {
            const signal = request?.signal as AbortSignal | undefined;
            if (!signal) return;
            if (signal.aborted) {
              reject(signal.reason ?? new Error("aborted"));
              return;
            }
            signal.addEventListener(
              "abort",
              () => {
                reject(signal.reason ?? new Error("aborted"));
              },
              { once: true },
            );
          });
        }),
      };

      const engine = new AiCellFunctionEngine({
        llmClient: llmClient as any,
        model: "test-model",
        auditStore: new MemoryAIAuditStore(),
        limits: { requestTimeoutMs },
      });

      const pending = evaluateFormula('=AI("summarize", "hello")', () => null, { ai: engine, cellAddress: "Sheet1!A1" });
      expect(pending).toBe(AI_CELL_PLACEHOLDER);
      expect(llmClient.chat).toHaveBeenCalledTimes(1);

      const firstCall = llmClient.chat.mock.calls[0]?.[0];
      expect(firstCall?.signal).toBeDefined();

      vi.advanceTimersByTime(requestTimeoutMs + 1);
      await engine.waitForIdle();

      expect((firstCall?.signal as AbortSignal | undefined)?.aborted).toBe(true);

      const errored = evaluateFormula('=AI("summarize", "hello")', () => null, { ai: engine, cellAddress: "Sheet1!A1" });
      expect(errored).toBe(AI_CELL_ERROR);
      expect(llmClient.chat).toHaveBeenCalledTimes(1);
    } finally {
      vi.useRealTimers();
    }
  });

  it("cache hit avoids re-calling the LLM", async () => {
    const llmClient = {
      chat: vi.fn(async (_request: any) => ({
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

  it("retries cached #AI! errors after the configured TTL", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date(0));

    try {
      const deferred1 = defer<any>();
      const deferred2 = defer<any>();
      let callCount = 0;
      const llmClient = {
        chat: vi.fn((_request: any) => {
          callCount += 1;
          return callCount === 1 ? deferred1.promise : deferred2.promise;
        }),
      };

      const engine = new AiCellFunctionEngine({
        llmClient: llmClient as any,
        auditStore: new MemoryAIAuditStore(),
        cache: { errorTtlMs: 60_000 },
      });

      const first = evaluateFormula('=AI("summarize", "hello")', () => null, { ai: engine, cellAddress: "Sheet1!A1" });
      expect(first).toBe(AI_CELL_PLACEHOLDER);
      expect(llmClient.chat).toHaveBeenCalledTimes(1);

      deferred1.reject(new Error("transient"));
      await engine.waitForIdle();

      // Immediate reevaluation should stick with the cached error (no tight retry loop).
      const errored = evaluateFormula('=AI("summarize", "hello")', () => null, { ai: engine, cellAddress: "Sheet1!A1" });
      expect(errored).toBe(AI_CELL_ERROR);
      expect(llmClient.chat).toHaveBeenCalledTimes(1);

      // After TTL, the cached error expires and we retry.
      vi.advanceTimersByTime(60_001);
      const retrying = evaluateFormula('=AI("summarize", "hello")', () => null, { ai: engine, cellAddress: "Sheet1!A1" });
      expect(retrying).toBe(AI_CELL_PLACEHOLDER);
      expect(llmClient.chat).toHaveBeenCalledTimes(2);

      deferred2.resolve({ message: { role: "assistant", content: "ok" }, usage: { promptTokens: 1, completionTokens: 1 } });
      await engine.waitForIdle();

      const resolved = evaluateFormula('=AI("summarize", "hello")', () => null, { ai: engine, cellAddress: "Sheet1!A1" });
      expect(resolved).toBe("ok");
    } finally {
      vi.useRealTimers();
    }
  });

  it("bounds concurrent LLM requests and queues additional work", async () => {
    const maxConcurrentRequests = 2;
    const totalRequests = 5;

    const deferreds: Array<Deferred<any>> = [];
    const llmClient = {
      chat: vi.fn((_request: any) => {
        const d = defer<any>();
        deferreds.push(d);
        return d.promise;
      }),
    };

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      model: "test-model",
      limits: { maxConcurrentRequests },
    });

    for (let i = 0; i < totalRequests; i += 1) {
      const value = engine.evaluateAiFunction({
        name: "AI",
        args: [`prompt-${i}`, `input-${i}`],
        cellAddress: "Sheet1!A1",
      });
      expect(value).toBe(AI_CELL_PLACEHOLDER);
    }

    // Only the first `maxConcurrentRequests` should start immediately.
    expect(llmClient.chat).toHaveBeenCalledTimes(maxConcurrentRequests);
    expect(deferreds).toHaveLength(maxConcurrentRequests);

    // Completing an in-flight request should kick off the next queued request.
    deferreds[0]!.resolve({
      message: { role: "assistant", content: "r0" },
      usage: { promptTokens: 1, completionTokens: 1 },
    });
    await Promise.resolve();
    expect(llmClient.chat).toHaveBeenCalledTimes(3);
    expect(deferreds).toHaveLength(3);

    deferreds[1]!.resolve({
      message: { role: "assistant", content: "r1" },
      usage: { promptTokens: 1, completionTokens: 1 },
    });
    await Promise.resolve();
    expect(llmClient.chat).toHaveBeenCalledTimes(4);
    expect(deferreds).toHaveLength(4);

    deferreds[2]!.resolve({
      message: { role: "assistant", content: "r2" },
      usage: { promptTokens: 1, completionTokens: 1 },
    });
    await Promise.resolve();
    expect(llmClient.chat).toHaveBeenCalledTimes(totalRequests);
    expect(deferreds).toHaveLength(totalRequests);

    deferreds[3]!.resolve({
      message: { role: "assistant", content: "r3" },
      usage: { promptTokens: 1, completionTokens: 1 },
    });
    deferreds[4]!.resolve({
      message: { role: "assistant", content: "r4" },
      usage: { promptTokens: 1, completionTokens: 1 },
    });

    await engine.waitForIdle();
    expect(llmClient.chat).toHaveBeenCalledTimes(totalRequests);
  });

  it("coerces numeric-looking model outputs into numbers", async () => {
    const llmClient = {
      chat: vi.fn(async () => ({
        message: { role: "assistant", content: "42" },
        usage: { promptTokens: 1, completionTokens: 1 },
      })),
    };

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
    });

    const pending = evaluateFormula('=AI("give me a number", "x")', () => null, { ai: engine, cellAddress: "Sheet1!A1" });
    expect(pending).toBe(AI_CELL_PLACEHOLDER);
    await engine.waitForIdle();

    const resolved = evaluateFormula('=AI("give me a number", "x")', () => null, { ai: engine, cellAddress: "Sheet1!A1" });
    expect(resolved).toBe(42);
    expect(typeof resolved).toBe("number");
  });

  it("coerces TRUE/FALSE model outputs into booleans", async () => {
    const llmClient = {
      chat: vi.fn(async () => ({
        message: { role: "assistant", content: "TRUE" },
        usage: { promptTokens: 1, completionTokens: 1 },
      })),
    };

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
    });

    const pending = evaluateFormula('=AI("return true", "x")', () => null, { ai: engine, cellAddress: "Sheet1!A1" });
    expect(pending).toBe(AI_CELL_PLACEHOLDER);
    await engine.waitForIdle();

    const resolved = evaluateFormula('=AI("return true", "x")', () => null, { ai: engine, cellAddress: "Sheet1!A1" });
    expect(resolved).toBe(true);
    expect(typeof resolved).toBe("boolean");
  });

  it("truncates large string outputs before caching to keep the grid responsive", async () => {
    const maxOutputChars = 10_000;
    const long = "X".repeat(20_000);
    const llmClient = {
      chat: vi.fn(async () => ({
        message: { role: "assistant", content: long },
        usage: { promptTokens: 1, completionTokens: 1 },
      })),
    };

    const auditStore = new MemoryAIAuditStore();
    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore,
      sessionId: "truncate-session",
      limits: { maxOutputChars: maxOutputChars, maxAuditPreviewChars: 200 },
    });

    const pending = evaluateFormula('=AI("summarize", "hello")', () => null, { ai: engine, cellAddress: "Sheet1!A1" });
    expect(pending).toBe(AI_CELL_PLACEHOLDER);
    await engine.waitForIdle();

    const resolved = evaluateFormula('=AI("summarize", "hello")', () => null, { ai: engine, cellAddress: "Sheet1!A1" });
    expect(typeof resolved).toBe("string");
    expect((resolved as string).length).toBe(maxOutputChars);
    expect(resolved).toBe(`${long.slice(0, maxOutputChars - 1)}…`);

    // Audit entries should remain bounded: store previews/hashes, not the full output.
    const entries = await auditStore.listEntries({ session_id: "truncate-session" });
    expect(entries).toHaveLength(1);
    const input = entries[0]?.input as any;
    expect(input?.output_preview?.length).toBe(200);
    expect(input?.output_preview).not.toBe(resolved);
    expect(input?.output_hash).toMatch(/^[0-9a-f]{16}$/);
    expect(input?.output_value).toBeUndefined();
  });

  it("formats range inputs in LLM prompts using sheet display names (Excel quoting)", async () => {
    const sheetId = "sheet_test_1";
    const sheetIdToName = new Map([[sheetId, "O'Brien"]]);
    const sheetNameResolver = createSheetNameResolverFromIdToNameMap(sheetIdToName);

    const llmClient = {
      chat: vi.fn(async (_request: any) => ({
        message: { role: "assistant", content: "ok" },
        usage: { promptTokens: 1, completionTokens: 1 },
      })),
    };

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      sheetNameResolver,
    });

    const getCellValue = (addr: string) => (addr === "A1" ? "Name" : addr === "A2" ? "Alice" : null);

    const pending = evaluateFormula('=AI("summarize", A1:A2)', getCellValue, { ai: engine, cellAddress: `${sheetId}!B1` });
    expect(pending).toBe(AI_CELL_PLACEHOLDER);
    expect(llmClient.chat).toHaveBeenCalledTimes(1);

    const request = llmClient.chat.mock.calls[0]![0];
    const content = String(request?.messages?.[1]?.content ?? "");
    expect(content).toContain("'O''Brien'!A1:A2");
    expect(content).not.toContain(`${sheetId}!A1:A2`);
  });

  it("purges legacy persisted cache keys that embed raw prompt text", async () => {
    const persistKey = "ai-cache-legacy";
    const legacyPrompt = "summarize";
    const legacyPromptHex = "0123456789abcdef";
    globalThis.localStorage?.setItem(
      persistKey,
      JSON.stringify([
        { key: `test-model\u0000AI\u0000${legacyPrompt}\u0000deadbeef`, value: "legacy", updatedAtMs: 0 },
        // Guard against regressions: a legacy raw prompt that "looks like a hash" should still be rejected.
        { key: `test-model\u0000AI\u0000${legacyPromptHex}\u0000deadbeef`, value: "legacy_hex", updatedAtMs: 0 },
      ]),
    );

    const llmClient = {
      chat: vi.fn(async (_request: any) => ({
        message: { role: "assistant", content: "ok" },
        usage: { promptTokens: 1, completionTokens: 1 },
      })),
    };

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      model: "test-model",
      cache: { persistKey },
    });

    const pending = evaluateFormula('=AI("summarize", "hello")', () => null, { ai: engine, cellAddress: "Sheet1!A1" });
    expect(pending).toBe(AI_CELL_PLACEHOLDER);
    await engine.waitForIdle();

    const stored = globalThis.localStorage?.getItem(persistKey) ?? "";
    expect(stored).not.toContain("legacy");
    expect(stored).not.toContain(legacyPrompt);
    expect(stored).not.toContain(legacyPromptHex);
  });

  it("loads persisted cache keys with legacy 8-hex hashes", async () => {
    const persistKey = "ai-cache-compat";
    const legacyKey = `test-model\u0000AI\u00001234abcd\u0000deadbeef`;
    globalThis.localStorage?.setItem(persistKey, JSON.stringify([{ key: legacyKey, value: "legacy", updatedAtMs: 0 }]));

    const llmClient = {
      chat: vi.fn(async (_request: any) => ({
        message: { role: "assistant", content: "ok" },
        usage: { promptTokens: 1, completionTokens: 1 },
      })),
    };

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      model: "test-model",
      cache: { persistKey, maxEntries: 10 },
    });

    const pending = evaluateFormula('=AI("summarize", "hello")', () => null, { ai: engine, cellAddress: "Sheet1!A1" });
    expect(pending).toBe(AI_CELL_PLACEHOLDER);
    await engine.waitForIdle();

    const stored = globalThis.localStorage?.getItem(persistKey) ?? "";
    const parsed = JSON.parse(stored) as Array<any>;
    expect(parsed.some((entry) => entry?.key === legacyKey && entry?.value === "legacy")).toBe(true);
  });

  it("persists AI cell cache keys using hashes (no raw prompt/input strings)", async () => {
    const persistKey = "ai-cache-hashed";
    const llmClient = {
      chat: vi.fn(async (_request: any) => ({
        message: { role: "assistant", content: "ok" },
        usage: { promptTokens: 1, completionTokens: 1 },
      })),
    };

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      model: "test-model",
      cache: { persistKey },
    });

    const pending = evaluateFormula('=AI("summarize", "hello")', () => null, { ai: engine, cellAddress: "Sheet1!A1" });
    expect(pending).toBe(AI_CELL_PLACEHOLDER);
    await engine.waitForIdle();

    const stored = globalThis.localStorage?.getItem(persistKey) ?? "";
    expect(stored).not.toContain("summarize");
    expect(stored).not.toContain("hello");

    // Keys should be `${model}\0${fn}\0${promptHash}\0${inputsHash}`.
    expect(stored).toMatch(/test-model\\u0000AI\\u0000[0-9a-f]{16}\\u0000[0-9a-f]{16}/);
  });

  it("batches localStorage persistence across multiple cache writes", async () => {
    vi.useFakeTimers();
    let setItemSpy: ReturnType<typeof vi.spyOn> | null = null;
    try {
      const persistKey = "ai-cache-batched";
      setItemSpy = vi.spyOn(globalThis.localStorage as any, "setItem");

      let callIndex = 0;
      const llmClient = {
        chat: vi.fn(async () => {
          callIndex += 1;
          return {
            message: { role: "assistant", content: `ok_${callIndex}` },
            usage: { promptTokens: 1, completionTokens: 1 },
          };
        }),
      };

      const engine = new AiCellFunctionEngine({
        llmClient: llmClient as any,
        auditStore: new MemoryAIAuditStore(),
        model: "test-model",
        cache: { persistKey },
      });

      const formulas = ['=AI("task1", "hello")', '=AI("task2", "hello")', '=AI("task3", "hello")'];
      for (const formula of formulas) {
        const pending = evaluateFormula(formula, () => null, { ai: engine, cellAddress: "Sheet1!A1" });
        expect(pending).toBe(AI_CELL_PLACEHOLDER);
      }

      await engine.waitForIdle();

      const callsForKey = setItemSpy.mock.calls.filter((call) => call[0] === persistKey);
      expect(callsForKey).toHaveLength(1);
      expect(llmClient.chat).toHaveBeenCalledTimes(formulas.length);
    } finally {
      setItemSpy?.mockRestore();
      vi.useRealTimers();
    }
  });

  it("changing referenced cells invalidates the cache key", async () => {
    const deferred1 = defer<any>();
    const deferred2 = defer<any>();

    let callCount = 0;
    const llmClient = {
      chat: vi.fn((_request: any) => {
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

  it("DLP blocks disallowed inputs, avoids the LLM call, and records a blocked audit entry", async () => {
    const workbookId = "dlp-block-workbook";
    const llmClient = { chat: vi.fn() };

    markCellRestricted({ workbookId, sheetId: "Sheet1", row: 0, col: 0 }); // A1
    setBlockPolicy(workbookId);

    const auditStore = new MemoryAIAuditStore();
    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore,
      workbookId,
      sessionId: "dlp-block-session",
    });

    const getCellValue = (addr: string) => (addr === "A1" ? "top secret" : null);
    const value = evaluateFormula('=AI("summarize", A1)', getCellValue, { ai: engine, cellAddress: "Sheet1!B1" });
    expect(value).toBe(AI_CELL_DLP_ERROR);
    expect(llmClient.chat).not.toHaveBeenCalled();

    await engine.waitForIdle();
    const entries = await auditStore.listEntries({ session_id: "dlp-block-session" });
    expect(entries).toHaveLength(1);
    expect((entries[0]?.input as any)?.blocked).toBe(true);

    const input = entries[0]?.input as any;
    const events = getAiDlpAuditLogger().list();
    const dlpEvent = events.find((e: any) => e.details?.type === "ai.cell_function" && e.details?.documentId === workbookId);
    expect(dlpEvent?.details?.inputs_hash).toBe(input?.inputs_hash);
    expect(dlpEvent?.details?.prompt_hash).toBe(input?.prompt_hash);
  });

  it("enforces DLP BLOCK decisions for all AI cell function variants", () => {
    const variants = [
      { fn: "AI", formula: '=AI("summarize", A1)' },
      { fn: "AI.EXTRACT", formula: '=AI.EXTRACT("email", A1)' },
      { fn: "AI.CLASSIFY", formula: '=AI.CLASSIFY("A/B", A1)' },
      { fn: "AI.TRANSLATE", formula: '=AI.TRANSLATE("French", A1)' },
    ] as const;

    for (const variant of variants) {
      const workbookId = `dlp-block-${variant.fn}`;
      const llmClient = { chat: vi.fn() };

      markCellRestricted({ workbookId, sheetId: "Sheet1", row: 0, col: 0 }); // A1
      setBlockPolicy(workbookId);

      const engine = new AiCellFunctionEngine({
        llmClient: llmClient as any,
        auditStore: new MemoryAIAuditStore(),
        workbookId,
      });

      const value = evaluateFormula(variant.formula, (ref) => (ref === "A1" ? "top secret" : null), {
        ai: engine,
        cellAddress: "Sheet1!B1",
      });
      expect(value).toBe(AI_CELL_DLP_ERROR);
      expect(llmClient.chat).not.toHaveBeenCalled();
    }
  });

  it("does not persist restricted prompt text in audit logs for blocked runs", async () => {
    const llmClient = { chat: vi.fn() };
    const workbookId = "dlp-blocked-prompt-workbook";
    const sessionId = "dlp-blocked-prompt-session";

    markCellRestricted({ workbookId, sheetId: "Sheet1", row: 0, col: 0 }); // A1
    setBlockPolicy(workbookId);

    const auditStore = new MemoryAIAuditStore();
    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore,
      workbookId,
      sessionId,
    });

    const value = evaluateFormula('=AI(A1, "hello")', (ref) => (ref === "A1" ? "top secret" : null), {
      ai: engine,
      cellAddress: "Sheet1!B1",
    });
    expect(value).toBe(AI_CELL_DLP_ERROR);
    expect(llmClient.chat).not.toHaveBeenCalled();

    await engine.waitForIdle();
    const entries = await auditStore.listEntries({ session_id: sessionId });
    expect(entries).toHaveLength(1);
    const input = entries[0]?.input as any;
    expect(input?.prompt).toBe("[REDACTED]");
    expect(input?.prompt).not.toContain("top secret");
    expect(input?.prompt_hash).toMatch(/^[0-9a-f]{16}$/);
    expect(input?.inputs_hash).toMatch(/^[0-9a-f]{16}$/);
    expect(input?.inputs_compaction).toBeDefined();
    expect(input?.blocked).toBe(true);
  });

  it("DLP does not allow restricted cells to be smuggled via nested formulas (e.g. IF)", () => {
    const workbookId = "dlp-smuggle-if";
    const llmClient = { chat: vi.fn() };

    markCellRestricted({ workbookId, sheetId: "Sheet1", row: 0, col: 0 }); // A1
    setBlockPolicy(workbookId);

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      workbookId,
    });

    const value = evaluateFormula('=AI("summarize", IF(TRUE, A1, "x"))', (ref) => (ref === "A1" ? "top secret" : null), {
      ai: engine,
      cellAddress: "Sheet1!B1",
    });
    expect(value).toBe(AI_CELL_DLP_ERROR);
    expect(llmClient.chat).not.toHaveBeenCalled();
  });

  it("DLP does not allow restricted cells to be smuggled via conditional outputs (e.g. IF(A1,\"Y\",\"N\"))", () => {
    const workbookId = "dlp-smuggle-conditional";
    const llmClient = { chat: vi.fn() };

    markCellRestricted({ workbookId, sheetId: "Sheet1", row: 0, col: 0 }); // A1
    setBlockPolicy(workbookId);

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      workbookId,
    });

    const value = evaluateFormula('=AI("summarize", IF(A1, "Y", "N"))', (ref) => (ref === "A1" ? 1 : null), {
      ai: engine,
      cellAddress: "Sheet1!B1",
    });
    expect(value).toBe(AI_CELL_DLP_ERROR);
    expect(llmClient.chat).not.toHaveBeenCalled();
  });

  it("DLP does not allow restricted cells to be smuggled via arithmetic coercion (e.g. A1+0)", () => {
    const workbookId = "dlp-smuggle-arithmetic";
    const llmClient = { chat: vi.fn() };

    markCellRestricted({ workbookId, sheetId: "Sheet1", row: 0, col: 0 }); // A1
    setBlockPolicy(workbookId);

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      workbookId,
    });

    const value = evaluateFormula('=AI("summarize", A1+0)', (ref) => (ref === "A1" ? 123 : null), {
      ai: engine,
      cellAddress: "Sheet1!B1",
    });
    expect(value).toBe(AI_CELL_DLP_ERROR);
    expect(llmClient.chat).not.toHaveBeenCalled();
  });

  it("DLP does not allow restricted cells to be smuggled via derived aggregations (e.g. SUM(range))", () => {
    const workbookId = "dlp-smuggle-sum";
    const llmClient = { chat: vi.fn() };

    markCellRestricted({ workbookId, sheetId: "Sheet1", row: 0, col: 0 }); // A1
    setBlockPolicy(workbookId);

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      workbookId,
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

  it("resolves DLP classifications for quoted sheet names containing semicolons", () => {
    const workbookId = "dlp-sheet-semicolons";
    const llmClient = { chat: vi.fn() };

    markCellRestricted({ workbookId, sheetId: "A;B", row: 0, col: 0 }); // A1 on A;B
    setBlockPolicy(workbookId);

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      workbookId,
    });

    const value = evaluateFormula('=AI("summarize", \'A;B\'!A1)', (ref) => (ref === "A;B!A1" ? "top secret" : null), {
      ai: engine,
      cellAddress: "Sheet1!B1",
    });
    expect(value).toBe(AI_CELL_DLP_ERROR);
    expect(llmClient.chat).not.toHaveBeenCalled();
  });

  it("heuristically redacts sensitive referenced values even without structured classifications", async () => {
    // Ensure the classification store is empty so enforcement relies on heuristics.
    globalThis.localStorage?.clear();

    const workbookId = "dlp-heuristic-redact-workbook";
    const llmClient = {
      chat: vi.fn(async (_request: any) => ({
        message: { role: "assistant", content: "ok" },
        usage: { promptTokens: 1, completionTokens: 1 },
      })),
    };

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      workbookId,
    });

    const getCellValue = (addr: string) => (addr === "A1" ? "user@example.com" : null);

    const pending = evaluateFormula('=AI("summarize", A1)', getCellValue, { ai: engine, cellAddress: "Sheet1!B1" });
    expect(pending).toBe(AI_CELL_PLACEHOLDER);
    await engine.waitForIdle();

    expect(llmClient.chat).toHaveBeenCalledTimes(1);
    const call = llmClient.chat.mock.calls[0]?.[0];
    const userMessage = call?.messages?.find((m: any) => m.role === "user")?.content ?? "";
    expect(userMessage).toContain("[REDACTED]");
    expect(userMessage).not.toContain("user@example.com");

    const resolved = evaluateFormula('=AI("summarize", A1)', getCellValue, { ai: engine, cellAddress: "Sheet1!B1" });
    expect(resolved).toBe("ok");
  });

  it("heuristically redacts referenced private key blocks even when cell values are truncated", async () => {
    // Ensure the classification store is empty so enforcement relies on heuristics.
    globalThis.localStorage?.clear();

    const workbookId = "dlp-heuristic-private-key";
    const llmClient = {
      chat: vi.fn(async (_request: any) => ({
        message: { role: "assistant", content: "ok" },
        usage: { promptTokens: 1, completionTokens: 1 },
      })),
    };

    // Force aggressive truncation so the prompt-formatted scalar would *not* include the full
    // begin/end markers. Heuristic scanning should still detect the full value and redact it.
    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      workbookId,
      limits: { maxCellChars: 50 },
    });

    const privateKey = `-----BEGIN PRIVATE KEY-----\n${"A".repeat(200)}\n-----END PRIVATE KEY-----`;
    const getCellValue = (addr: string) => (addr === "A1" ? privateKey : null);

    const pending = evaluateFormula('=AI("summarize", A1)', getCellValue, { ai: engine, cellAddress: "Sheet1!B1" });
    expect(pending).toBe(AI_CELL_PLACEHOLDER);
    await engine.waitForIdle();

    expect(llmClient.chat).toHaveBeenCalledTimes(1);
    const call = llmClient.chat.mock.calls[0]?.[0];
    const userMessage = call?.messages?.find((m: any) => m.role === "user")?.content ?? "";
    expect(userMessage).toContain("[REDACTED]");
    expect(userMessage).not.toContain("BEGIN PRIVATE KEY");
    expect(userMessage).not.toContain("END PRIVATE KEY");

    const resolved = evaluateFormula('=AI("summarize", A1)', getCellValue, { ai: engine, cellAddress: "Sheet1!B1" });
    expect(resolved).toBe("ok");
  });

  it("heuristically redacts sensitive values inside referenced ranges even without structured classifications", async () => {
    // Ensure the classification store is empty so enforcement relies on heuristics.
    globalThis.localStorage?.clear();

    const workbookId = "dlp-heuristic-range-redact";
    const llmClient = {
      chat: vi.fn(async (_request: any) => ({
        message: { role: "assistant", content: "ok" },
        usage: { promptTokens: 1, completionTokens: 1 },
      })),
    };

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      workbookId,
    });

    const getCellValue = (addr: string) => (addr === "A1" ? "user@example.com" : addr === "A2" ? "public payload" : null);

    const pending = evaluateFormula('=AI("summarize", A1:A2)', getCellValue, { ai: engine, cellAddress: "Sheet1!B1" });
    expect(pending).toBe(AI_CELL_PLACEHOLDER);
    await engine.waitForIdle();

    expect(llmClient.chat).toHaveBeenCalledTimes(1);
    const call = llmClient.chat.mock.calls[0]?.[0];
    const userMessage = call?.messages?.find((m: any) => m.role === "user")?.content ?? "";
    expect(userMessage).toContain("[REDACTED]");
    expect(userMessage).toContain("public payload");
    expect(userMessage).not.toContain("user@example.com");

    const resolved = evaluateFormula('=AI("summarize", A1:A2)', getCellValue, { ai: engine, cellAddress: "Sheet1!B1" });
    expect(resolved).toBe("ok");
  });

  it("blocks heuristic Restricted values under block policy even without structured classifications", async () => {
    // Ensure the classification store is empty so enforcement relies on heuristics.
    globalThis.localStorage?.clear();

    const workbookId = "dlp-heuristic-block-workbook";
    const sessionId = "dlp-heuristic-block-session";
    const llmClient = { chat: vi.fn() };

    setBlockPolicy(workbookId);

    const auditStore = new MemoryAIAuditStore();
    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore,
      workbookId,
      sessionId,
    });

    const getCellValue = (addr: string) => (addr === "A1" ? "user@example.com" : null);

    const value = evaluateFormula('=AI("summarize", A1)', getCellValue, { ai: engine, cellAddress: "Sheet1!B1" });
    expect(value).toBe(AI_CELL_DLP_ERROR);
    expect(llmClient.chat).not.toHaveBeenCalled();

    await engine.waitForIdle();
    const entries = await auditStore.listEntries({ session_id: sessionId });
    expect(entries).toHaveLength(1);
    const input = entries[0]?.input as any;
    expect(input?.blocked).toBe(true);
    expect(JSON.stringify(input)).not.toContain("user@example.com");
  });

  it("heuristically redacts referenced API keys even without structured classifications", async () => {
    globalThis.localStorage?.clear();

    const workbookId = "dlp-heuristic-api-key";
    const llmClient = {
      chat: vi.fn(async (_request: any) => ({
        message: { role: "assistant", content: "ok" },
        usage: { promptTokens: 1, completionTokens: 1 },
      })),
    };

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      workbookId,
    });

    // Stripe-like secret key (matches ai-context API_KEY_RE).
    const apiKey = "sk_live_1234567890abcdef12345678";
    const getCellValue = (addr: string) => (addr === "A1" ? apiKey : null);

    const pending = evaluateFormula('=AI("summarize", A1)', getCellValue, { ai: engine, cellAddress: "Sheet1!B1" });
    expect(pending).toBe(AI_CELL_PLACEHOLDER);
    await engine.waitForIdle();

    expect(llmClient.chat).toHaveBeenCalledTimes(1);
    const call = llmClient.chat.mock.calls[0]?.[0];
    const userMessage = call?.messages?.find((m: any) => m.role === "user")?.content ?? "";
    expect(userMessage).toContain("[REDACTED]");
    expect(userMessage).not.toContain(apiKey);

    const resolved = evaluateFormula('=AI("summarize", A1)', getCellValue, { ai: engine, cellAddress: "Sheet1!B1" });
    expect(resolved).toBe("ok");
  });

  it("DLP redacts inputs before sending to the LLM", async () => {
    const workbookId = "dlp-redact-workbook";
    const llmClient = {
      chat: vi.fn(async (_request: any) => ({
        message: { role: "assistant", content: "ok" },
        usage: { promptTokens: 1, completionTokens: 1 },
      })),
    };

    markCellRestricted({ workbookId, sheetId: "Sheet1", row: 0, col: 0 }); // A1

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      workbookId,
    });

    const getCellValue = (addr: string) => (addr === "A1" ? "secret payload" : null);
    const pending = evaluateFormula('=AI("summarize", A1)', getCellValue, { ai: engine, cellAddress: "Sheet1!B1" });
    expect(pending).toBe(AI_CELL_PLACEHOLDER);
    await engine.waitForIdle();

    const call = llmClient.chat.mock.calls[0]?.[0];
    const userMessage = call?.messages?.find((m: any) => m.role === "user")?.content ?? "";
    expect(userMessage).toContain("[REDACTED]");
    expect(userMessage).not.toContain("secret payload");
  });

  it("DLP heuristically redacts sensitive patterns in referenced cells without classification records", async () => {
    const workbookId = "dlp-heuristic-redact-cell";
    const llmClient = {
      chat: vi.fn(async (_request: any) => ({
        message: { role: "assistant", content: "ok" },
        usage: { promptTokens: 1, completionTokens: 1 },
      })),
    };

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      workbookId,
    });

    const getCellValue = (addr: string) => (addr === "A1" ? "user@example.com" : null);
    const pending = evaluateFormula('=AI("summarize", A1)', getCellValue, { ai: engine, cellAddress: "Sheet1!B1" });
    expect(pending).toBe(AI_CELL_PLACEHOLDER);
    await engine.waitForIdle();

    const call = llmClient.chat.mock.calls[0]?.[0];
    const userMessage = call?.messages?.find((m: any) => m.role === "user")?.content ?? "";
    expect(userMessage).toContain("[REDACTED]");
    expect(userMessage).not.toContain("user@example.com");
  });

  it("DLP heuristically redacts sensitive patterns within referenced ranges without classification records", async () => {
    const workbookId = "dlp-heuristic-redact-range";
    const llmClient = {
      chat: vi.fn(async (_request: any) => ({
        message: { role: "assistant", content: "ok" },
        usage: { promptTokens: 1, completionTokens: 1 },
      })),
    };

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      workbookId,
    });

    const getCellValue = (addr: string) => (addr === "A1" ? "user@example.com" : addr === "A2" ? "public payload" : null);
    const pending = evaluateFormula('=AI("summarize", A1:A2)', getCellValue, { ai: engine, cellAddress: "Sheet1!B1" });
    expect(pending).toBe(AI_CELL_PLACEHOLDER);
    await engine.waitForIdle();

    const call = llmClient.chat.mock.calls[0]?.[0];
    const userMessage = call?.messages?.find((m: any) => m.role === "user")?.content ?? "";
    expect(userMessage).toContain("[REDACTED]");
    expect(userMessage).toContain("public payload");
    expect(userMessage).not.toContain("user@example.com");
  });

  it("DLP heuristically detects long private key blocks in referenced cells", async () => {
    const workbookId = "dlp-heuristic-private-key";
    const llmClient = {
      chat: vi.fn(async (_request: any) => ({
        message: { role: "assistant", content: "ok" },
        usage: { promptTokens: 1, completionTokens: 1 },
      })),
    };

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      workbookId,
    });

    // Long enough that a prefix-only truncation would omit the END marker, causing a miss.
    const privateKey = `-----BEGIN PRIVATE KEY-----\n${"A".repeat(5000)}\n-----END PRIVATE KEY-----`;
    const getCellValue = (addr: string) => (addr === "A1" ? privateKey : null);
    const pending = evaluateFormula('=AI("summarize", A1)', getCellValue, { ai: engine, cellAddress: "Sheet1!B1" });
    expect(pending).toBe(AI_CELL_PLACEHOLDER);
    await engine.waitForIdle();

    const call = llmClient.chat.mock.calls[0]?.[0];
    const userMessage = call?.messages?.find((m: any) => m.role === "user")?.content ?? "";
    expect(userMessage).toContain("[REDACTED]");
    expect(userMessage).not.toContain("BEGIN PRIVATE KEY");
    expect(userMessage).not.toContain("END PRIVATE KEY");
  });

  it("DLP redacts only disallowed cells within a mixed-classification range", async () => {
    const workbookId = "dlp-redact-range-workbook";
    const llmClient = {
      chat: vi.fn(async (_request: any) => ({
        message: { role: "assistant", content: "ok" },
        usage: { promptTokens: 1, completionTokens: 1 },
      })),
    };

    markCellRestricted({ workbookId, sheetId: "Sheet1", row: 0, col: 0 }); // A1

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      workbookId,
    });

    const pending = evaluateFormula(
      '=AI("summarize", A1:A2)',
      (ref) => (ref === "A1" ? "secret payload" : ref === "A2" ? "public payload" : null),
      {
        ai: engine,
        cellAddress: "Sheet1!B1",
      },
    );
    expect(pending).toBe(AI_CELL_PLACEHOLDER);
    await engine.waitForIdle();

    const call = llmClient.chat.mock.calls[0]?.[0];
    const userMessage = call?.messages?.find((m: any) => m.role === "user")?.content ?? "";
    expect(userMessage).toContain("[REDACTED]");
    expect(userMessage).toContain("public payload");
    expect(userMessage).not.toContain("secret payload");
  });

  it("DLP redacts only disallowed cells in provenance arrays without a range ref", async () => {
    const workbookId = "dlp-provenance-array";
    const llmClient = {
      chat: vi.fn(async (_request: any) => ({
        message: { role: "assistant", content: "ok" },
        usage: { promptTokens: 1, completionTokens: 1 },
      })),
    };

    markCellRestricted({ workbookId, sheetId: "Sheet1", row: 0, col: 0 }); // A1

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      workbookId,
    });

    const input = [
      { __cellRef: "Sheet1!A1", value: "secret payload" },
      { __cellRef: "Sheet1!A2", value: "public payload" },
    ];

    const pending = engine.evaluateAiFunction({
      name: "AI",
      args: ["summarize", input as any],
      cellAddress: "Sheet1!B1",
    });
    expect(pending).toBe(AI_CELL_PLACEHOLDER);
    await engine.waitForIdle();

    const call = llmClient.chat.mock.calls[0]?.[0];
    const userMessage = call?.messages?.find((m: any) => m.role === "user")?.content ?? "";
    expect(userMessage).toContain("[REDACTED]");
    expect(userMessage).toContain("public payload");
    expect(userMessage).not.toContain("secret payload");
  });

  it("does not iterate over full array inputs when compacting large AI arguments", async () => {
    const workbookId = "budget-large-array";
    const llmClient = {
      chat: vi.fn(async () => ({
        message: { role: "assistant", content: "ok" },
        usage: { promptTokens: 1, completionTokens: 1 },
      })),
    };

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      workbookId,
    });

    const length = 100_000;
    let indexReads = 0;
    const backing: any[] = [];
    const large = new Proxy(backing, {
      get(_target, prop) {
        if (prop === "length") return length;
        if (typeof prop === "string" && /^\d+$/.test(prop)) {
          indexReads += 1;
          const idx = Number(prop);
          return `V${idx}`;
        }
        return undefined;
      },
    });

    const pending = engine.evaluateAiFunction({
      name: "AI",
      args: ["summarize", large as any],
      cellAddress: "Sheet1!B1",
    });
    expect(pending).toBe(AI_CELL_PLACEHOLDER);
    await engine.waitForIdle();

    // We should only touch the handful of indices used for preview + sampling, not 100k.
    expect(indexReads).toBeLessThan(500);
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

  it("does not treat arbitrary #prefixed strings as spreadsheet errors", async () => {
    const workbookId = "hash-text-workbook";
    const llmClient = {
      chat: vi.fn(async (_request: any) => ({
        message: { role: "assistant", content: "ok" },
        usage: { promptTokens: 1, completionTokens: 1 },
      })),
    };
    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      workbookId,
    });

    const pending = evaluateFormula('=AI("summarize", A1)', (ref) => (ref === "A1" ? "#hashtag" : null), {
      ai: engine,
      cellAddress: "Sheet1!B1",
    });
    expect(pending).toBe(AI_CELL_PLACEHOLDER);
    await engine.waitForIdle();

    expect(llmClient.chat).toHaveBeenCalledTimes(1);
    const call = llmClient.chat.mock.calls[0]?.[0];
    const userMessage = call?.messages?.find((m: any) => m.role === "user")?.content ?? "";
    expect(userMessage).toContain("#hashtag");
  });

  it("budgeting compacts large ranges in the prompt", async () => {
    const workbookId = "budget-workbook";
    const llmClient = {
      chat: vi.fn(async (_request: any) => ({
        message: { role: "assistant", content: "ok" },
        usage: { promptTokens: 1, completionTokens: 1 },
      })),
    };

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      workbookId,
    });

    const getCellValue = (addr: string) => {
      const match = /^A(\d+)$/.exec(addr);
      if (!match) return null;
      const n = Number(match[1]);
      return `CELL_${String(n).padStart(4, "0")}`;
    };

    const pending = evaluateFormula('=AI("summarize", A1:A1000)', getCellValue, { ai: engine, cellAddress: "Sheet1!B1" });
    expect(pending).toBe(AI_CELL_PLACEHOLDER);
    await engine.waitForIdle();

    const call = llmClient.chat.mock.calls[0]?.[0];
    const userMessage = call?.messages?.find((m: any) => m.role === "user")?.content ?? "";
    expect(userMessage.length).toBeLessThan(8_000);
    expect(userMessage).toContain('"total_cells":1000');
    expect(userMessage).toContain('"sampled_cells":200');
    expect(userMessage).toContain('"truncated":true');

    const occurrences = userMessage.match(/CELL_\d{4}/g)?.length ?? 0;
    expect(occurrences).toBeGreaterThan(0);
    expect(occurrences).toBeLessThan(200);

    // Should include sampled cells beyond the first N rows (not just a prefix).
    const matches = Array.from(userMessage.matchAll(/CELL_(\d{4})/g) as Iterable<RegExpMatchArray>);
    const maxCell = Math.max(...matches.map((match) => Number(match[1])));
    expect(maxCell).toBeGreaterThan(200);
  });

  it("truncates large cell values before serializing inputs into the prompt", async () => {
    const llmClient = {
      chat: vi.fn(async (_request: any) => ({
        message: { role: "assistant", content: "ok" },
        usage: { promptTokens: 1, completionTokens: 1 },
      })),
    };

    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore: new MemoryAIAuditStore(),
      limits: { maxCellChars: 50 },
    });

    const long = `PREFIX_${"X".repeat(200)}_SUFFIX`;
    const pending = evaluateFormula('=AI("summarize", A1)', (ref) => (ref === "A1" ? long : null), {
      ai: engine,
      cellAddress: "Sheet1!B1",
    });
    expect(pending).toBe(AI_CELL_PLACEHOLDER);
    await engine.waitForIdle();

    const call = llmClient.chat.mock.calls[0]?.[0];
    const userMessage = call?.messages?.find((m: any) => m.role === "user")?.content ?? "";
    expect(userMessage).toContain("PREFIX_");
    expect(userMessage).not.toContain("_SUFFIX");
    expect(userMessage).toContain("[TRUNCATED]");
    expect(userMessage).toContain("…");
  });

  it("truncates long prompts in audit entries", async () => {
    const llmClient = {
      chat: vi.fn(async () => ({
        message: { role: "assistant", content: "ok" },
        usage: { promptTokens: 1, completionTokens: 1 },
      })),
    };

    const auditStore = new MemoryAIAuditStore();
    const engine = new AiCellFunctionEngine({
      llmClient: llmClient as any,
      auditStore,
      sessionId: "audit-session",
      limits: { maxAuditPreviewChars: 200 },
    });

    const longPrompt = "P".repeat(500);
    const pending = evaluateFormula(`=AI("${longPrompt}", "hello")`, () => null, { ai: engine, cellAddress: "Sheet1!A1" });
    expect(pending).toBe(AI_CELL_PLACEHOLDER);
    await engine.waitForIdle();

    const entries = await auditStore.listEntries({ session_id: "audit-session" });
    expect(entries).toHaveLength(1);
    const input = entries[0]?.input as any;
    expect(typeof input?.prompt).toBe("string");
    expect(input.prompt).not.toBe(longPrompt);
    expect(input.prompt.length).toBe(200);
    expect(input.prompt).toContain("[TRUNCATED]");
    expect(input.prompt.endsWith("…")).toBe(true);
    expect(input.prompt_hash).toMatch(/^[0-9a-f]{16}$/);
  });
});
