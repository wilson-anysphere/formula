import { beforeEach, describe, expect, it, vi } from "vitest";

import { evaluateFormula } from "./evaluateFormula.js";
import { AI_CELL_ERROR, AI_CELL_PLACEHOLDER, AiCellFunctionEngine } from "./AiCellFunctionEngine.js";

import { MemoryAIAuditStore } from "../../../../packages/ai-audit/src/memory-store.js";

describe("AiCellFunctionEngine.dispose()", () => {
  beforeEach(() => {
    globalThis.localStorage?.clear();
  });

  it("aborts in-flight requests without persisting #AI! errors", async () => {
    vi.useFakeTimers();
    try {
      let observedSignal: AbortSignal | undefined;
      const llmClient = {
        chat: vi.fn((request: any) => {
          observedSignal = request?.signal;
          return new Promise((_resolve, reject) => {
            const signal = request?.signal as AbortSignal | undefined;
            if (!signal) {
              reject(new Error("missing abort signal"));
              return;
            }
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

      const persistKey = "test:ai-cell-cache";
      const engine = new AiCellFunctionEngine({
        llmClient: llmClient as any,
        model: "test-model",
        auditStore: new MemoryAIAuditStore(),
        cache: { persistKey },
      });

      const pending = evaluateFormula('=AI("summarize", "hello")', () => null, { ai: engine, cellAddress: "Sheet1!A1" });
      expect(pending).toBe(AI_CELL_PLACEHOLDER);
      expect(llmClient.chat).toHaveBeenCalledTimes(1);

      engine.dispose();
      expect(observedSignal?.aborted).toBe(true);

      // Allow any abort-driven catch/finally handlers to run.
      await Promise.resolve();
      vi.advanceTimersByTime(100);
      await Promise.resolve();

      const persisted = globalThis.localStorage?.getItem(persistKey);
      expect(persisted ?? "").not.toContain(AI_CELL_ERROR);
    } finally {
      vi.useRealTimers();
    }
  });
});

