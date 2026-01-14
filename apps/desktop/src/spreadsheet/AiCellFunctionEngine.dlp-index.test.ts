import { beforeEach, describe, expect, it, vi } from "vitest";

import { evaluateFormula } from "./evaluateFormula.js";
import { AI_CELL_PLACEHOLDER, AiCellFunctionEngine, __dlpIndexBuilder } from "./AiCellFunctionEngine.js";

import { LocalClassificationStore } from "../../../../packages/security/dlp/src/classificationStore.js";

function makeRecordListInstrumenter() {
  let passes = 0;
  let elementGets = 0;
  let propGets = 0;
  const objectProxyCache = new WeakMap<object, any>();

  const wrapObject = (value: any): any => {
    if (!value || typeof value !== "object") return value;
    if (Array.isArray(value)) return value;
    // Avoid proxying built-ins with internal slots (e.g. Map/Set) since their methods
    // can throw when `this` is a Proxy.
    if (value instanceof Map || value instanceof Set || value instanceof Date) return value;
    const cached = objectProxyCache.get(value);
    if (cached) return cached;
    const proxy = new Proxy(value, {
      get(target, prop, receiver) {
        propGets += 1;
        return wrapObject(Reflect.get(target, prop, receiver));
      },
    });
    objectProxyCache.set(value, proxy);
    return proxy;
  };

  const wrap = (records: any[]) =>
    new Proxy((records ?? []).map((r) => wrapObject(r)), {
      get(target, prop, receiver) {
        if (prop === Symbol.iterator) {
          return function () {
            passes += 1;
            // Bind iterator to proxy so numeric index access is observable.
            return Array.prototype[Symbol.iterator].call(receiver);
          };
        }
        if (typeof prop === "string" && /^[0-9]+$/.test(prop)) {
          elementGets += 1;
        }
        return Reflect.get(target, prop, receiver);
      },
    });

  return {
    wrap,
    getPasses: () => passes,
    getElementGets: () => elementGets,
    getPropGets: () => propGets,
  };
}

describe("AiCellFunctionEngine DLP indexing", () => {
  beforeEach(() => {
    globalThis.localStorage?.clear();
  });

  it("does not scan all classification records per referenced cell", async () => {
    const workbookId = "ai-cell-dlp-index";
    const sheetId = "Sheet1";

    // Many per-cell classification records (structured), plus many referenced cells in the formula.
    // This amplifies any O(records * referencedCells) regressions.
    const records = [];
    const updatedAt = new Date().toISOString();
    // Keep the record set reasonably sized so the test stays fast while still catching
    // per-referenced-cell scans (which would multiply this count by ~200).
    for (let row = 0; row < 500; row++) {
      records.push({
        selector: { scope: "cell", documentId: workbookId, sheetId, row, col: 0 },
        classification: { level: row === 0 ? "Restricted" : "Public", labels: [] },
        updatedAt,
      });
    }

    globalThis.localStorage?.setItem(`dlp:classifications:${workbookId}`, JSON.stringify(records));

    const instrumenter = makeRecordListInstrumenter();
    const originalList = LocalClassificationStore.prototype.list;
    const listSpy = vi.spyOn(LocalClassificationStore.prototype, "list").mockImplementation(function (documentId: string) {
      const out = originalList.call(this, documentId) as any[];
      return instrumenter.wrap(out);
    });

    try {
      const llmClient = {
        chat: vi.fn(async () => ({
          message: { role: "assistant", content: "ok" },
          usage: { promptTokens: 1, completionTokens: 1 },
        })),
      };

      const engine = new AiCellFunctionEngine({ llmClient: llmClient as any, workbookId, model: "test-model" });

      const refs = Array.from({ length: 200 }, (_v, idx) => `A${idx + 1}`);
      const formula = `=AI(\"summarize\", ${refs.join(", ")})`;

      const getCellValue = (addr: string) => (addr === "A1" ? "top secret" : addr.startsWith("A") ? "x" : null);
      const value = evaluateFormula(formula, getCellValue, { ai: engine, cellAddress: "Sheet1!B1" });

      expect(value).toBe(AI_CELL_PLACEHOLDER);

      // Perf proxy: building the index and hashing records can require a handful of linear scans.
      // A per-referenced-cell scan regression would multiply this by ~200.
      expect(instrumenter.getPasses()).toBeLessThan(50);
      expect(instrumenter.getElementGets()).toBeLessThan(20_000);
      // Defense-in-depth: catch per-referenced-cell scans even if the record list is cloned once and scanned repeatedly.
      expect(instrumenter.getPropGets()).toBeLessThan(200_000);

      await engine.waitForIdle();
      expect(llmClient.chat).toHaveBeenCalledTimes(1);
    } finally {
      listSpy.mockRestore();
    }
  });

  it("memoizes DLP cell indices across evaluations and invalidates on record changes", async () => {
    const workbookId = "ai-cell-dlp-index-memo";
    const sheetId = "Sheet1";
    const updatedAt = new Date().toISOString();

    const records = [
      {
        selector: { scope: "cell", documentId: workbookId, sheetId, row: 0, col: 0 },
        classification: { level: "Restricted", labels: [] },
        updatedAt,
      },
      {
        selector: { scope: "cell", documentId: workbookId, sheetId, row: 0, col: 1 },
        classification: { level: "Public", labels: [] },
        updatedAt,
      },
    ];

    globalThis.localStorage?.setItem(`dlp:classifications:${workbookId}`, JSON.stringify(records));

    const llmClient = {
      chat: vi.fn(async () => ({
        message: { role: "assistant", content: "ok" },
        usage: { promptTokens: 1, completionTokens: 1 },
      })),
    };

    const engine = new AiCellFunctionEngine({ llmClient: llmClient as any, workbookId, model: "test-model" });

    const buildSpy = vi.spyOn(__dlpIndexBuilder, "buildDlpCellIndex");

    const getCellValue = (addr: string) => (addr === "A1" ? "top secret" : null);

    const v1 = evaluateFormula('=AI("one", A1)', getCellValue, { ai: engine, cellAddress: "Sheet1!B1" });
    expect(v1).toBe(AI_CELL_PLACEHOLDER);
    expect(buildSpy).toHaveBeenCalledTimes(1);

    const v2 = evaluateFormula('=AI("two", A1)', getCellValue, { ai: engine, cellAddress: "Sheet1!C1" });
    expect(v2).toBe(AI_CELL_PLACEHOLDER);
    expect(buildSpy).toHaveBeenCalledTimes(1);

    await engine.waitForIdle();
    expect(llmClient.chat).toHaveBeenCalledTimes(2);

    const next = [
      ...records,
      {
        selector: { scope: "cell", documentId: workbookId, sheetId, row: 1, col: 0 },
        classification: { level: "Public", labels: [] },
        updatedAt,
      },
    ];
    globalThis.localStorage?.setItem(`dlp:classifications:${workbookId}`, JSON.stringify(next));

    const v3 = evaluateFormula('=AI("three", A1)', getCellValue, { ai: engine, cellAddress: "Sheet1!D1" });
    expect(v3).toBe(AI_CELL_PLACEHOLDER);
    expect(buildSpy).toHaveBeenCalledTimes(2);

    await engine.waitForIdle();
    expect(llmClient.chat).toHaveBeenCalledTimes(3);

    buildSpy.mockRestore();
  });
});
