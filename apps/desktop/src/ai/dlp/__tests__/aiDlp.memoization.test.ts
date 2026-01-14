import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { createDefaultOrgPolicy } from "../../../../../../packages/security/dlp/src/policy.js";
import { LocalClassificationStore } from "../../../../../../packages/security/dlp/src/classificationStore.js";
import { getAiCloudDlpOptions } from "../aiDlp.js";

function createInMemoryLocalStorage(): Storage {
  const store = new Map<string, string>();
  return {
    getItem: (key: string) => (store.has(key) ? store.get(key)! : null),
    setItem: (key: string, value: string) => {
      store.set(String(key), String(value));
    },
    removeItem: (key: string) => {
      store.delete(String(key));
    },
    clear: () => {
      store.clear();
    },
    key: (index: number) => Array.from(store.keys())[index] ?? null,
    get length() {
      return store.size;
    }
  } as Storage;
}

const orgId = "memoization-test-org";
const documentId = "memoization-test-doc";

function makeRecordListInstrumenter() {
  let passes = 0;
  let elementGets = 0;

  const wrap = (records: any[]) =>
    new Proxy(records, {
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
  };
}

describe("getAiCloudDlpOptions memoization", () => {
  beforeEach(() => {
    const storage = createInMemoryLocalStorage();
    vi.stubGlobal("localStorage", storage);
    storage.clear();
  });

  afterEach(() => {
    try {
      (globalThis.localStorage as any)?.clear?.();
      // Ensure cached entries don't keep large arrays alive across the broader test suite.
      getAiCloudDlpOptions({ documentId, orgId });
    } finally {
      vi.unstubAllGlobals();
      vi.restoreAllMocks();
    }
  });

  it("reuses parsed policy + classification records until underlying localStorage strings change", () => {
    const storage = globalThis.localStorage;

    const classificationRecords = Array.from({ length: 2000 }, (_, i) => ({
      selector: { idx: i },
      classification: { level: "PUBLIC", labels: [] },
      updatedAt: "2024-01-01T00:00:00.000Z",
    }));

    storage.setItem(`dlp:orgPolicy:${orgId}`, JSON.stringify(createDefaultOrgPolicy()));
    storage.setItem(`dlp:classifications:${documentId}`, JSON.stringify(classificationRecords));

    const parseSpy = vi.spyOn(JSON, "parse");

    const instrumenter = makeRecordListInstrumenter();
    const originalList = LocalClassificationStore.prototype.list;
    const listSpy = vi.spyOn(LocalClassificationStore.prototype, "list").mockImplementation(function (docId: string) {
      const out = originalList.call(this, docId) as any[];
      return instrumenter.wrap(out);
    });

    try {
      const first = getAiCloudDlpOptions({ documentId, orgId });
      const callsAfterFirst = parseSpy.mock.calls.length;
      const readsAfterFirst = instrumenter.getElementGets();

      const second = getAiCloudDlpOptions({ documentId, orgId });
      expect(parseSpy).toHaveBeenCalledTimes(callsAfterFirst);
      // Perf proxy: the cached fast path should not re-scan all classification records to recompute cache keys.
      // If a regression calls computeDlpCacheKey again, we'd see ~2000 additional element reads here.
      expect(instrumenter.getElementGets() - readsAfterFirst).toBeLessThan(50);
      expect(second.policy).toBe(first.policy);
      expect(second.classificationStore).toBe(first.classificationStore);
      expect(second.classificationRecords).toBe(first.classificationRecords);

      storage.setItem(
        `dlp:classifications:${documentId}`,
        JSON.stringify([
          ...classificationRecords,
          {
            selector: { idx: 2001 },
            classification: { level: "PUBLIC", labels: [] },
            updatedAt: "2024-01-01T00:00:00.000Z",
          },
        ]),
      );

      const callsBeforeThird = parseSpy.mock.calls.length;
      const third = getAiCloudDlpOptions({ documentId, orgId });
      expect(parseSpy.mock.calls.length).toBeGreaterThan(callsBeforeThird);
      expect(third.classificationStore).toBe(first.classificationStore);
      expect(third.classificationRecords).not.toBe(first.classificationRecords);
    } finally {
      listSpy.mockRestore();
    }
  });
});
