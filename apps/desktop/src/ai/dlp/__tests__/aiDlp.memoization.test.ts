import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { createDefaultOrgPolicy } from "../../../../../../packages/security/dlp/src/policy.js";

vi.mock("../dlpCacheKey.js", async () => {
  const actual = await vi.importActual<any>("../dlpCacheKey.js");
  return {
    ...actual,
    computeDlpCacheKey: vi.fn(actual.computeDlpCacheKey),
  };
});

import { computeDlpCacheKey } from "../dlpCacheKey.js";
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

    const first = getAiCloudDlpOptions({ documentId, orgId });
    expect(computeDlpCacheKey).toHaveBeenCalledTimes(1);
    const callsAfterFirst = parseSpy.mock.calls.length;

    const second = getAiCloudDlpOptions({ documentId, orgId });
    expect(parseSpy).toHaveBeenCalledTimes(callsAfterFirst);
    expect(computeDlpCacheKey).toHaveBeenCalledTimes(1);
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
    expect(computeDlpCacheKey).toHaveBeenCalledTimes(2);
    expect(parseSpy.mock.calls.length).toBeGreaterThan(callsBeforeThird);
    expect(third.classificationStore).toBe(first.classificationStore);
    expect(third.classificationRecords).not.toBe(first.classificationRecords);
  });
});
