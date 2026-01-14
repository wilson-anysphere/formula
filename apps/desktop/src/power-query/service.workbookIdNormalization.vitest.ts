import { afterEach, beforeEach, describe, expect, it } from "vitest";

import { loadQueriesFromStorage, saveQueriesToStorage } from "./service.ts";

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
    },
  } as Storage;
}

describe("Power query storage helpers", () => {
  const originalLocalStorage = Object.getOwnPropertyDescriptor(globalThis, "localStorage");

  beforeEach(() => {
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: createInMemoryLocalStorage() });
  });

  afterEach(() => {
    if (originalLocalStorage) {
      Object.defineProperty(globalThis, "localStorage", originalLocalStorage);
    } else {
      // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
      delete (globalThis as any).localStorage;
    }
  });

  it("normalizes workbook ids by trimming before reading/writing localStorage keys", () => {
    const queries = [
      {
        id: "q1",
        name: "Query 1",
        source: { type: "range", range: { values: [] } },
        steps: [],
      },
    ];

    saveQueriesToStorage("  workbook-1  ", queries as any);

    // Trimmed id should find the persisted queries.
    expect(loadQueriesFromStorage("workbook-1")).toEqual(queries);

    // The normalized key should not include whitespace.
    expect((globalThis as any).localStorage.getItem("formula.desktop.powerQuery.queries:workbook-1")).not.toBeNull();
    expect((globalThis as any).localStorage.getItem("formula.desktop.powerQuery.queries:  workbook-1  ")).toBeNull();
  });
});

