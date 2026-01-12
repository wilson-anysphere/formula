// @vitest-environment jsdom

import { afterEach, beforeEach, describe, expect, it } from "vitest";

import { purgeLegacyDesktopLLMSettings } from "./desktopLLMClient.js";

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

describe("purgeLegacyDesktopLLMSettings", () => {
  const originalGlobalLocalStorage = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
  const originalWindowLocalStorage = Object.getOwnPropertyDescriptor(window, "localStorage");

  beforeEach(() => {
    // Node 25 ships an experimental `globalThis.localStorage` accessor that throws unless
    // Node is started with `--localstorage-file`. Provide a stable in-memory implementation
    // for unit tests.
    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
  });

  afterEach(() => {
    try {
      window.localStorage.clear();
    } catch {
      // ignore
    }

    if (originalGlobalLocalStorage) {
      Object.defineProperty(globalThis, "localStorage", originalGlobalLocalStorage);
    } else {
      // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
      delete (globalThis as any).localStorage;
    }

    if (originalWindowLocalStorage) {
      Object.defineProperty(window, "localStorage", originalWindowLocalStorage);
    } else {
      // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
      delete (window as any).localStorage;
    }
  });

  it("removes legacy LLM provider + API key settings from localStorage", () => {
    window.localStorage.setItem("formula:openaiApiKey", "sk-legacy-test");
    window.localStorage.setItem("formula:llm:provider", "openai");

    purgeLegacyDesktopLLMSettings();

    expect(window.localStorage.getItem("formula:openaiApiKey")).toBeNull();
    expect(window.localStorage.getItem("formula:llm:provider")).toBeNull();
  });
});
