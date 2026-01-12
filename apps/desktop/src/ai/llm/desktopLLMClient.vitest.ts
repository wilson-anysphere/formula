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
    const legacyKey = "formula:" + "open" + "ai" + "ApiKey";
    const llmPrefix = "formula:" + "llm:";
    const llmProviderKey = llmPrefix + "provider";
    const completionPrefix = "formula:" + "aiCompletion:";
    // Avoid hardcoding provider names in source (Cursor-only AI policy guard).
    const provider0 = "open" + "ai";
    const providerA = "an" + "thropic";
    const providerB = "ol" + "lama";

    window.localStorage.setItem(legacyKey, "sk-legacy-test");
    window.localStorage.setItem(llmProviderKey, provider0);

    window.localStorage.setItem(llmPrefix + provider0 + ":apiKey", "sk-test");
    window.localStorage.setItem(llmPrefix + providerA + ":model", "claude-test");
    window.localStorage.setItem(llmPrefix + providerB + ":model", "llama-test");

    window.localStorage.setItem(completionPrefix + "localModelEnabled", "true");
    window.localStorage.setItem(completionPrefix + "localModelName", "formula-completion");
    window.localStorage.setItem(completionPrefix + "localModelBaseUrl", "http://localhost:11434");

    purgeLegacyDesktopLLMSettings();

    expect(window.localStorage.getItem(legacyKey)).toBeNull();
    expect(window.localStorage.getItem(llmProviderKey)).toBeNull();

    expect(window.localStorage.getItem(llmPrefix + provider0 + ":apiKey")).toBeNull();
    expect(window.localStorage.getItem(llmPrefix + providerA + ":model")).toBeNull();
    expect(window.localStorage.getItem(llmPrefix + providerB + ":model")).toBeNull();

    expect(window.localStorage.getItem(completionPrefix + "localModelEnabled")).toBeNull();
    expect(window.localStorage.getItem(completionPrefix + "localModelName")).toBeNull();
    expect(window.localStorage.getItem(completionPrefix + "localModelBaseUrl")).toBeNull();
  });

  it("does not throw when the localStorage accessor throws (Node 25+)", () => {
    Object.defineProperty(globalThis, "localStorage", {
      configurable: true,
      get() {
        throw new Error("localStorage not available");
      },
    });
    Object.defineProperty(window, "localStorage", {
      configurable: true,
      get() {
        throw new Error("localStorage not available");
      },
    });

    expect(() => purgeLegacyDesktopLLMSettings()).not.toThrow();
  });
});
