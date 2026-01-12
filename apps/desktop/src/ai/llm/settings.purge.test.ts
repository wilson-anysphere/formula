import { beforeEach, describe, expect, it } from "vitest";

import { loadDesktopLLMConfig } from "./settings.js";

class MemoryLocalStorage implements Storage {
  private readonly store = new Map<string, string>();

  get length(): number {
    return this.store.size;
  }

  clear(): void {
    this.store.clear();
  }

  getItem(key: string): string | null {
    return this.store.get(String(key)) ?? null;
  }

  key(index: number): string | null {
    if (index < 0) return null;
    return Array.from(this.store.keys())[index] ?? null;
  }

  removeItem(key: string): void {
    this.store.delete(String(key));
  }

  setItem(key: string, value: string): void {
    this.store.set(String(key), String(value));
  }
}

function getLocalStorageForTest(): Storage {
  try {
    const storage = globalThis.localStorage as Storage | undefined;
    if (storage && typeof storage.getItem === "function" && typeof storage.setItem === "function") {
      return storage;
    }
  } catch {
    // fall through
  }

  const storage = new MemoryLocalStorage();
  Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
  return storage;
}

describe("desktop AI localStorage purge", () => {
  beforeEach(() => {
    const storage = getLocalStorageForTest();
    storage.clear();
  });

  it("removes legacy tab-completion local model keys", () => {
    const storage = getLocalStorageForTest();
    const completionPrefix = "formula:" + "aiCompletion:";
    const enabledKey = completionPrefix + "localModelEnabled";
    const modelKey = completionPrefix + "localModelName";
    const baseUrlKey = completionPrefix + "localModelBaseUrl";

    storage.setItem(enabledKey, "true");
    storage.setItem(modelKey, "formula-completion");
    storage.setItem(baseUrlKey, "http://localhost:11434");

    // Trigger purge (runs on first load of AI settings).
    loadDesktopLLMConfig();

    expect(storage.getItem(enabledKey)).toBeNull();
    expect(storage.getItem(modelKey)).toBeNull();
    expect(storage.getItem(baseUrlKey)).toBeNull();
  });

  it("is fully guarded when the localStorage accessor throws", () => {
    const original = Object.getOwnPropertyDescriptor(globalThis, "localStorage");

    try {
      Object.defineProperty(globalThis, "localStorage", {
        configurable: true,
        get() {
          throw new Error("localStorage not available");
        },
      });

      expect(() => loadDesktopLLMConfig()).not.toThrow();
    } finally {
      if (original) {
        Object.defineProperty(globalThis, "localStorage", original);
      } else {
        // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
        delete (globalThis as any).localStorage;
      }
    }
  });
});
