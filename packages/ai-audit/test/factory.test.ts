import { afterEach, describe, expect, it } from "vitest";

import { createDefaultAIAuditStore } from "../src/factory.js";
import { LocalStorageAIAuditStore } from "../src/local-storage-store.js";
import { MemoryAIAuditStore } from "../src/memory-store.js";

class MemoryLocalStorage implements Storage {
  #data = new Map<string, string>();

  get length(): number {
    return this.#data.size;
  }

  clear(): void {
    this.#data.clear();
  }

  getItem(key: string): string | null {
    return this.#data.get(key) ?? null;
  }

  key(index: number): string | null {
    return Array.from(this.#data.keys())[index] ?? null;
  }

  removeItem(key: string): void {
    this.#data.delete(key);
  }

  setItem(key: string, value: string): void {
    this.#data.set(key, value);
  }
}

const originalWindowDescriptor = Object.getOwnPropertyDescriptor(globalThis, "window");
const originalLocalStorageDescriptor = Object.getOwnPropertyDescriptor(globalThis, "localStorage");

function restoreGlobals() {
  if (originalWindowDescriptor) Object.defineProperty(globalThis, "window", originalWindowDescriptor);
  else {
    // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
    delete (globalThis as any).window;
  }

  if (originalLocalStorageDescriptor) Object.defineProperty(globalThis, "localStorage", originalLocalStorageDescriptor);
  else {
    // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
    delete (globalThis as any).localStorage;
  }
}

describe("createDefaultAIAuditStore", () => {
  afterEach(() => {
    restoreGlobals();
  });

  it("prefers LocalStorageAIAuditStore in a browser-like runtime", async () => {
    const storage = new MemoryLocalStorage();
    Object.defineProperty(globalThis, "window", { value: { localStorage: storage }, configurable: true });

    const store = await createDefaultAIAuditStore({ retention: { max_entries: 42 } });

    expect(store).toBeInstanceOf(LocalStorageAIAuditStore);
    expect((store as LocalStorageAIAuditStore).maxEntries).toBe(42);
  });

  it("falls back to MemoryAIAuditStore when localStorage access throws", async () => {
    const win: any = {};
    Object.defineProperty(win, "localStorage", {
      configurable: true,
      get() {
        throw new Error("no localStorage");
      }
    });
    Object.defineProperty(globalThis, "window", { value: win, configurable: true });

    const store = await createDefaultAIAuditStore();
    expect(store).toBeInstanceOf(MemoryAIAuditStore);
  });

  it("defaults to MemoryAIAuditStore in Node runtimes (no window)", async () => {
    // Ensure we don't accidentally treat the test environment as browser-like.
    // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
    delete (globalThis as any).window;

    const store = await createDefaultAIAuditStore();
    expect(store).toBeInstanceOf(MemoryAIAuditStore);
  });
});

