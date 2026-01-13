// @vitest-environment jsdom
import { describe, expect, it } from "vitest";

import { LocalStorageBinaryStorage } from "@formula/ai-audit/browser";

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

describe("LocalStorageBinaryStorage (browser base64 path)", () => {
  it(
    "round-trips a multi-MB blob without using Buffer (chunked btoa/atob)",
    async () => {
      const originalBuffer = (globalThis as any).Buffer;
      const originalLocalStorage = Object.getOwnPropertyDescriptor(globalThis, "localStorage");

      Object.defineProperty(globalThis, "localStorage", { value: new MemoryLocalStorage(), configurable: true });

      try {
        const storage = new LocalStorageBinaryStorage("ai_audit_large_blob_test");

        const size = 2 * 1024 * 1024; // 2MiB
        const data = new Uint8Array(size);
        for (let i = 0; i < data.length; i++) data[i] = i & 0xff;

        try {
          (globalThis as any).Buffer = undefined;

          await storage.save(data);
          const loaded = await storage.load();

          expect(loaded).not.toBeNull();
          expect(loaded!.byteLength).toBe(data.byteLength);
          expect(loaded).toEqual(data);
        } finally {
          (globalThis as any).Buffer = originalBuffer;
        }
      } finally {
        if (originalLocalStorage) {
          Object.defineProperty(globalThis, "localStorage", originalLocalStorage);
        } else {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          delete (globalThis as any).localStorage;
        }
      }
    },
    60_000
  );
});

