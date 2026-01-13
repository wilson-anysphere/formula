// @vitest-environment jsdom

import { afterAll, beforeAll, beforeEach, expect, test } from "vitest";

import { LocalStorageBinaryStorage } from "../src/store/binaryStorage.js";
import { ensureTestLocalStorage } from "./testLocalStorage.js";

ensureTestLocalStorage();

function getTestLocalStorage(): Storage {
  // Vitest's jsdom environment exposes the actual JSDOM instance on `globalThis.jsdom`.
  // Node 25 also exposes an experimental `globalThis.localStorage` accessor that can
  // throw unless Node is started with `--localstorage-file`, so we avoid it here.
  const jsdomStorage = (globalThis as any)?.jsdom?.window?.localStorage as Storage | undefined;
  if (!jsdomStorage) {
    throw new Error("Expected vitest jsdom environment to provide globalThis.jsdom.window.localStorage");
  }
  return jsdomStorage;
}

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

const originalLocalStorage = Object.getOwnPropertyDescriptor(globalThis, "localStorage");

beforeAll(() => {
  // Node 25 ships an experimental `localStorage` accessor that throws unless Node is
  // started with `--localstorage-file`. Force a deterministic in-memory storage for tests.
  Object.defineProperty(globalThis, "localStorage", { value: new MemoryLocalStorage(), configurable: true });
});

afterAll(() => {
  if (originalLocalStorage) {
    Object.defineProperty(globalThis, "localStorage", originalLocalStorage);
  }
});

beforeEach(() => {
  getTestLocalStorage().clear();
});

test("LocalStorageBinaryStorage round-trips bytes (key is namespaced per workbook)", async () => {
  const storage = new LocalStorageBinaryStorage({ namespace: "formula.test.rag", workbookId: "wb-123" });
  expect(storage.key).toContain("wb-123");

  const bytes = new Uint8Array([1, 2, 3, 4, 255]);
  await storage.save(bytes);

  const raw = getTestLocalStorage().getItem(storage.key);
  expect(raw).toBeTypeOf("string");

  const loaded = await storage.load();
  expect(loaded).toBeInstanceOf(Uint8Array);
  expect(Array.from(loaded ?? [])).toEqual(Array.from(bytes));
});

test("LocalStorageBinaryStorage returns null when missing", async () => {
  const storage = new LocalStorageBinaryStorage({ namespace: "formula.test.rag", workbookId: "missing" });
  expect(await storage.load()).toBeNull();
});

test("LocalStorageBinaryStorage round-trips multi-MB payloads via browser base64 fallback", async () => {
  const storage = new LocalStorageBinaryStorage({ namespace: "formula.test.rag", workbookId: "large" });

  // ~3MB payload (base64-encoded value is ~4MB). Large enough to stress the browser
  // base64 path without exceeding typical localStorage quotas in jsdom.
  const size = 3 * 1024 * 1024;
  const bytes = new Uint8Array(size);
  for (let i = 0; i < bytes.length; i += 1) bytes[i] = i % 256;

  const originalBuffer = (globalThis as any).Buffer;
  try {
    // Force the browser code-path (chunked btoa/atob conversion).
    (globalThis as any).Buffer = undefined;

    await storage.save(bytes);
    const loaded = await storage.load();

    expect(loaded).toBeInstanceOf(Uint8Array);
    expect(loaded?.byteLength).toBe(bytes.byteLength);

    // Compare without creating huge intermediate arrays / diffs.
    const a = loaded as Uint8Array;
    let equal = a.length === bytes.length;
    for (let i = 0; equal && i < a.length; i += 1) {
      if (a[i] !== bytes[i]) equal = false;
    }
    expect(equal).toBe(true);
  } finally {
    getTestLocalStorage().clear();
    (globalThis as any).Buffer = originalBuffer;
  }
});
