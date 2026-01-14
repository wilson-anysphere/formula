// @vitest-environment jsdom

import { afterAll, beforeAll, beforeEach, expect, test } from "vitest";

import { ChunkedLocalStorageBinaryStorage, LocalStorageBinaryStorage } from "../src/store/binaryStorage.js";
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

function listKeysWithPrefix(storage: Storage, prefix: string): string[] {
  const keys: string[] = [];
  for (let i = 0; i < storage.length; i += 1) {
    const key = storage.key(i);
    if (key?.startsWith(prefix)) keys.push(key);
  }
  return keys.sort();
}

test("ChunkedLocalStorageBinaryStorage round-trips a large payload across multiple localStorage keys", async () => {
  const storage = new ChunkedLocalStorageBinaryStorage({
    namespace: "formula.test.rag",
    workbookId: "wb-123",
    chunkSizeChars: 64,
  });

  const bytes = new Uint8Array(2048);
  for (let i = 0; i < bytes.length; i += 1) bytes[i] = i % 256;

  await storage.save(bytes);

  const rawMeta = getTestLocalStorage().getItem(`${storage.key}:meta`);
  expect(rawMeta).toBeTypeOf("string");
  const meta = JSON.parse(rawMeta ?? "{}");
  expect(meta.chunks).toBeGreaterThan(1);

  // Sanity check that chunks were written as separate keys.
  const keys = listKeysWithPrefix(getTestLocalStorage(), `${storage.key}:`);
  expect(keys).toContain(`${storage.key}:meta`);
  expect(keys).toContain(`${storage.key}:0`);

  const loaded = await storage.load();
  expect(loaded).toBeInstanceOf(Uint8Array);
  expect(Array.from(loaded ?? [])).toEqual(Array.from(bytes));
});

test("ChunkedLocalStorageBinaryStorage overwrites and removes old chunks when saving a smaller payload", async () => {
  const storage = new ChunkedLocalStorageBinaryStorage({
    namespace: "formula.test.rag",
    workbookId: "wb-123",
    chunkSizeChars: 64,
  });

  const large = new Uint8Array(4096);
  for (let i = 0; i < large.length; i += 1) large[i] = (i * 17) % 256;

  await storage.save(large);
  const firstMeta = JSON.parse(getTestLocalStorage().getItem(`${storage.key}:meta`) ?? "{}");
  expect(firstMeta.chunks).toBeGreaterThan(1);

  const small = new Uint8Array([1, 2, 3, 4, 5, 255]);
  await storage.save(small);

  const secondMeta = JSON.parse(getTestLocalStorage().getItem(`${storage.key}:meta`) ?? "{}");
  expect(secondMeta.chunks).toBe(1);

  const keysAfter = listKeysWithPrefix(getTestLocalStorage(), `${storage.key}:`);
  expect(keysAfter).toEqual([`${storage.key}:0`, `${storage.key}:meta`]);

  const loaded = await storage.load();
  expect(Array.from(loaded ?? [])).toEqual(Array.from(small));
});

test("ChunkedLocalStorageBinaryStorage round-trips multi-MB payloads via browser base64 streaming path", async () => {
  const storage = new ChunkedLocalStorageBinaryStorage({
    namespace: "formula.test.rag",
    workbookId: "large-chunked",
    chunkSizeChars: 16_384,
  });

  // ~1MB payload; large enough to exercise the chunked btoa/atob code-path without
  // creating excessive test runtime.
  const size = 1024 * 1024;
  const bytes = new Uint8Array(size);
  for (let i = 0; i < bytes.length; i += 1) bytes[i] = i % 256;

  const originalBuffer = (globalThis as any).Buffer;
  try {
    // Force the browser code-path (streaming btoa/atob conversion).
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

test("ChunkedLocalStorageBinaryStorage clears corrupted base64 chunk payloads on load", async () => {
  const storage = new ChunkedLocalStorageBinaryStorage({
    namespace: "formula.test.rag",
    workbookId: "corrupt-chunk",
    chunkSizeChars: 8,
  });

  const originalBuffer = (globalThis as any).Buffer;
  try {
    // Force the browser decode path so invalid base64 throws (Buffer's base64 decoder is permissive).
    (globalThis as any).Buffer = undefined;

    const ls = getTestLocalStorage();
    ls.setItem(`${storage.key}:meta`, JSON.stringify({ chunks: 2 }));
    ls.setItem(`${storage.key}:0`, "not-base64");
    ls.setItem(`${storage.key}:1`, "AAAA");

    const loaded = await storage.load();
    expect(loaded).toBeNull();

    // Corrupted payloads should be cleared so the next load is a clean miss.
    expect(listKeysWithPrefix(ls, `${storage.key}:`)).toEqual([]);
    expect(ls.getItem(storage.key)).toBeNull();
  } finally {
    getTestLocalStorage().clear();
    (globalThis as any).Buffer = originalBuffer;
  }
});

test("ChunkedLocalStorageBinaryStorage clears absurd chunk counts in meta (corruption guard)", async () => {
  const storage = new ChunkedLocalStorageBinaryStorage({
    namespace: "formula.test.rag",
    workbookId: "corrupt-meta-huge",
    chunkSizeChars: 8,
  });

  const ls = getTestLocalStorage();
  ls.setItem(`${storage.key}:meta`, JSON.stringify({ chunks: 2_000_000 }));
  ls.setItem(`${storage.key}:0`, "AAAA");

  // Ensure the corruption guard short-circuits before trying to read chunk keys.
  const originalGetItem = (Storage.prototype as any).getItem as (...args: any[]) => any;
  (Storage.prototype as any).getItem = function (key: string) {
    if (key === `${storage.key}:0`) throw new Error("Should not read chunk keys when meta is corrupted");
    return originalGetItem.call(this, key);
  };

  try {
    const loaded = await storage.load();
    expect(loaded).toBeNull();
    expect(listKeysWithPrefix(ls, `${storage.key}:`)).toEqual([]);
  } finally {
    (Storage.prototype as any).getItem = originalGetItem;
  }
});

test("ChunkedLocalStorageBinaryStorage works with minimal Storage implementations (no key/length/removeItem)", async () => {
  const jsdomWindow = (globalThis as any)?.jsdom?.window as Window | undefined;
  if (!jsdomWindow) throw new Error("Expected vitest jsdom environment to provide globalThis.jsdom.window");
  const originalLocalStorage = Object.getOwnPropertyDescriptor(jsdomWindow, "localStorage");

  const data = new Map<string, string>();
  const minimalStorage = {
    getItem(key: string) {
      return data.get(key) ?? null;
    },
    setItem(key: string, value: string) {
      data.set(key, String(value));
    },
  } as any;

  Object.defineProperty(jsdomWindow, "localStorage", { value: minimalStorage, configurable: true });

  try {
    const storage = new ChunkedLocalStorageBinaryStorage({
      namespace: "formula.test.rag",
      workbookId: "minimal-storage",
      chunkSizeChars: 8,
    });

    const bytes = new Uint8Array([1, 2, 3, 4, 5, 255]);
    await storage.save(bytes);
    const loaded = await storage.load();
    expect(Array.from(loaded ?? [])).toEqual(Array.from(bytes));

    // Should not throw even though removeItem/key/length are missing.
    await expect(storage.remove()).resolves.toBeUndefined();
  } finally {
    if (originalLocalStorage) {
      Object.defineProperty(jsdomWindow, "localStorage", originalLocalStorage);
    } else {
      // eslint-disable-next-line no-undef
      delete (jsdomWindow as any).localStorage;
    }
  }
});

test("ChunkedLocalStorageBinaryStorage does not corrupt existing data if a save fails mid-write", async () => {
  const storage = new ChunkedLocalStorageBinaryStorage({
    namespace: "formula.test.rag",
    workbookId: "wb-rollback",
    chunkSizeChars: 4,
  });

  const original = new Uint8Array(128);
  for (let i = 0; i < original.length; i += 1) original[i] = i % 256;
  await storage.save(original);

  // Fail when writing the second chunk (`:1`) to simulate a quota/storage error after
  // chunk 0 has already been overwritten.
  const originalSetItem = (Storage.prototype as any).setItem as (...args: any[]) => any;
  (Storage.prototype as any).setItem = function (key: string, value: string) {
    if (key === `${storage.key}:1`) {
      throw new Error("QuotaExceededError");
    }
    return originalSetItem.call(this, key, value);
  };

  try {
    const next = new Uint8Array(128);
    for (let i = 0; i < next.length; i += 1) next[i] = (i * 17) % 256;
    await expect(storage.save(next)).rejects.toThrow(/QuotaExceededError/);
  } finally {
    (Storage.prototype as any).setItem = originalSetItem;
  }

  const loaded = await storage.load();
  expect(Array.from(loaded ?? [])).toEqual(Array.from(original));
});

test("ChunkedLocalStorageBinaryStorage remove deletes meta + chunk keys", async () => {
  const storage = new ChunkedLocalStorageBinaryStorage({
    namespace: "formula.test.rag",
    workbookId: "wb-remove",
    chunkSizeChars: 64,
  });

  const bytes = new Uint8Array(2048);
  for (let i = 0; i < bytes.length; i += 1) bytes[i] = i % 256;

  await storage.save(bytes);
  expect(listKeysWithPrefix(getTestLocalStorage(), `${storage.key}:`).length).toBeGreaterThan(1);

  await storage.remove();
  expect(await storage.load()).toBeNull();
  expect(listKeysWithPrefix(getTestLocalStorage(), `${storage.key}:`)).toEqual([]);
});

test("LocalStorageBinaryStorage remove deletes persisted bytes", async () => {
  const storage = new LocalStorageBinaryStorage({ namespace: "formula.test.rag", workbookId: "remove-me" });
  const bytes = new Uint8Array([9, 8, 7]);

  await storage.save(bytes);
  expect(getTestLocalStorage().getItem(storage.key)).toBeTypeOf("string");

  await storage.remove();
  expect(await storage.load()).toBeNull();
  expect(getTestLocalStorage().getItem(storage.key)).toBeNull();
});

test("ChunkedLocalStorageBinaryStorage can load legacy single-key LocalStorageBinaryStorage values (and migrates them)", async () => {
  const legacy = new LocalStorageBinaryStorage({ namespace: "formula.test.rag", workbookId: "legacy" });
  const bytes = new Uint8Array([1, 2, 3, 4, 5, 255]);
  await legacy.save(bytes);
  expect(getTestLocalStorage().getItem(legacy.key)).toBeTypeOf("string");

  const chunked = new ChunkedLocalStorageBinaryStorage({
    namespace: "formula.test.rag",
    workbookId: "legacy",
    chunkSizeChars: 4,
  });

  const loaded = await chunked.load();
  expect(Array.from(loaded ?? [])).toEqual(Array.from(bytes));

  // Migration should remove the legacy single-key entry and write chunked keys instead.
  expect(getTestLocalStorage().getItem(chunked.key)).toBeNull();
  const metaRaw = getTestLocalStorage().getItem(`${chunked.key}:meta`);
  expect(metaRaw).toBeTypeOf("string");
  const meta = JSON.parse(metaRaw ?? "{}");
  expect(meta.chunks).toBeGreaterThan(0);
  expect(getTestLocalStorage().getItem(`${chunked.key}:0`)).toBeTypeOf("string");
});

test("ChunkedLocalStorageBinaryStorage can load legacy misaligned chunk boundaries (and migrates them)", async () => {
  const namespace = "formula.test.rag";
  const workbookId = "misaligned";
  const key = `${namespace}:${workbookId}`;

  const bytes = new Uint8Array(64);
  for (let i = 0; i < bytes.length; i += 1) bytes[i] = (i * 31) % 256;
  const encoded = Buffer.from(bytes).toString("base64");

  // Simulate a previous buggy implementation that used a non-base64-aligned chunk size.
  const oldChunkSize = 5; // not divisible by 4
  const oldChunks = Math.ceil(encoded.length / oldChunkSize);
  const ls = getTestLocalStorage();
  ls.setItem(`${key}:meta`, JSON.stringify({ chunks: oldChunks }));
  for (let i = 0; i < oldChunks; i += 1) {
    const start = i * oldChunkSize;
    ls.setItem(`${key}:${i}`, encoded.slice(start, start + oldChunkSize));
  }

  const storage = new ChunkedLocalStorageBinaryStorage({ namespace, workbookId, chunkSizeChars: 8 });
  const loaded = await storage.load();
  expect(Array.from(loaded ?? [])).toEqual(Array.from(bytes));

  // Migration should rewrite chunks with a base64-aligned chunkSizeChars (8 -> aligned) and
  // ensure each stored chunk is independently decodable (length % 4 === 0).
  const metaRaw = ls.getItem(`${key}:meta`);
  expect(metaRaw).toBeTypeOf("string");
  const meta = JSON.parse(metaRaw ?? "{}");
  expect(typeof meta.chunks).toBe("number");
  expect(meta.chunks).toBeGreaterThan(0);

  for (let i = 0; i < meta.chunks; i += 1) {
    const part = ls.getItem(`${key}:${i}`);
    expect(part).toBeTypeOf("string");
    expect((part as string).length % 4).toBe(0);
  }
});
