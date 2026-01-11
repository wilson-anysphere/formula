// @vitest-environment jsdom

import { beforeEach, expect, test } from "vitest";

import { LocalStorageBinaryStorage } from "../src/store/binaryStorage.js";
import { SqliteVectorStore } from "../src/store/sqliteVectorStore.js";
import { ensureTestLocalStorage } from "./testLocalStorage.js";

ensureTestLocalStorage();

function getTestLocalStorage(): Storage {
  const jsdomStorage = (globalThis as any)?.jsdom?.window?.localStorage as Storage | undefined;
  if (!jsdomStorage) {
    throw new Error("Expected vitest jsdom environment to provide globalThis.jsdom.window.localStorage");
  }
  return jsdomStorage;
}

beforeEach(() => {
  getTestLocalStorage().clear();
});

let sqlJsAvailable = true;
try {
  await import("sql.js");
} catch {
  sqlJsAvailable = false;
}

const maybeTest = sqlJsAvailable ? test : test.skip;

maybeTest("SqliteVectorStore persists and reloads via BinaryStorage", async () => {
  const storage = new LocalStorageBinaryStorage({
    namespace: "formula.test.rag.sqlite",
    workbookId: "sqlite-store",
  });

  const store1 = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: true });
  await store1.upsert([
    { id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb", label: "A" } },
    { id: "b", vector: [0, 1, 0], metadata: { workbookId: "wb", label: "B" } },
  ]);
  await store1.close();

  const store2 = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: false });
  const rec = await store2.get("a");
  expect(rec).not.toBeNull();
  expect(rec?.metadata?.label).toBe("A");

  const hits = await store2.query([1, 0, 0], 1, { workbookId: "wb" });
  expect(hits[0]?.id).toBe("a");
  await store2.close();
});
