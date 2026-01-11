// @vitest-environment jsdom

import { beforeEach, expect, test } from "vitest";

import { LocalStorageBinaryStorage } from "../src/store/binaryStorage.js";

beforeEach(() => {
  localStorage.clear();
});

test("LocalStorageBinaryStorage round-trips bytes (key is namespaced per workbook)", async () => {
  const storage = new LocalStorageBinaryStorage({ namespace: "formula.test.rag", workbookId: "wb-123" });
  expect(storage.key).toContain("wb-123");

  const bytes = new Uint8Array([1, 2, 3, 4, 255]);
  await storage.save(bytes);

  const raw = localStorage.getItem(storage.key);
  expect(raw).toBeTypeOf("string");

  const loaded = await storage.load();
  expect(loaded).toBeInstanceOf(Uint8Array);
  expect(Array.from(loaded ?? [])).toEqual(Array.from(bytes));
});

test("LocalStorageBinaryStorage returns null when missing", async () => {
  const storage = new LocalStorageBinaryStorage({ namespace: "formula.test.rag", workbookId: "missing" });
  expect(await storage.load()).toBeNull();
});

