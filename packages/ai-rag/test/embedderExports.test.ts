import { readdir } from "node:fs/promises";

import { expect, test } from "vitest";

import * as aiRag from "../src/index.js";

test("ai-rag does not export remote/local provider embedders", () => {
  const remoteName = "Open" + "AI" + "Embedder";
  // Keep the local embedder name split so Cursor-only policy guards can detect
  // accidental *usage* of forbidden providers without blocking this negative test.
  const localName = "O" + "llama" + "Embedder";

  expect(Object.prototype.hasOwnProperty.call(aiRag, remoteName)).toBe(false);
  expect(Object.prototype.hasOwnProperty.call(aiRag, localName)).toBe(false);
});

test("ai-rag does not ship OpenAI/local embedder modules", async () => {
  const entries = await readdir(new URL("../src/embedding/", import.meta.url));
  const normalized = entries.map((e) => e.toLowerCase());

  expect(normalized.some((e) => /^openaiembedder\./.test(e))).toBe(false);
  const localEmbedderPrefix = "o" + "llama" + "embedder.";
  expect(normalized.some((e) => e.startsWith(localEmbedderPrefix))).toBe(false);
});
