import assert from "node:assert/strict";
import test from "node:test";

import { ContextManager } from "../src/contextManager.js";

test("buildWorkbookContext aborts while awaiting query embedding", async () => {
  const abortController = new AbortController();
  let embedCalled = false;

  const cm = new ContextManager({
    tokenBudgetTokens: 1_000,
    workbookRag: {
      vectorStore: {
        query: async () => [],
      },
      embedder: {
        embedTexts: async () => {
          embedCalled = true;
          // Never resolve unless aborted.
          return new Promise(() => {});
        },
      },
      topK: 3,
    },
  });

  const promise = cm.buildWorkbookContext({
    workbook: { id: "wb_abort_embed", sheets: [] },
    query: "hello",
    skipIndexing: true,
    signal: abortController.signal,
  });

  assert.equal(embedCalled, true);
  abortController.abort();

  await assert.rejects(promise, { name: "AbortError" });
});

test("buildWorkbookContext aborts while awaiting vectorStore.query", async () => {
  const abortController = new AbortController();
  let queryCalled = false;
  let resolveQueryStarted = () => {};
  const queryStarted = new Promise((resolve) => {
    resolveQueryStarted = resolve;
  });

  const cm = new ContextManager({
    tokenBudgetTokens: 1_000,
    workbookRag: {
      vectorStore: {
        query: async () => {
          queryCalled = true;
          resolveQueryStarted();
          // Never resolve unless aborted.
          return new Promise(() => {});
        },
      },
      embedder: {
        embedTexts: async () => [new Float32Array(8)],
      },
      topK: 3,
    },
  });

  const promise = cm.buildWorkbookContext({
    workbook: { id: "wb_abort_query", sheets: [] },
    query: "hello",
    skipIndexing: true,
    signal: abortController.signal,
  });

  await queryStarted;
  assert.equal(queryCalled, true);
  abortController.abort();

  await assert.rejects(promise, { name: "AbortError" });
});
