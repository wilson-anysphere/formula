import assert from "node:assert/strict";
import test from "node:test";

import http from "node:http";
import https from "node:https";

import { HashEmbedder } from "../src/embedding/hashEmbedder.js";

test("HashEmbedder does not perform network requests", async () => {
  const originalFetch = globalThis.fetch;
  const originalHttpRequest = http.request;
  const originalHttpsRequest = https.request;

  /** @type {string[]} */
  const calls = [];

  try {
    // HashEmbedder should be fully offline; if this ever starts making network
    // requests, treat it as a regression.
    globalThis.fetch = /** @type {any} */ (() => {
      calls.push("fetch");
      throw new Error("Unexpected fetch call from HashEmbedder");
    });

    // Some clients might use node:http(s) directly. Block those too.
    http.request = /** @type {any} */ (() => {
      calls.push("http.request");
      throw new Error("Unexpected http.request call from HashEmbedder");
    });
    https.request = /** @type {any} */ (() => {
      calls.push("https.request");
      throw new Error("Unexpected https.request call from HashEmbedder");
    });

    const embedder = new HashEmbedder({ dimension: 64 });
    const [vec] = await embedder.embedTexts(["hello world"]);
    assert.equal(vec.length, 64);
    assert.deepEqual(calls, []);
  } finally {
    globalThis.fetch = originalFetch;
    http.request = originalHttpRequest;
    https.request = originalHttpsRequest;
  }
});

