import assert from "node:assert/strict";
import test from "node:test";

import http from "node:http";
import https from "node:https";
import net from "node:net";
import tls from "node:tls";

test("HashEmbedder does not perform network requests", async () => {
  const originalFetch = globalThis.fetch;
  const originalWebSocket = globalThis.WebSocket;
  const originalHttpRequest = http.request;
  const originalHttpsRequest = https.request;
  const originalNetConnect = net.connect;
  const originalNetCreateConnection = net.createConnection;
  const originalTlsConnect = tls.connect;
  const originalTlsCreateConnection = tls.createConnection;

  /** @type {string[]} */
  const calls = [];

  try {
    // HashEmbedder should be fully offline; if this ever starts making network
    // requests, treat it as a regression.
    globalThis.fetch = /** @type {any} */ (() => {
      calls.push("fetch");
      throw new Error("Unexpected fetch call from HashEmbedder");
    });

    globalThis.WebSocket = /** @type {any} */ (class {
      constructor() {
        calls.push("WebSocket");
        throw new Error("Unexpected WebSocket call from HashEmbedder");
      }
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

    // Also block low-level socket creation in case a future implementation uses
    // custom networking without going through fetch/http.
    const connectError = (name) => {
      calls.push(name);
      throw new Error(`Unexpected ${name} call from HashEmbedder`);
    };

    net.connect = /** @type {any} */ ((..._args) => connectError("net.connect"));
    net.createConnection = /** @type {any} */ ((..._args) => connectError("net.createConnection"));
    tls.connect = /** @type {any} */ ((..._args) => connectError("tls.connect"));
    tls.createConnection = /** @type {any} */ ((..._args) => connectError("tls.createConnection"));

    // Import after stubbing so a future implementation can't capture references
    // to the real network functions at module init time.
    const { HashEmbedder } = await import("../src/embedding/hashEmbedder.js");

    const embedder = new HashEmbedder({ dimension: 64 });
    const [vec] = await embedder.embedTexts(["hello world"]);
    assert.equal(vec.length, 64);
    assert.deepEqual(calls, []);
  } finally {
    globalThis.fetch = originalFetch;
    globalThis.WebSocket = originalWebSocket;
    http.request = originalHttpRequest;
    https.request = originalHttpsRequest;
    net.connect = originalNetConnect;
    net.createConnection = originalNetCreateConnection;
    tls.connect = originalTlsConnect;
    tls.createConnection = originalTlsCreateConnection;
  }
});
