import test from "node:test";
import assert from "node:assert/strict";

import { applyNetworkSandbox } from "../src/pyodide-main-thread.js";

test("applyNetworkSandbox is best-effort when fetch/WebSocket are non-writable", (t) => {
  const originalFetchDescriptor = Object.getOwnPropertyDescriptor(globalThis, "fetch");
  const originalWebSocketDescriptor = Object.getOwnPropertyDescriptor(globalThis, "WebSocket");

  const canOverrideFetch = !originalFetchDescriptor || originalFetchDescriptor.configurable === true;
  const canOverrideWebSocket = !originalWebSocketDescriptor || originalWebSocketDescriptor.configurable === true;

  if (!canOverrideFetch || !canOverrideWebSocket) {
    t.skip();
    return;
  }

  const originalFetch = globalThis.fetch;
  const originalWebSocket = globalThis.WebSocket;

  let fetchSetCalls = 0;
  let webSocketSetCalls = 0;

  Object.defineProperty(globalThis, "fetch", {
    configurable: true,
    enumerable: true,
    get() {
      return originalFetch;
    },
    set(_value) {
      fetchSetCalls += 1;
      throw new TypeError("fetch is read-only");
    },
  });

  Object.defineProperty(globalThis, "WebSocket", {
    configurable: true,
    enumerable: true,
    get() {
      return originalWebSocket;
    },
    set(_value) {
      webSocketSetCalls += 1;
      throw new TypeError("WebSocket is read-only");
    },
  });

  try {
    const restore = applyNetworkSandbox({ network: "none" });
    assert.equal(typeof restore, "function");
    assert.ok(fetchSetCalls > 0);
    assert.ok(webSocketSetCalls > 0);
    assert.doesNotThrow(() => restore());
  } finally {
    if (originalFetchDescriptor) {
      Object.defineProperty(globalThis, "fetch", originalFetchDescriptor);
    } else {
      // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
      delete globalThis.fetch;
    }

    if (originalWebSocketDescriptor) {
      Object.defineProperty(globalThis, "WebSocket", originalWebSocketDescriptor);
    } else {
      // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
      delete globalThis.WebSocket;
    }
  }
});

