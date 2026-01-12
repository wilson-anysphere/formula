import assert from "node:assert/strict";
import test from "node:test";

import { CursorTabCompletionClient } from "../src/cursorTabCompletionClient.js";

test("CursorTabCompletionClient sends a structured request body", async () => {
  /** @type {{ url: string, init: any } | null} */
  let seen = null;

  const fetchImpl = async (url, init) => {
    seen = { url, init };
    return {
      ok: true,
      async json() {
        return { completion: "2" };
      },
    };
  };

  const client = new CursorTabCompletionClient({
    baseUrl: "http://example.test",
    fetchImpl,
    timeoutMs: 500,
  });

  const completion = await client.completeTabCompletion({ input: "=1+", cursorPosition: 3, cellA1: "A1" });
  assert.equal(completion, "2");

  assert.equal(seen?.url, "http://example.test/api/ai/tab-completion");
  assert.equal(seen?.init?.method, "POST");
  assert.equal(seen?.init?.headers?.["content-type"], "application/json");
  assert.equal(typeof seen?.init?.signal?.aborted, "boolean");

  const body = JSON.parse(seen?.init?.body ?? "{}");
  assert.deepEqual(body, { input: "=1+", cursorPosition: 3, cellA1: "A1" });
});

test("CursorTabCompletionClient aborts the request when the timeout budget is exceeded", async () => {
  let sawAbort = false;

  const fetchImpl = async (_url, init) => {
    return await new Promise((resolve, reject) => {
      init.signal.addEventListener("abort", () => {
        sawAbort = true;
        const err = new Error("aborted");
        // Match the shape thrown by fetch() in many runtimes.
        err.name = "AbortError";
        reject(err);
      });

      // Never resolve; rely on abort.
      // (Intentionally left pending.)
    });
  };

  const client = new CursorTabCompletionClient({
    baseUrl: "http://example.test",
    fetchImpl,
    timeoutMs: 10,
  });

  const completion = await client.completeTabCompletion({ input: "=1+", cursorPosition: 3, cellA1: "A1" });
  assert.equal(completion, "");
  assert.equal(sawAbort, true);
});

test("CursorTabCompletionClient returns empty string on non-2xx responses", async () => {
  const fetchImpl = async () => {
    return {
      ok: false,
      status: 500,
      async json() {
        return { completion: "should not be read" };
      },
    };
  };

  const client = new CursorTabCompletionClient({
    baseUrl: "http://example.test",
    fetchImpl,
    timeoutMs: 500,
  });

  const completion = await client.completeTabCompletion({ input: "=1+", cursorPosition: 3, cellA1: "A1" });
  assert.equal(completion, "");
});
