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
  assert.equal(
    seen?.init?.headers?.["Content-Type"] ?? seen?.init?.headers?.["content-type"],
    "application/json",
  );
  assert.equal(seen?.init?.credentials, "include");
  assert.equal(typeof seen?.init?.signal?.aborted, "boolean");

  const body = JSON.parse(seen?.init?.body ?? "{}");
  assert.deepEqual(body, { input: "=1+", cursorPosition: 3, cellA1: "A1" });
});

test("CursorTabCompletionClient accepts a fully-qualified endpoint URL (does not append /api/ai/tab-completion twice)", async () => {
  /** @type {string | null} */
  let urlSeen = null;
  const fetchImpl = async (url) => {
    urlSeen = url;
    return {
      ok: true,
      async json() {
        return { completion: "ok" };
      },
    };
  };

  const client = new CursorTabCompletionClient({
    baseUrl: "http://example.test/api/ai/tab-completion",
    fetchImpl,
    timeoutMs: 500,
  });

  const completion = await client.completeTabCompletion({ input: "=", cursorPosition: 1, cellA1: "A1" });
  assert.equal(completion, "ok");
  assert.equal(urlSeen, "http://example.test/api/ai/tab-completion");
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

test(
  "CursorTabCompletionClient enforces the timeout budget even if fetchImpl ignores AbortSignal",
  { timeout: 1000 },
  async () => {
    let fetchCalls = 0;

    const fetchImpl = async () => {
      fetchCalls += 1;
      // Intentionally ignore init.signal and never resolve.
      return await new Promise(() => {});
    };

    const client = new CursorTabCompletionClient({
      baseUrl: "http://example.test",
      fetchImpl,
      timeoutMs: 10,
    });

    const completion = await client.completeTabCompletion({ input: "=1+", cursorPosition: 3, cellA1: "A1" });
    assert.equal(completion, "");
    assert.equal(fetchCalls, 1);
  },
);

test("CursorTabCompletionClient merges headers from getAuthHeaders", async () => {
  /** @type {Record<string, string> | null} */
  let headersSeen = null;

  const fetchImpl = async (_url, init) => {
    headersSeen = init.headers;
    return {
      ok: true,
      async json() {
        return { completion: "ok" };
      },
    };
  };

  const client = new CursorTabCompletionClient({
    baseUrl: "http://example.test",
    fetchImpl,
    timeoutMs: 500,
    getAuthHeaders() {
      return { "x-cursor-test-auth": "yes" };
    },
  });

  const completion = await client.completeTabCompletion({ input: "=", cursorPosition: 1, cellA1: "A1" });
  assert.equal(completion, "ok");
  assert.equal(headersSeen?.["x-cursor-test-auth"], "yes");
  assert.equal(headersSeen?.["Content-Type"] ?? headersSeen?.["content-type"], "application/json");
});

test("CursorTabCompletionClient forces Content-Type to application/json even if getAuthHeaders sets it", async () => {
  /** @type {Record<string, string> | null} */
  let headersSeen = null;

  const fetchImpl = async (_url, init) => {
    headersSeen = init.headers;
    return {
      ok: true,
      async json() {
        return { completion: "ok" };
      },
    };
  };

  const client = new CursorTabCompletionClient({
    baseUrl: "http://example.test",
    fetchImpl,
    timeoutMs: 500,
    getAuthHeaders() {
      return { "Content-Type": "text/plain", "x-cursor-test-auth": "yes" };
    },
  });

  const completion = await client.completeTabCompletion({ input: "=", cursorPosition: 1, cellA1: "A1" });
  assert.equal(completion, "ok");
  assert.equal(headersSeen?.["x-cursor-test-auth"], "yes");
  const hasUpper = Object.prototype.hasOwnProperty.call(headersSeen ?? {}, "Content-Type");
  const hasLower = Object.prototype.hasOwnProperty.call(headersSeen ?? {}, "content-type");
  assert.equal(
    (hasUpper ? 1 : 0) + (hasLower ? 1 : 0),
    1,
    "Expected exactly one Content-Type header key (no duplicates by casing)",
  );
  assert.equal(headersSeen?.["Content-Type"] ?? headersSeen?.["content-type"], "application/json");
});

test("CursorTabCompletionClient resolves to empty string when an external AbortSignal is aborted", async () => {
  let sawAbort = false;

  const fetchImpl = async (_url, init) => {
    return await new Promise((_resolve, reject) => {
      init.signal.addEventListener("abort", () => {
        sawAbort = true;
        const err = new Error("aborted");
        err.name = "AbortError";
        reject(err);
      });
    });
  };

  const client = new CursorTabCompletionClient({
    baseUrl: "http://example.test",
    fetchImpl,
    timeoutMs: 500,
  });

  const controller = new AbortController();
  const promise = client.completeTabCompletion({
    input: "=1+",
    cursorPosition: 3,
    cellA1: "A1",
    signal: controller.signal,
  });

  controller.abort();

  const completion = await promise;
  assert.equal(completion, "");
  assert.equal(sawAbort, true);
});

test("CursorTabCompletionClient returns empty string immediately when the external AbortSignal is already aborted", async () => {
  let fetchCalls = 0;
  let authCalls = 0;

  const fetchImpl = async () => {
    fetchCalls += 1;
    return { ok: true, async json() {} };
  };

  const client = new CursorTabCompletionClient({
    baseUrl: "http://example.test",
    fetchImpl,
    timeoutMs: 500,
    getAuthHeaders: () => {
      authCalls += 1;
      return { "x-cursor-test-auth": "yes" };
    },
  });

  const controller = new AbortController();
  controller.abort();

  const completion = await client.completeTabCompletion({
    input: "=1+",
    cursorPosition: 3,
    cellA1: "A1",
    signal: controller.signal,
  });

  assert.equal(completion, "");
  assert.equal(fetchCalls, 0);
  assert.equal(authCalls, 0);
});

test(
  "CursorTabCompletionClient applies the timeout budget while awaiting getAuthHeaders",
  { timeout: 1000 },
  async () => {
  let fetchCalls = 0;

  const fetchImpl = async () => {
    fetchCalls += 1;
    return { ok: true, async json() {} };
  };

  const client = new CursorTabCompletionClient({
    baseUrl: "http://example.test",
    fetchImpl,
    timeoutMs: 10,
    getAuthHeaders: () => new Promise(() => {}), // never resolves
  });

  const completion = await client.completeTabCompletion({ input: "=1+2", cursorPosition: 5, cellA1: "A1" });

  assert.equal(completion, "");
  assert.equal(fetchCalls, 0);
  },
);

test("CursorTabCompletionClient awaits async getAuthHeaders", async () => {
  /** @type {Record<string, string> | null} */
  let headersSeen = null;
  let authCalls = 0;

  const fetchImpl = async (_url, init) => {
    headersSeen = init.headers;
    return {
      ok: true,
      async json() {
        return { completion: "ok" };
      },
    };
  };

  const client = new CursorTabCompletionClient({
    baseUrl: "http://example.test",
    fetchImpl,
    timeoutMs: 500,
    getAuthHeaders: async () => {
      authCalls += 1;
      await Promise.resolve();
      return { "x-cursor-test-auth": "async" };
    },
  });

  const completion = await client.completeTabCompletion({ input: "=", cursorPosition: 1, cellA1: "A1" });
  assert.equal(completion, "ok");
  assert.equal(authCalls, 1);
  assert.equal(headersSeen?.["x-cursor-test-auth"], "async");
});

test("CursorTabCompletionClient de-dupes Authorization header case-insensitively", async () => {
  /** @type {Record<string, string> | null} */
  let headersSeen = null;

  const fetchImpl = async (_url, init) => {
    headersSeen = init.headers;
    return {
      ok: true,
      async json() {
        return { completion: "ok" };
      },
    };
  };

  const client = new CursorTabCompletionClient({
    baseUrl: "http://example.test",
    fetchImpl,
    timeoutMs: 500,
    getAuthHeaders() {
      return {
        Authorization: "Bearer first",
        authorization: "Bearer second",
      };
    },
  });

  const completion = await client.completeTabCompletion({ input: "=", cursorPosition: 1, cellA1: "A1" });
  assert.equal(completion, "ok");
  assert.equal(headersSeen?.Authorization, "Bearer second");
  assert.ok(!("authorization" in (headersSeen ?? {})));
});

test(
  "CursorTabCompletionClient resolves to empty string when aborted while awaiting getAuthHeaders",
  { timeout: 1000 },
  async () => {
    let fetchCalls = 0;

    const fetchImpl = async () => {
      fetchCalls += 1;
      return { ok: true, async json() {} };
    };

    const client = new CursorTabCompletionClient({
      baseUrl: "http://example.test",
      fetchImpl,
      timeoutMs: 500,
      getAuthHeaders: () => new Promise(() => {}), // never resolves
    });

    const controller = new AbortController();
    const promise = client.completeTabCompletion({
      input: "=1+2",
      cursorPosition: 4,
      cellA1: "A1",
      signal: controller.signal,
    });

    controller.abort();

    const completion = await promise;
    assert.equal(completion, "");
    assert.equal(fetchCalls, 0);
  },
);

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
