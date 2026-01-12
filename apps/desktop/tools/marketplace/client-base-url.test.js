import test from "node:test";
import assert from "node:assert/strict";

import { MarketplaceClient } from "./client.js";

test("node MarketplaceClient normalizes baseUrl with /api suffix", async () => {
  const originalFetch = globalThis.fetch;
  /** @type {string | null} */
  let requested = null;

  try {
    // eslint-disable-next-line no-global-assign
    globalThis.fetch = async (url) => {
      requested = String(url);
      return {
        ok: true,
        status: 200,
        json: async () => ({ results: [] }),
      };
    };

    const client = new MarketplaceClient({ baseUrl: "https://example.com/api/" });
    await client.search({ q: "hello" });

    assert.ok(requested);
    const parsed = new URL(requested);
    assert.equal(parsed.origin, "https://example.com");
    assert.equal(parsed.pathname, "/api/search");
    assert.equal(parsed.searchParams.get("q"), "hello");
  } finally {
    // eslint-disable-next-line no-global-assign
    globalThis.fetch = originalFetch;
  }
});

test("node MarketplaceClient strips query/hash from baseUrl", async () => {
  const originalFetch = globalThis.fetch;
  /** @type {string | null} */
  let requested = null;

  try {
    // eslint-disable-next-line no-global-assign
    globalThis.fetch = async (url) => {
      requested = String(url);
      return {
        ok: true,
        status: 200,
        json: async () => ({ results: [] }),
      };
    };

    const client = new MarketplaceClient({ baseUrl: "https://example.com/api?x=y#frag" });
    await client.search({ q: "hello" });

    assert.ok(requested);
    const parsed = new URL(requested);
    assert.equal(parsed.origin, "https://example.com");
    assert.equal(parsed.pathname, "/api/search");
  } finally {
    // eslint-disable-next-line no-global-assign
    globalThis.fetch = originalFetch;
  }
});

test("node MarketplaceClient strips trailing slash from baseUrl", async () => {
  const originalFetch = globalThis.fetch;
  /** @type {string | null} */
  let requested = null;

  try {
    // eslint-disable-next-line no-global-assign
    globalThis.fetch = async (url) => {
      requested = String(url);
      return {
        ok: true,
        status: 200,
        json: async () => ({ results: [] }),
      };
    };

    const client = new MarketplaceClient({ baseUrl: "https://example.com/" });
    await client.search({ q: "hello" });

    assert.ok(requested);
    const parsed = new URL(requested);
    assert.equal(parsed.origin, "https://example.com");
    assert.equal(parsed.pathname, "/api/search");
  } finally {
    // eslint-disable-next-line no-global-assign
    globalThis.fetch = originalFetch;
  }
});

