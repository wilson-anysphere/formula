import assert from "node:assert/strict";
import test from "node:test";

import { loadConfigFromEnv } from "../src/config.js";

function withEnv<T>(overrides: Record<string, string | undefined>, fn: () => T): T {
  const previous: Record<string, string | undefined> = {};
  for (const key of Object.keys(overrides)) {
    previous[key] = process.env[key];
    const next = overrides[key];
    if (next === undefined) {
      delete process.env[key];
    } else {
      process.env[key] = next;
    }
  }

  try {
    return fn();
  } finally {
    for (const key of Object.keys(overrides)) {
      const prev = previous[key];
      if (prev === undefined) {
        delete process.env[key];
      } else {
        process.env[key] = prev;
      }
    }
  }
}

test("SYNC_SERVER_TRUST_PROXY defaults to false", () => {
  const config = withEnv(
    {
      NODE_ENV: "test",
      SYNC_SERVER_TRUST_PROXY: undefined,
      // Ensure auth is configured so config parsing succeeds.
      SYNC_SERVER_AUTH_TOKEN: "token",
      SYNC_SERVER_JWT_SECRET: "",
    },
    () => loadConfigFromEnv()
  );
  assert.equal(config.trustProxy, false);
});

test("SYNC_SERVER_TRUST_PROXY=true enables trusting x-forwarded-for", () => {
  const config = withEnv(
    {
      NODE_ENV: "test",
      SYNC_SERVER_TRUST_PROXY: "true",
      SYNC_SERVER_AUTH_TOKEN: "token",
      SYNC_SERVER_JWT_SECRET: "",
    },
    () => loadConfigFromEnv()
  );
  assert.equal(config.trustProxy, true, "expected trustProxy=true when SYNC_SERVER_TRUST_PROXY is true");
});

