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

test("SYNC_SERVER_SHUTDOWN_GRACE_MS defaults to 10_000", () => {
  const config = withEnv(
    {
      NODE_ENV: "test",
      SYNC_SERVER_SHUTDOWN_GRACE_MS: undefined,
      // Ensure auth is configured so config parsing succeeds.
      SYNC_SERVER_AUTH_TOKEN: "token",
      SYNC_SERVER_JWT_SECRET: "",
    },
    () => loadConfigFromEnv()
  );
  assert.equal(config.shutdownGraceMs, 10_000);
});

test("SYNC_SERVER_SHUTDOWN_GRACE_MS accepts 0 (immediate termination)", () => {
  const config = withEnv(
    {
      NODE_ENV: "test",
      SYNC_SERVER_SHUTDOWN_GRACE_MS: "0",
      SYNC_SERVER_AUTH_TOKEN: "token",
      SYNC_SERVER_JWT_SECRET: "",
    },
    () => loadConfigFromEnv()
  );
  assert.equal(config.shutdownGraceMs, 0);
});

test("SYNC_SERVER_SHUTDOWN_GRACE_MS clamps negative values to 0", () => {
  const config = withEnv(
    {
      NODE_ENV: "test",
      SYNC_SERVER_SHUTDOWN_GRACE_MS: "-5",
      SYNC_SERVER_AUTH_TOKEN: "token",
      SYNC_SERVER_JWT_SECRET: "",
    },
    () => loadConfigFromEnv()
  );
  assert.equal(config.shutdownGraceMs, 0);
});

test("SYNC_SERVER_SHUTDOWN_GRACE_MS parses positive integers", () => {
  const config = withEnv(
    {
      NODE_ENV: "test",
      SYNC_SERVER_SHUTDOWN_GRACE_MS: "1234",
      SYNC_SERVER_AUTH_TOKEN: "token",
      SYNC_SERVER_JWT_SECRET: "",
    },
    () => loadConfigFromEnv()
  );
  assert.equal(config.shutdownGraceMs, 1234);
});

