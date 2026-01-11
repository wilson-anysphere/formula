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

test("SYNC_SERVER_AUTH_MODE=introspect requires SYNC_SERVER_INTROSPECT_URL", () => {
  assert.throws(
    () =>
      withEnv(
        {
          NODE_ENV: "test",
          SYNC_SERVER_AUTH_MODE: "introspect",
          SYNC_SERVER_INTROSPECT_URL: "",
          SYNC_SERVER_INTROSPECT_TOKEN: "admin-token",
          // Ensure other auth env vars don't mask the failure.
          SYNC_SERVER_AUTH_TOKEN: "",
          SYNC_SERVER_JWT_SECRET: "",
        },
        () => loadConfigFromEnv()
      ),
    /SYNC_SERVER_INTROSPECT_URL/
  );
});

test("SYNC_SERVER_AUTH_MODE=introspect requires SYNC_SERVER_INTROSPECT_TOKEN", () => {
  assert.throws(
    () =>
      withEnv(
        {
          NODE_ENV: "test",
          SYNC_SERVER_AUTH_MODE: "introspect",
          SYNC_SERVER_INTROSPECT_URL: "http://127.0.0.1:1234",
          SYNC_SERVER_INTROSPECT_TOKEN: "",
          SYNC_SERVER_AUTH_TOKEN: "",
          SYNC_SERVER_JWT_SECRET: "",
        },
        () => loadConfigFromEnv()
      ),
    /SYNC_SERVER_INTROSPECT_TOKEN/
  );
});

test("SYNC_SERVER_INTROSPECT_FAIL_OPEN is disabled in production", () => {
  const config = withEnv(
    {
      NODE_ENV: "production",
      SYNC_SERVER_AUTH_MODE: "introspect",
      SYNC_SERVER_INTROSPECT_URL: "http://127.0.0.1:1234",
      SYNC_SERVER_INTROSPECT_TOKEN: "admin-token",
      SYNC_SERVER_INTROSPECT_CACHE_MS: "1234",
      SYNC_SERVER_INTROSPECT_FAIL_OPEN: "true",
      SYNC_SERVER_AUTH_TOKEN: "",
      SYNC_SERVER_JWT_SECRET: "",
    },
    () => loadConfigFromEnv()
  );

  assert.equal(config.auth.mode, "introspect");
  assert.equal(config.auth.cacheMs, 1234);
  assert.equal(config.auth.failOpen, false);
});

test("SYNC_SERVER_INTROSPECT_FAIL_OPEN is honored in non-production environments", () => {
  const config = withEnv(
    {
      NODE_ENV: "test",
      SYNC_SERVER_AUTH_MODE: "introspect",
      SYNC_SERVER_INTROSPECT_URL: "http://127.0.0.1:1234",
      SYNC_SERVER_INTROSPECT_TOKEN: "admin-token",
      SYNC_SERVER_INTROSPECT_FAIL_OPEN: "true",
      SYNC_SERVER_AUTH_TOKEN: "",
      SYNC_SERVER_JWT_SECRET: "",
    },
    () => loadConfigFromEnv()
  );

  assert.equal(config.auth.mode, "introspect");
  assert.equal(config.auth.failOpen, true);
});

