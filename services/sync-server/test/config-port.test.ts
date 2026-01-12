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

test("SYNC_SERVER_PORT accepts port 0", () => {
  const config = withEnv(
    {
      NODE_ENV: "test",
      SYNC_SERVER_PORT: "0",
      // Ensure auth is configured so config parsing succeeds.
      SYNC_SERVER_AUTH_TOKEN: "token",
      SYNC_SERVER_JWT_SECRET: "",
    },
    () => loadConfigFromEnv()
  );
  assert.equal(config.port, 0);
});

test("SYNC_SERVER_PORT rejects negative ports", () => {
  assert.throws(() => {
    withEnv(
      {
        NODE_ENV: "test",
        SYNC_SERVER_PORT: "-1",
        SYNC_SERVER_AUTH_TOKEN: "token",
        SYNC_SERVER_JWT_SECRET: "",
      },
      () => loadConfigFromEnv()
    );
  });
});

test("SYNC_SERVER_PORT rejects ports > 65535", () => {
  assert.throws(() => {
    withEnv(
      {
        NODE_ENV: "test",
        SYNC_SERVER_PORT: "65536",
        SYNC_SERVER_AUTH_TOKEN: "token",
        SYNC_SERVER_JWT_SECRET: "",
      },
      () => loadConfigFromEnv()
    );
  });
});

