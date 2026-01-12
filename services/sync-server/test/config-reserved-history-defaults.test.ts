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

test("reserved history quotas default to non-zero in production", () => {
  const config = withEnv(
    {
      NODE_ENV: "production",
      // Ensure auth is configured so config parsing succeeds.
      SYNC_SERVER_AUTH_TOKEN: "token",
      SYNC_SERVER_JWT_SECRET: "",
      // Unset quotas so defaults apply.
      SYNC_SERVER_MAX_VERSIONS_PER_DOC: undefined,
      SYNC_SERVER_MAX_BRANCHING_COMMITS_PER_DOC: undefined,
    },
    () => loadConfigFromEnv()
  );

  assert.equal(config.limits.maxVersionsPerDoc, 500);
  assert.equal(config.limits.maxBranchingCommitsPerDoc, 5_000);
});

test("reserved history quotas default to 0 outside production", () => {
  const config = withEnv(
    {
      NODE_ENV: "test",
      SYNC_SERVER_AUTH_TOKEN: "token",
      SYNC_SERVER_JWT_SECRET: "",
      SYNC_SERVER_MAX_VERSIONS_PER_DOC: undefined,
      SYNC_SERVER_MAX_BRANCHING_COMMITS_PER_DOC: undefined,
    },
    () => loadConfigFromEnv()
  );

  assert.equal(config.limits.maxVersionsPerDoc, 0);
  assert.equal(config.limits.maxBranchingCommitsPerDoc, 0);
});

test("reserved history quota env vars override defaults", () => {
  const config = withEnv(
    {
      NODE_ENV: "production",
      SYNC_SERVER_AUTH_TOKEN: "token",
      SYNC_SERVER_JWT_SECRET: "",
      SYNC_SERVER_MAX_VERSIONS_PER_DOC: "123",
      SYNC_SERVER_MAX_BRANCHING_COMMITS_PER_DOC: "456",
    },
    () => loadConfigFromEnv()
  );

  assert.equal(config.limits.maxVersionsPerDoc, 123);
  assert.equal(config.limits.maxBranchingCommitsPerDoc, 456);

  const disabled = withEnv(
    {
      NODE_ENV: "production",
      SYNC_SERVER_AUTH_TOKEN: "token",
      SYNC_SERVER_JWT_SECRET: "",
      SYNC_SERVER_MAX_VERSIONS_PER_DOC: "0",
      SYNC_SERVER_MAX_BRANCHING_COMMITS_PER_DOC: "0",
    },
    () => loadConfigFromEnv()
  );

  assert.equal(disabled.limits.maxVersionsPerDoc, 0);
  assert.equal(disabled.limits.maxBranchingCommitsPerDoc, 0);
});

