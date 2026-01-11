import assert from "node:assert/strict";
import test from "node:test";

import * as powerQuery from "../src/node.js";

test("power-query node entrypoint exports Node-only helpers", () => {
  assert.equal(typeof powerQuery.FileSystemCacheStore, "function");
  assert.equal(typeof powerQuery.EncryptedFileSystemCacheStore, "function");
  assert.equal(typeof powerQuery.createNodeCryptoCacheProvider, "function");
  assert.equal(typeof powerQuery.createNodeCredentialStore, "function");
});

test("power-query node entrypoint exports core scalar wrapper types", () => {
  assert.equal(powerQuery.MS_PER_DAY, 24 * 60 * 60 * 1000);
  assert.equal(typeof powerQuery.PqDecimal, "function");
  assert.equal(typeof powerQuery.PqTime, "function");
  assert.equal(typeof powerQuery.PqDuration, "function");
  assert.equal(typeof powerQuery.PqDateTimeZone, "function");
});
