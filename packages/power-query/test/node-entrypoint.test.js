import assert from "node:assert/strict";
import test from "node:test";

import * as powerQuery from "../src/node.js";

test("power-query node entrypoint exports Node-only helpers", () => {
  assert.equal(typeof powerQuery.FileSystemCacheStore, "function");
  assert.equal(typeof powerQuery.EncryptedFileSystemCacheStore, "function");
  assert.equal(typeof powerQuery.createNodeCryptoCacheProvider, "function");
  assert.equal(typeof powerQuery.createNodeCredentialStore, "function");
});
