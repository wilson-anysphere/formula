import assert from "node:assert/strict";
import test from "node:test";

import { InMemoryKeychainProvider } from "../../../security/crypto/keychain/inMemoryKeychain.js";
import { CredentialManager } from "../../src/credentials/manager.js";
import { KeychainCredentialStore } from "../../src/credentials/stores/keychain.js";

test("CredentialManager: resolves scopes and persists secrets via KeychainCredentialStore", async () => {
  const keychain = new InMemoryKeychainProvider();
  const store = new KeychainCredentialStore({ keychainProvider: keychain, service: "pq-test" });

  let promptCalls = 0;
  const prompt = async ({ connectorId, scope }) => {
    promptCalls += 1;
    if (connectorId === "http") {
      assert.deepEqual(scope, { type: "http", origin: "https://api.example.com" });
      return { headers: { Authorization: "Bearer token-1" } };
    }
    if (connectorId === "file") {
      assert.deepEqual(scope, { type: "file", match: "exact", path: "/tmp/data.csv" });
      return { access: "granted" };
    }
    if (connectorId === "sql") {
      assert.deepEqual(scope, { type: "sql", server: "db.example.com:5432", database: "analytics", user: "alice" });
      return { password: "pw-1" };
    }
    return null;
  };

  const manager = new CredentialManager({ store, prompt });

  const httpHandle1 = await manager.onCredentialRequest("http", { request: { url: "https://api.example.com/v1/data" } });
  assert.ok(httpHandle1);
  assert.equal(typeof httpHandle1.credentialId, "string");
  assert.equal(httpHandle1.credentialId, httpHandle1.id);
  assert.deepEqual(await httpHandle1.getSecret(), { headers: { Authorization: "Bearer token-1" } });

  // Same origin => same scope => should not prompt again and should return the same credential id.
  const httpHandle2 = await manager.onCredentialRequest("http", { request: { url: "https://api.example.com/v2/other" } });
  assert.ok(httpHandle2);
  assert.equal(httpHandle2.credentialId, httpHandle1.credentialId);
  assert.equal(promptCalls, 1);

  // File scope uses exact path matching by default.
  const fileHandle = await manager.onCredentialRequest("file", { request: { format: "csv", path: "/tmp/data.csv" } });
  assert.ok(fileHandle);
  assert.deepEqual(await fileHandle.getSecret(), { access: "granted" });

  // SQL scope supports URL-style connection strings.
  const sqlHandle = await manager.onCredentialRequest("sql", {
    request: { connection: "postgres://alice:pw-ignored@db.example.com:5432/analytics", sql: "select 1" },
  });
  assert.ok(sqlHandle);
  assert.deepEqual(await sqlHandle.getSecret(), { password: "pw-1" });

  // New manager instance should still read credentials from the shared keychain provider without prompting.
  let promptCalledAgain = false;
  const manager2 = new CredentialManager({
    store: new KeychainCredentialStore({ keychainProvider: keychain, service: "pq-test" }),
    prompt: async () => {
      promptCalledAgain = true;
      return null;
    },
  });
  const httpHandle3 = await manager2.onCredentialRequest("http", { request: { url: "https://api.example.com/v3" } });
  assert.ok(httpHandle3);
  assert.equal(httpHandle3.credentialId, httpHandle1.credentialId);
  assert.equal(promptCalledAgain, false);
});

