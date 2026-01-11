import assert from "node:assert/strict";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";

import { EncryptedFileCredentialStore } from "../../src/credentials/stores/encryptedFile.node.js";
import { isEncryptedFileBytes } from "../../../security/crypto/encryptedFile.js";

import { OAuth2Manager } from "../../src/oauth2/manager.js";
import { CredentialStoreOAuthTokenStore } from "../../src/oauth2/credentialStoreTokenStore.js";
import { normalizeScopes } from "../../src/oauth2/tokenStore.js";

/**
 * @param {any} data
 * @param {number} [status]
 */
function jsonResponse(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { "content-type": "application/json" },
  });
}

test("OAuth2 persistence: refresh token stored via EncryptedFileCredentialStore", async () => {
  const tmpDir = await mkdtemp(path.join(os.tmpdir(), "pq-oauth2-"));
  try {
    const filePath = path.join(tmpDir, "credentials.bin");
    const credentialStore = new EncryptedFileCredentialStore({ filePath, keychainProvider: null });
    const tokenStore = new CredentialStoreOAuthTokenStore(credentialStore);

    const now = () => 1_000;
    let refreshCalls = 0;
    /** @type {typeof fetch} */
    const mockFetch = async (url, init) => {
      if (url !== "https://auth.example/token") throw new Error(`Unexpected URL: ${url}`);
      const body = init?.body;
      const params = body instanceof URLSearchParams ? body : new URLSearchParams(String(body ?? ""));
      assert.equal(params.get("grant_type"), "refresh_token");
      refreshCalls++;
      return jsonResponse({
        access_token: `access-${refreshCalls}`,
        token_type: "Bearer",
        expires_in: 3600,
        refresh_token: "refresh-1",
      });
    };

    const { scopesHash, scopes } = normalizeScopes(["read"]);
    await tokenStore.set(
      { providerId: "example", scopesHash },
      { providerId: "example", scopesHash, scopes, refreshToken: "refresh-1" },
    );

    const manager1 = new OAuth2Manager({ tokenStore, fetch: mockFetch, now });
    manager1.registerProvider({ id: "example", clientId: "client", tokenEndpoint: "https://auth.example/token" });

    const token1 = await manager1.getAccessToken({ providerId: "example", scopes: ["read"] });
    assert.equal(token1.accessToken, "access-1");

    const fileBytes = await readFile(filePath);
    assert.ok(isEncryptedFileBytes(fileBytes), "credential store file should be encrypted");

    const credentialStore2 = new EncryptedFileCredentialStore({ filePath, keychainProvider: null });
    const tokenStore2 = new CredentialStoreOAuthTokenStore(credentialStore2);
    const manager2 = new OAuth2Manager({ tokenStore: tokenStore2, fetch: mockFetch, now });
    manager2.registerProvider({ id: "example", clientId: "client", tokenEndpoint: "https://auth.example/token" });

    const token2 = await manager2.getAccessToken({ providerId: "example", scopes: ["read"] });
    assert.equal(token2.accessToken, "access-2");
    assert.equal(refreshCalls, 2);
  } finally {
    await rm(tmpDir, { recursive: true, force: true });
  }
});

