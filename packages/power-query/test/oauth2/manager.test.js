import assert from "node:assert/strict";
import test from "node:test";

import { HttpConnector } from "../../src/connectors/http.js";
import { OAuth2Manager } from "../../src/oauth2/manager.js";
import { InMemoryOAuthTokenStore, normalizeScopes } from "../../src/oauth2/tokenStore.js";

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

test("OAuth2Manager: exchanges authorization code and refreshes with rotation", async () => {
  const store = new InMemoryOAuthTokenStore();
  const now = () => 1_000;

  let exchangeCalls = 0;
  let refreshCalls = 0;

  /** @type {typeof fetch} */
  const mockFetch = async (url, init) => {
    if (url !== "https://auth.example/token") throw new Error(`Unexpected URL: ${url}`);
    const body = init?.body;
    const params = body instanceof URLSearchParams ? body : new URLSearchParams(String(body ?? ""));
    const grantType = params.get("grant_type");

    if (grantType === "authorization_code") {
      exchangeCalls++;
      assert.equal(params.get("code"), "auth-code");
      assert.equal(params.get("client_id"), "client");
      assert.equal(params.get("redirect_uri"), "https://app.example/callback");
      assert.equal(params.get("code_verifier"), "verifier");
      return jsonResponse({
        access_token: "access-1",
        token_type: "Bearer",
        expires_in: "3600",
        refresh_token: "refresh-1",
      });
    }

    if (grantType === "refresh_token") {
      refreshCalls++;
      assert.equal(params.get("refresh_token"), refreshCalls === 1 ? "refresh-1" : "refresh-2");
      return jsonResponse({
        access_token: refreshCalls === 1 ? "access-2" : "access-3",
        token_type: "Bearer",
        expires_in: "3600",
        refresh_token: refreshCalls === 1 ? "refresh-2" : "refresh-3",
      });
    }

    throw new Error(`Unexpected grant_type: ${grantType}`);
  };

  const manager = new OAuth2Manager({ tokenStore: store, fetch: mockFetch, now });
  manager.registerProvider({
    id: "example",
    clientId: "client",
    tokenEndpoint: "https://auth.example/token",
    redirectUri: "https://app.example/callback",
  });

  const exchanged = await manager.exchangeAuthorizationCode({
    providerId: "example",
    code: "auth-code",
    redirectUri: "https://app.example/callback",
    codeVerifier: "verifier",
    scopes: ["read"],
  });
  assert.equal(exchanged.accessToken, "access-1");
  assert.equal(exchangeCalls, 1);
  assert.equal(refreshCalls, 0);

  const cached = await manager.getAccessToken({ providerId: "example", scopes: ["read"] });
  assert.equal(cached.accessToken, "access-1");
  assert.equal(refreshCalls, 0);

  const refreshed = await manager.getAccessToken({ providerId: "example", scopes: ["read"], forceRefresh: true });
  assert.equal(refreshed.accessToken, "access-2");
  assert.equal(refreshCalls, 1);

  const { scopesHash, scopes } = normalizeScopes(["read"]);
  const entry = await store.get({ providerId: "example", scopesHash });
  assert.equal(entry?.refreshToken, "refresh-2");
  assert.deepEqual(entry?.scopes, scopes);
});

test("OAuth2Manager: dedupes concurrent refresh calls", async () => {
  const store = new InMemoryOAuthTokenStore();
  const now = () => 1_000;

  const { scopesHash, scopes } = normalizeScopes(["read"]);
  await store.set(
    { providerId: "example", scopesHash },
    { providerId: "example", scopesHash, scopes, refreshToken: "refresh-1" },
  );

  let refreshCalls = 0;
  /** @type {typeof fetch} */
  const mockFetch = async (url, init) => {
    if (url !== "https://auth.example/token") throw new Error(`Unexpected URL: ${url}`);
    const body = init?.body;
    const params = body instanceof URLSearchParams ? body : new URLSearchParams(String(body ?? ""));
    assert.equal(params.get("grant_type"), "refresh_token");
    refreshCalls++;
    return jsonResponse({
      access_token: "access-1",
      token_type: "Bearer",
      expires_in: 3600,
      refresh_token: "refresh-1",
    });
  };

  const manager = new OAuth2Manager({ tokenStore: store, fetch: mockFetch, now });
  manager.registerProvider({ id: "example", clientId: "client", tokenEndpoint: "https://auth.example/token" });

  const results = await Promise.all(
    Array.from({ length: 10 }, () => manager.getAccessToken({ providerId: "example", scopes: ["read"] })),
  );

  assert.equal(refreshCalls, 1);
  assert.ok(results.every((r) => r.accessToken === "access-1"));
});

test("HttpConnector: retries once on 401 by forcing an OAuth2 refresh", async () => {
  const store = new InMemoryOAuthTokenStore();
  const now = () => 1_000;

  const { scopesHash, scopes } = normalizeScopes(["read"]);
  await store.set(
    { providerId: "example", scopesHash },
    {
      providerId: "example",
      scopesHash,
      scopes,
      refreshToken: "refresh-1",
      accessToken: "access-1",
      expiresAtMs: now() + 3_600_000,
    },
  );

  let tokenCalls = 0;
  let apiCalls = 0;

  /** @type {typeof fetch} */
  const mockFetch = async (url, init) => {
    if (url === "https://auth.example/token") {
      tokenCalls++;
      const body = init?.body;
      const params = body instanceof URLSearchParams ? body : new URLSearchParams(String(body ?? ""));
      assert.equal(params.get("grant_type"), "refresh_token");
      return jsonResponse({
        access_token: "access-2",
        token_type: "Bearer",
        expires_in: 3600,
        refresh_token: "refresh-1",
      });
    }

    if (url === "https://api.example/data") {
      apiCalls++;
      const auth = /** @type {any} */ (init?.headers)?.Authorization;
      if (auth === "Bearer access-1") {
        return new Response("unauthorized", { status: 401 });
      }
      assert.equal(auth, "Bearer access-2");
      return jsonResponse([{ id: 1 }], 200);
    }

    throw new Error(`Unexpected URL: ${url}`);
  };

  const manager = new OAuth2Manager({ tokenStore: store, fetch: mockFetch, now });
  manager.registerProvider({ id: "example", clientId: "client", tokenEndpoint: "https://auth.example/token" });

  const connector = new HttpConnector({ fetch: mockFetch, oauth2Manager: manager });
  const result = await connector.execute(
    {
      url: "https://api.example/data",
      responseType: "json",
      auth: { type: "oauth2", providerId: "example", scopes: ["read"] },
    },
    { now },
  );

  assert.equal(apiCalls, 2);
  assert.equal(tokenCalls, 1);
  assert.deepEqual(result.table.toGrid(), [["id"], [1]]);
});

test("OAuth2Manager: refresh token survives store reload", async () => {
  const now = () => 1_000;

  const store1 = new InMemoryOAuthTokenStore();
  const { scopesHash, scopes } = normalizeScopes(["read"]);
  await store1.set(
    { providerId: "example", scopesHash },
    { providerId: "example", scopesHash, scopes, refreshToken: "refresh-1" },
  );

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

  const manager1 = new OAuth2Manager({ tokenStore: store1, fetch: mockFetch, now });
  manager1.registerProvider({ id: "example", clientId: "client", tokenEndpoint: "https://auth.example/token" });

  const token1 = await manager1.getAccessToken({ providerId: "example", scopes: ["read"] });
  assert.equal(token1.accessToken, "access-1");

  const store2 = new InMemoryOAuthTokenStore(store1.snapshot());
  const manager2 = new OAuth2Manager({ tokenStore: store2, fetch: mockFetch, now });
  manager2.registerProvider({ id: "example", clientId: "client", tokenEndpoint: "https://auth.example/token" });

  const token2 = await manager2.getAccessToken({ providerId: "example", scopes: ["read"] });
  assert.equal(token2.accessToken, "access-2");
  assert.equal(refreshCalls, 2);
});

test("OAuth2Manager: clears persisted refresh token when the server returns invalid_grant", async () => {
  const store = new InMemoryOAuthTokenStore();
  const now = () => 1_000;
  const { scopesHash, scopes } = normalizeScopes(["read"]);
  await store.set(
    { providerId: "example", scopesHash },
    { providerId: "example", scopesHash, scopes, refreshToken: "bad-refresh" },
  );

  let refreshCalls = 0;
  /** @type {typeof fetch} */
  const mockFetch = async (url, init) => {
    if (url !== "https://auth.example/token") throw new Error(`Unexpected URL: ${url}`);
    const body = init?.body;
    const params = body instanceof URLSearchParams ? body : new URLSearchParams(String(body ?? ""));
    assert.equal(params.get("grant_type"), "refresh_token");
    refreshCalls++;
    return jsonResponse({ error: "invalid_grant", error_description: "revoked" }, 400);
  };

  const manager = new OAuth2Manager({ tokenStore: store, fetch: mockFetch, now });
  manager.registerProvider({ id: "example", clientId: "client", tokenEndpoint: "https://auth.example/token" });

  await assert.rejects(
    () => manager.getAccessToken({ providerId: "example", scopes: ["read"] }),
    /Re-authentication is required/,
  );
  assert.equal(refreshCalls, 1);

  const entry = await store.get({ providerId: "example", scopesHash });
  assert.equal(entry, null);
});
