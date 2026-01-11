import assert from "node:assert/strict";
import test from "node:test";

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

test("OAuth2Manager: authorizeWithPkce uses broker redirect + exchanges code", async () => {
  const store = new InMemoryOAuthTokenStore();
  const now = () => 1_000;

  let openedUrl = null;

  /** @type {typeof fetch} */
  const mockFetch = async (url, init) => {
    if (url !== "https://auth.example/token") throw new Error(`Unexpected URL: ${url}`);
    const body = init?.body;
    const params = body instanceof URLSearchParams ? body : new URLSearchParams(String(body ?? ""));
    assert.equal(params.get("grant_type"), "authorization_code");
    assert.equal(params.get("code"), "auth-code");
    assert.equal(params.get("client_id"), "client");
    assert.equal(params.get("redirect_uri"), "https://app.example/callback");
    const verifier = params.get("code_verifier");
    assert.ok(typeof verifier === "string" && verifier.length >= 43, "code_verifier should be present (PKCE)");
    return jsonResponse({
      access_token: "access-1",
      token_type: "Bearer",
      expires_in: 3600,
      refresh_token: "refresh-1",
    });
  };

  const manager = new OAuth2Manager({ tokenStore: store, fetch: mockFetch, now });
  manager.registerProvider({
    id: "example",
    clientId: "client",
    tokenEndpoint: "https://auth.example/token",
    authorizationEndpoint: "https://auth.example/authorize",
    redirectUri: "https://app.example/callback",
  });

  const result = await manager.authorizeWithPkce({
    providerId: "example",
    scopes: ["read"],
    broker: {
      openAuthUrl: async (url) => {
        openedUrl = url;
      },
      waitForRedirect: async (redirectUri) => {
        assert.equal(redirectUri, "https://app.example/callback");
        assert.ok(openedUrl, "openAuthUrl should be called before waiting for redirect");
        const parsed = new URL(openedUrl ?? "");
        const state = parsed.searchParams.get("state");
        assert.ok(state, "state param must be present");
        return `${redirectUri}?code=auth-code&state=${encodeURIComponent(state)}`;
      },
    },
  });

  assert.equal(result.accessToken, "access-1");
  const { scopesHash } = normalizeScopes(["read"]);
  const entry = await store.get({ providerId: "example", scopesHash });
  assert.equal(entry?.refreshToken, "refresh-1");
});

test("OAuth2Manager: authorizeWithDeviceCode starts device flow and polls token endpoint", async () => {
  const store = new InMemoryOAuthTokenStore();
  const now = () => 1_000;

  let openedUrl = null;
  let prompted = null;

  /** @type {typeof fetch} */
  const mockFetch = async (url, init) => {
    const body = init?.body;
    const params = body instanceof URLSearchParams ? body : new URLSearchParams(String(body ?? ""));
    if (url === "https://auth.example/device") {
      assert.equal(params.get("client_id"), "client");
      assert.equal(params.get("scope"), "read");
      return jsonResponse({
        device_code: "device-code",
        user_code: "ABCD-EFGH",
        verification_uri: "https://auth.example/verify",
        verification_uri_complete: "https://auth.example/verify?user_code=ABCD-EFGH",
        expires_in: 600,
        interval: 5,
      });
    }
    if (url === "https://auth.example/token") {
      assert.equal(params.get("grant_type"), "urn:ietf:params:oauth:grant-type:device_code");
      assert.equal(params.get("device_code"), "device-code");
      assert.equal(params.get("client_id"), "client");
      return jsonResponse({
        access_token: "access-1",
        token_type: "Bearer",
        expires_in: 3600,
        refresh_token: "refresh-1",
      });
    }
    throw new Error(`Unexpected URL: ${url}`);
  };

  const manager = new OAuth2Manager({ tokenStore: store, fetch: mockFetch, now });
  manager.registerProvider({
    id: "example",
    clientId: "client",
    tokenEndpoint: "https://auth.example/token",
    deviceAuthorizationEndpoint: "https://auth.example/device",
  });

  const result = await manager.authorizeWithDeviceCode({
    providerId: "example",
    scopes: ["read"],
    broker: {
      openAuthUrl: async (url) => {
        openedUrl = url;
      },
      deviceCodePrompt: async (code, verificationUri) => {
        prompted = { code, verificationUri };
      },
    },
  });

  assert.equal(result.accessToken, "access-1");
  assert.equal(openedUrl, "https://auth.example/verify?user_code=ABCD-EFGH");
  assert.deepEqual(prompted, { code: "ABCD-EFGH", verificationUri: "https://auth.example/verify?user_code=ABCD-EFGH" });

  const { scopesHash } = normalizeScopes(["read"]);
  const entry = await store.get({ providerId: "example", scopesHash });
  assert.equal(entry?.refreshToken, "refresh-1");
});

