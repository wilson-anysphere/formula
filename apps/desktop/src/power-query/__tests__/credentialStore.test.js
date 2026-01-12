import assert from "node:assert/strict";
import test from "node:test";

import { CacheManager, MemoryCacheStore, QueryEngine, httpScope } from "@formula/power-query";

import { createPowerQueryCredentialManager } from "../credentialManager.ts";
import { createDesktopOAuth2Manager } from "../oauth2Manager.ts";

function createMockCredentialInvoke() {
  /** @type {Map<string, { id: string; secret: any }>} */
  const entries = new Map();
  let idCounter = 0;

  return {
    invoke: async (cmd, args) => {
      if (cmd === "power_query_credential_get") {
        const scopeKey = args?.scope_key;
        return scopeKey ? entries.get(String(scopeKey)) ?? null : null;
      }
      if (cmd === "power_query_credential_set") {
        const scopeKey = String(args?.scope_key ?? "");
        if (!scopeKey) throw new Error("missing scope_key");
        const entry = { id: `id-${++idCounter}`, secret: args?.secret };
        entries.set(scopeKey, entry);
        return entry;
      }
      if (cmd === "power_query_credential_delete") {
        const scopeKey = String(args?.scope_key ?? "");
        if (scopeKey) entries.delete(scopeKey);
        return null;
      }
      if (cmd === "power_query_credential_list") {
        return Array.from(entries.entries()).map(([scopeKey, entry]) => ({ scopeKey, id: entry.id }));
      }
      throw new Error(`Unexpected invoke: ${cmd}`);
    },
  };
}

test("Power Query desktop credential store persists across manager instances (Tauri)", async () => {
  const originalTauri = globalThis.__TAURI__;
  const backend = createMockCredentialInvoke();
  globalThis.__TAURI__ = { core: { invoke: backend.invoke } };

  try {
    const scope = httpScope({ url: "https://example.com/api" });
    const secret = { username: "user", password: "supersecret" };

    const mgr1 = createPowerQueryCredentialManager();
    const created = await mgr1.store.set(scope, secret);
    assert.equal(created.secret.password, "supersecret");
    assert.ok(typeof created.id === "string" && created.id.length > 0);
    assert.ok(!created.id.includes("supersecret"));

    const mgr2 = createPowerQueryCredentialManager();
    const loaded = await mgr2.store.get(scope);
    assert.deepEqual(loaded, created);

    const updated = await mgr2.store.set(scope, { username: "user", password: "supersecret2" });
    assert.notEqual(updated.id, created.id, "expected id to change when the secret changes (cache invalidation)");
  } finally {
    globalThis.__TAURI__ = originalTauri;
  }
});

test("QueryEngine cache keys vary by credential id and do not call getSecret()", async () => {
  let getSecretCalls = 0;
  let credentialId = "cred-1";

  const engine = new QueryEngine({
    cache: new CacheManager({ store: new MemoryCacheStore() }),
    onCredentialRequest: async () => ({
      id: credentialId,
      getSecret: async () => {
        getSecretCalls += 1;
        return "supersecret";
      },
    }),
  });

  const query = {
    id: "q_api",
    name: "API",
    source: { type: "api", url: "https://example.com/api", method: "GET", headers: {}, auth: null },
    steps: [],
  };

  const key1 = await engine.getCacheKey(query, {}, {});
  assert.ok(typeof key1 === "string" && key1.length > 0);
  assert.ok(!key1.includes("supersecret"));
  assert.equal(getSecretCalls, 0, "expected cache key computation to avoid retrieving secret material");

  credentialId = "cred-2";
  const key2 = await engine.getCacheKey(query, {}, {});
  assert.notEqual(key1, key2, "expected cache key to vary by credential id");
});

test("OAuth2 refresh tokens persist via CredentialStoreOAuthTokenStore (desktop)", async () => {
  const originalTauri = globalThis.__TAURI__;
  const backend = createMockCredentialInvoke();
  globalThis.__TAURI__ = { core: { invoke: backend.invoke } };

  try {
    const provider = {
      id: "provider",
      clientId: "client",
      tokenEndpoint: "https://example.com/oauth/token",
      defaultScopes: ["scope-a"],
    };

    const { oauth2: oauth2a } = createDesktopOAuth2Manager();
    oauth2a.registerProvider(provider);
    const storeKey = oauth2a.makeStoreKey(provider.id, provider.defaultScopes);
    await oauth2a.persistTokens(storeKey, {
      providerId: provider.id,
      scopesHash: storeKey.scopesHash,
      scopes: provider.defaultScopes,
      refreshToken: "rt-1",
      accessToken: undefined,
      expiresAtMs: undefined,
    });

    let fetchCalls = 0;
    /** @type {typeof fetch} */
    const fetch = async (url, options) => {
      fetchCalls += 1;
      assert.equal(url, provider.tokenEndpoint);
      assert.equal(options?.method, "POST");
      const body = /** @type {any} */ (options?.body);
      // Token client uses URLSearchParams.
      const params = body instanceof URLSearchParams ? body : new URLSearchParams(String(body ?? ""));
      assert.equal(params.get("grant_type"), "refresh_token");
      assert.equal(params.get("refresh_token"), "rt-1");
      return new Response(JSON.stringify({ access_token: "at-1", token_type: "bearer", expires_in: 3600, refresh_token: "rt-2" }), {
        status: 200,
        headers: { "content-type": "application/json" },
      });
    };

    const { oauth2: oauth2b } = createDesktopOAuth2Manager({ fetch });
    oauth2b.registerProvider(provider);
    const refreshed = await oauth2b.getAccessToken({ providerId: provider.id, scopes: provider.defaultScopes });
    assert.equal(refreshed.accessToken, "at-1");
    assert.equal(fetchCalls, 1);

    const { oauth2: oauth2c } = createDesktopOAuth2Manager();
    oauth2c.registerProvider(provider);
    const reloaded = await oauth2c.tokenStore.get(storeKey);
    assert.equal(reloaded?.refreshToken, "rt-2");
  } finally {
    globalThis.__TAURI__ = originalTauri;
  }
});
