import assert from "node:assert/strict";
import test from "node:test";

import { QueryEngine } from "../../src/engine.js";
import { HttpConnector } from "../../src/connectors/http.js";

import { CredentialManager } from "../../src/credentials/manager.js";
import { InMemoryCredentialStore } from "../../src/credentials/stores/inMemory.js";
import { httpScope } from "../../src/credentials/scopes.js";

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

test("CredentialManager + HttpConnector OAuth2: stored credential handle can provide oauth2 config", async () => {
  const now = () => 1_000;

  const credentialStore = new InMemoryCredentialStore();
  const cm = new CredentialManager({ store: credentialStore });

  const url = "https://api.example/data";
  const scope = httpScope({ url });
  await credentialStore.set(scope, { oauth2: { providerId: "example", scopes: ["read"] } });

  const tokenStore = new InMemoryOAuthTokenStore();
  const { scopesHash, scopes } = normalizeScopes(["read"]);
  await tokenStore.set(
    { providerId: "example", scopesHash },
    { providerId: "example", scopesHash, scopes, refreshToken: "refresh-1" },
  );

  let tokenCalls = 0;
  let apiCalls = 0;

  /** @type {typeof fetch} */
  const fetchFn = async (requestUrl, init) => {
    if (requestUrl === "https://auth.example/token") {
      tokenCalls++;
      const body = init?.body;
      const params = body instanceof URLSearchParams ? body : new URLSearchParams(String(body ?? ""));
      assert.equal(params.get("grant_type"), "refresh_token");
      return jsonResponse({
        access_token: "access-1",
        token_type: "Bearer",
        expires_in: 3600,
        refresh_token: "refresh-1",
      });
    }
    if (requestUrl === url) {
      apiCalls++;
      const auth = /** @type {any} */ (init?.headers)?.Authorization;
      assert.equal(auth, "Bearer access-1");
      return jsonResponse([{ id: 1 }]);
    }
    throw new Error(`Unexpected URL: ${requestUrl}`);
  };

  const oauth2Manager = new OAuth2Manager({ tokenStore, fetch: fetchFn, now });
  oauth2Manager.registerProvider({ id: "example", clientId: "client", tokenEndpoint: "https://auth.example/token" });

  const http = new HttpConnector({ fetch: fetchFn, oauth2Manager });
  const engine = new QueryEngine({
    connectors: { http },
    onCredentialRequest: cm.onCredentialRequest.bind(cm),
  });

  const query = {
    id: "q1",
    name: "Query 1",
    source: { type: "api", url, method: "GET" },
    steps: [],
  };

  const table = await engine.executeQuery(query, {}, { now });
  assert.deepEqual(table.toGrid(), [["id"], [1]]);
  assert.equal(tokenCalls, 1);
  assert.equal(apiCalls, 1);
});

