import assert from "node:assert/strict";
import test from "node:test";

import { QueryEngine } from "../../src/engine.js";
import { HttpConnector } from "../../src/connectors/http.js";

test("Http OAuth2: request.auth overrides credentials-provided oauth2 config", async () => {
  /** @type {any[]} */
  const observedTokenRequests = [];

  const oauth2Manager = {
    getAccessToken: async (opts) => {
      observedTokenRequests.push(opts);
      return { accessToken: `token-for-${opts.providerId}`, expiresAtMs: null, refreshToken: null };
    },
  };

  /** @type {typeof fetch} */
  const fetchFn = async (_url, init) => {
    const auth = /** @type {any} */ (init?.headers)?.Authorization;
    assert.equal(auth, "Bearer token-for-explicit");
    return new Response(JSON.stringify([{ id: 1 }]), { status: 200, headers: { "content-type": "application/json" } });
  };

  const http = new HttpConnector({ fetch: fetchFn, oauth2Manager });
  const engine = new QueryEngine({
    connectors: { http },
    onCredentialRequest: async () => ({ oauth2: { providerId: "from-credentials" } }),
  });

  const query = {
    id: "q1",
    name: "Query 1",
    source: {
      type: "api",
      url: "https://api.example/data",
      method: "GET",
      auth: { type: "oauth2", providerId: "explicit" },
    },
    steps: [],
    refreshPolicy: { type: "manual" },
  };

  const table = await engine.executeQuery(query, {}, { limit: 10 });
  assert.deepEqual(table.toGrid(), [["id"], [1]]);
  assert.equal(observedTokenRequests.length, 1);
  assert.equal(observedTokenRequests[0].providerId, "explicit");
});

