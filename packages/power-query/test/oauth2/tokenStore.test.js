import assert from "node:assert/strict";
import test from "node:test";

import { normalizeScopes } from "../../src/oauth2/tokenStore.js";

test("normalizeScopes trims, sorts, and dedupes scope strings", () => {
  const a = normalizeScopes([" read", "write", "read", "", " write ", "  "]);
  assert.deepEqual(a.scopes, ["read", "write"]);

  const b = normalizeScopes(["write", "read"]);
  assert.equal(a.scopesHash, b.scopesHash);
});

