import assert from "node:assert/strict";
import test from "node:test";

import { createCodeChallenge } from "../../src/oauth2/pkce.js";

test("PKCE: code challenge matches RFC 7636 test vector", async () => {
  // RFC 7636 Appendix B
  const verifier = "dBjftJeZ4CVP-mB92K27uhbUJU1p1r_wW1gFWFOEjXk";
  const expected = "E9Melhoa2OwvFrEMTJguCHaoeK1t8URWbuGJSstw-cM";

  const challenge = await createCodeChallenge(verifier);
  assert.equal(challenge, expected);
});

