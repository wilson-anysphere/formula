import assert from "node:assert/strict";
import test from "node:test";

import jwt from "jsonwebtoken";

import { AuthError, authenticateRequest } from "../src/auth.js";
import type { AuthMode } from "../src/config.js";

const JWT_SECRET = "test-secret";
const JWT_AUDIENCE = "formula-sync";

function signJwt(payload: Record<string, unknown>): string {
  return jwt.sign(payload, JWT_SECRET, {
    algorithm: "HS256",
    audience: JWT_AUDIENCE,
  });
}

const auth: AuthMode = {
  mode: "jwt-hs256",
  secret: JWT_SECRET,
  audience: JWT_AUDIENCE,
  requireSub: false,
  requireExp: false,
};

test("authenticateRequest accepts JWT rangeRestrictions claim", async () => {
  const docId = "doc-1";
  const token = signJwt({
    sub: "u1",
    docId,
    role: "editor",
    rangeRestrictions: [
      {
        range: {
          sheetId: "Sheet1",
          startRow: 0,
          startCol: 0,
          endRow: 0,
          endCol: 0,
        },
        editAllowlist: ["u1"],
      },
    ],
  });

  const ctx = await authenticateRequest(auth, token, docId);
  assert.equal(ctx.userId, "u1");
  assert.equal(ctx.docId, docId);
  assert.equal(ctx.role, "editor");
  assert.ok(Array.isArray(ctx.rangeRestrictions));
  assert.equal(ctx.rangeRestrictions.length, 1);
});

test("authenticateRequest accepts rangeRestrictions sheetName alias", async () => {
  const docId = "doc-1b";
  const token = signJwt({
    sub: "u1",
    docId,
    role: "editor",
    rangeRestrictions: [
      {
        sheetName: "Sheet1",
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 0,
        editAllowlist: ["u1"],
      },
    ],
  });

  const ctx = await authenticateRequest(auth, token, docId);
  assert.ok(Array.isArray(ctx.rangeRestrictions));
  assert.equal(ctx.rangeRestrictions.length, 1);
});

test("authenticateRequest rejects rangeRestrictions when it is not an array", async () => {
  const docId = "doc-2";
  const token = signJwt({
    sub: "u1",
    docId,
    role: "editor",
    rangeRestrictions: { not: "an-array" },
  });

  await assert.rejects(
    authenticateRequest(auth, token, docId),
    (err) => err instanceof AuthError && err.statusCode === 403
  );
});

test("authenticateRequest rejects invalid rangeRestrictions entries", async () => {
  const docId = "doc-3";
  const token = signJwt({
    sub: "u1",
    docId,
    role: "editor",
    rangeRestrictions: [
      {
        sheetId: "Sheet1",
        startRow: -1,
        startCol: 0,
        endRow: 0,
        endCol: 0,
      },
    ],
  });

  await assert.rejects(
    authenticateRequest(auth, token, docId),
    (err) => err instanceof AuthError && err.statusCode === 403
  );
});
