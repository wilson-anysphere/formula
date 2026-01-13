import { describe, expect, it } from "vitest";

import {
  resolveCollabSessionPermissionsFromToken,
  tryDecodeJwtPayload,
  tryDeriveCollabSessionPermissionsFromJwtToken,
} from "../jwt";

function encodeBase64Url(value: string): string {
  return Buffer.from(value, "utf8")
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/g, "");
}

function makeJwt(payload: unknown): string {
  const header = encodeBase64Url(JSON.stringify({ alg: "none", typ: "JWT" }));
  const body = encodeBase64Url(JSON.stringify(payload));
  return `${header}.${body}.sig`;
}

describe("collab/jwt", () => {
  describe("tryDecodeJwtPayload", () => {
    it("returns null for tokens that do not look like header.payload.signature", () => {
      expect(tryDecodeJwtPayload("")).toBeNull();
      expect(tryDecodeJwtPayload("not-a-jwt")).toBeNull();
      expect(tryDecodeJwtPayload("a.b")).toBeNull();
      expect(tryDecodeJwtPayload("a.b.c.d")).toBeNull();
      expect(tryDecodeJwtPayload("a..c")).toBeNull();
    });

    it("returns null for malformed base64url payloads", () => {
      expect(tryDecodeJwtPayload("a.%ZZ.c")).toBeNull();
      expect(tryDecodeJwtPayload("a.!!!.c")).toBeNull();
    });

    it("returns null when the decoded payload is not valid JSON", () => {
      const token = `${encodeBase64Url("header")}.${encodeBase64Url("not-json")}.sig`;
      expect(tryDecodeJwtPayload(token)).toBeNull();
    });

    it("supports unicode payloads", () => {
      const token = makeJwt({ sub: "user-1", name: "José", role: "viewer" });
      expect(tryDecodeJwtPayload(token)).toEqual({ sub: "user-1", name: "José", role: "viewer" });
    });
  });

  it("derives CollabSession permissions from JWT claims (best-effort)", () => {
    const token = makeJwt({
      sub: "user-123",
      role: "viewer",
      rangeRestrictions: [{ sheetId: "Sheet1", startRow: 0, endRow: 1, startCol: 0, endCol: 1, role: "viewer" }],
    });

    const perms = resolveCollabSessionPermissionsFromToken({ token, fallbackUserId: "fallback" });
    expect(perms).toEqual({
      role: "viewer",
      rangeRestrictions: [{ sheetId: "Sheet1", startRow: 0, endRow: 1, startCol: 0, endCol: 1, role: "viewer" }],
      userId: "user-123",
    });

    // Also assert the "raw" extractor returns the same claims.
    expect(tryDeriveCollabSessionPermissionsFromJwtToken(token)).toEqual({
      userId: "user-123",
      role: "viewer",
      rangeRestrictions: [{ sheetId: "Sheet1", startRow: 0, endRow: 1, startCol: 0, endCol: 1, role: "viewer" }],
    });
  });
});

