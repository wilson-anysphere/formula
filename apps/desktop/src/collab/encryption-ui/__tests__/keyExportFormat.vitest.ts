import { describe, expect, test } from "vitest";

import { parseEncryptionKeyExportString, serializeEncryptionKeyExportString } from "../keyExportFormat";

describe("encryption key export format", () => {
  test("roundtrips docId/keyId/keyBytes", () => {
    const docId = "doc-123";
    const keyId = "key-abc";
    const keyBytes = new Uint8Array(32);
    for (let i = 0; i < keyBytes.length; i += 1) keyBytes[i] = i;

    const encoded = serializeEncryptionKeyExportString({ docId, keyId, keyBytes });
    const parsed = parseEncryptionKeyExportString(encoded);

    expect(parsed.docId).toBe(docId);
    expect(parsed.keyId).toBe(keyId);
    expect(Array.from(parsed.keyBytes)).toEqual(Array.from(keyBytes));
  });

  test("parses raw base64url token without the formula-enc:// prefix", () => {
    const docId = "doc-123";
    const keyId = "key-abc";
    const keyBytes = new Uint8Array(32);
    keyBytes.fill(7);
    const encoded = serializeEncryptionKeyExportString({ docId, keyId, keyBytes });
    const token = encoded.replace(/^formula-enc:\/\//, "");

    const parsed = parseEncryptionKeyExportString(token);
    expect(parsed.docId).toBe(docId);
    expect(parsed.keyId).toBe(keyId);
    expect(parsed.keyBytes.byteLength).toBe(32);
  });

  test("rejects invalid key lengths", () => {
    expect(() => serializeEncryptionKeyExportString({ docId: "doc", keyId: "k", keyBytes: new Uint8Array(31) })).toThrow(
      /Invalid encryption key length/,
    );

    const docId = "doc-123";
    const keyId = "key-abc";
    const keyBytes = new Uint8Array(32);
    keyBytes.fill(1);
    const encoded = serializeEncryptionKeyExportString({ docId, keyId, keyBytes });
    const token = encoded.replace(/^formula-enc:\/\//, "");
    let b64 = token.replace(/-/g, "+").replace(/_/g, "/");
    while (b64.length % 4 !== 0) b64 += "=";
    const parsed = JSON.parse(Buffer.from(b64, "base64").toString("utf8"));
    parsed.keyBytesBase64 = Buffer.from(new Uint8Array(31)).toString("base64");
    const tampered = `formula-enc://${Buffer.from(JSON.stringify(parsed), "utf8")
      .toString("base64")
      .replace(/\+/g, "-")
      .replace(/\//g, "_")
      .replace(/=+$/g, "")}`;
    expect(() => parseEncryptionKeyExportString(tampered)).toThrow(/Invalid encryption key length/);
  });
});
