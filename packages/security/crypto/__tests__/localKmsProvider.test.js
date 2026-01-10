import test from "node:test";
import assert from "node:assert/strict";

import { LocalKmsProvider } from "../kms/localKmsProvider.js";

test("LocalKmsProvider wraps and unwraps keys", () => {
  const kms = new LocalKmsProvider();
  const dek = Buffer.from("12345678901234567890123456789012", "utf8"); // 32 bytes

  const context = { orgId: "org-1", purpose: "document" };
  const wrapped = kms.wrapKey({ plaintextKey: dek, encryptionContext: context });
  const unwrapped = kms.unwrapKey({ wrappedKey: wrapped, encryptionContext: context });

  assert.deepEqual(unwrapped, dek);
});

test("LocalKmsProvider unwrap fails when encryption context mismatches", () => {
  const kms = new LocalKmsProvider();
  const dek = Buffer.from("12345678901234567890123456789012", "utf8");

  const wrapped = kms.wrapKey({ plaintextKey: dek, encryptionContext: { v: 1 } });
  assert.throws(() => {
    kms.unwrapKey({ wrappedKey: wrapped, encryptionContext: { v: 2 } });
  });
});

test("LocalKmsProvider key rotation preserves ability to unwrap old keys", () => {
  const kms = new LocalKmsProvider();
  const dek1 = Buffer.from("aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa", "utf8");
  const ctx = { docId: "doc-1" };

  const wrapped1 = kms.wrapKey({ plaintextKey: dek1, encryptionContext: ctx });
  kms.rotateKey();

  const dek2 = Buffer.from("bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb", "utf8");
  const wrapped2 = kms.wrapKey({ plaintextKey: dek2, encryptionContext: ctx });

  assert.deepEqual(kms.unwrapKey({ wrappedKey: wrapped1, encryptionContext: ctx }), dek1);
  assert.deepEqual(kms.unwrapKey({ wrappedKey: wrapped2, encryptionContext: ctx }), dek2);
});

