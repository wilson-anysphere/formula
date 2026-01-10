import test from "node:test";
import assert from "node:assert/strict";

import { decryptAes256Gcm, encryptAes256Gcm, generateAes256Key } from "../aes256gcm.js";

test("aes-256-gcm round-trip", () => {
  const key = generateAes256Key();
  const plaintext = Buffer.from("top secret", "utf8");
  const aad = Buffer.from("context", "utf8");

  const encrypted = encryptAes256Gcm({ plaintext, key, aad });
  const decrypted = decryptAes256Gcm({ ...encrypted, key, aad });

  assert.equal(decrypted.toString("utf8"), "top secret");
});

test("aes-256-gcm rejects tampered ciphertext", () => {
  const key = generateAes256Key();
  const plaintext = Buffer.from("top secret", "utf8");

  const encrypted = encryptAes256Gcm({ plaintext, key });
  const tampered = Buffer.from(encrypted.ciphertext);
  tampered[0] ^= 0xff;

  assert.throws(() => {
    decryptAes256Gcm({ ...encrypted, ciphertext: tampered, key });
  });
});

test("aes-256-gcm rejects wrong AAD", () => {
  const key = generateAes256Key();
  const plaintext = Buffer.from("top secret", "utf8");
  const encrypted = encryptAes256Gcm({
    plaintext,
    key,
    aad: Buffer.from("aad:one", "utf8")
  });

  assert.throws(() => {
    decryptAes256Gcm({
      ...encrypted,
      key,
      aad: Buffer.from("aad:two", "utf8")
    });
  });
});

