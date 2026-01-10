import test from "node:test";
import assert from "node:assert/strict";

import { KeyRing } from "../keyring.js";
import { decodeEncryptedFileBytes, encodeEncryptedFileBytes, isEncryptedFileBytes } from "../encryptedFile.js";

test("encrypted file format round-trips with KeyRing.encryptBytes/decryptBytes", () => {
  const ring = KeyRing.create();
  const aadContext = { path: "/tmp/test.bin" };

  const plaintext = Buffer.from("not a sqlite file", "utf8");
  const encrypted = ring.encryptBytes(plaintext, { aadContext });
  const encoded = encodeEncryptedFileBytes({
    keyVersion: encrypted.keyVersion,
    iv: encrypted.iv,
    tag: encrypted.tag,
    ciphertext: encrypted.ciphertext
  });

  assert.equal(isEncryptedFileBytes(encoded), true);

  const decoded = decodeEncryptedFileBytes(encoded);
  const decrypted = ring.decryptBytes(decoded, { aadContext });
  assert.equal(decrypted.toString("utf8"), "not a sqlite file");
});

