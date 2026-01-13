import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs/promises";
import { fileURLToPath } from "node:url";

import {
  decryptAes256Gcm,
  encryptAes256Gcm,
  serializeEncryptedPayload
} from "../aes256gcm.js";
import { aadFromContext, fromBase64 } from "../utils.js";

const FIXTURE_PATH = fileURLToPath(
  new URL("../../../../fixtures/crypto/desktop-storage-encryption-v1.json", import.meta.url)
);

test("desktop storage AES-256-GCM vectors v1 match Rust implementation", async () => {
  const raw = await fs.readFile(FIXTURE_PATH, "utf8");
  const fixture = JSON.parse(raw);

  const key = fromBase64(fixture.key, "key");
  const iv = fromBase64(fixture.iv, "iv");
  const aad = aadFromContext(fixture.aadContext);

  let plaintext;
  switch (fixture.plaintext?.encoding) {
    case "utf8":
      plaintext = Buffer.from(fixture.plaintext.value, "utf8");
      break;
    case "base64":
      plaintext = Buffer.from(fixture.plaintext.value, "base64");
      break;
    default:
      throw new Error(`Unsupported plaintext encoding: ${fixture.plaintext?.encoding}`);
  }

  const encrypted = encryptAes256Gcm({ plaintext, key, aad, iv });
  const serialized = serializeEncryptedPayload(encrypted);

  assert.deepStrictEqual(serialized, {
    algorithm: fixture.algorithm,
    iv: fixture.iv,
    ciphertext: fixture.expected.ciphertext,
    tag: fixture.expected.tag
  });

  const decrypted = decryptAes256Gcm({ ...encrypted, key, aad });
  assert.deepStrictEqual(decrypted, plaintext);
});

