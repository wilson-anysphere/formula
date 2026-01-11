import assert from "node:assert/strict";
import test from "node:test";

import { loadConfigFromEnv } from "../src/config.js";

function withEnv<T>(overrides: Record<string, string | undefined>, fn: () => T): T {
  const previous: Record<string, string | undefined> = {};
  for (const key of Object.keys(overrides)) {
    previous[key] = process.env[key];
    const next = overrides[key];
    if (next === undefined) {
      delete process.env[key];
    } else {
      process.env[key] = next;
    }
  }

  try {
    return fn();
  } finally {
    for (const key of Object.keys(overrides)) {
      const prev = previous[key];
      if (prev === undefined) {
        delete process.env[key];
      } else {
        process.env[key] = prev;
      }
    }
  }
}

test("SYNC_SERVER_PERSISTENCE_ENCRYPTION_KEY_B64 enables keyring persistence encryption", () => {
  const keyBase64 = Buffer.alloc(32, 1).toString("base64");

  const config = withEnv(
    {
      NODE_ENV: "test",
      SYNC_SERVER_PERSISTENCE_ENCRYPTION: "off",
      SYNC_SERVER_PERSISTENCE_ENCRYPTION_KEY_B64: keyBase64,
      SYNC_SERVER_ENCRYPTION_KEYRING_JSON: "",
      SYNC_SERVER_ENCRYPTION_KEYRING_PATH: "",
    },
    () => loadConfigFromEnv()
  );

  assert.equal(config.persistence.encryption.mode, "keyring");
  const keyRing = config.persistence.encryption.keyRing;
  const plaintext = Buffer.from("hello");
  const encrypted = keyRing.encryptBytes(plaintext);
  const decrypted = keyRing.decryptBytes(encrypted);
  assert.ok(decrypted.equals(plaintext));
});

test("SYNC_SERVER_PERSISTENCE_ENCRYPTION_KEY_B64 must decode to 32 bytes", () => {
  const keyBase64 = Buffer.alloc(31, 2).toString("base64");

  assert.throws(
    () =>
      withEnv(
        {
          NODE_ENV: "test",
          SYNC_SERVER_PERSISTENCE_ENCRYPTION: "off",
          SYNC_SERVER_PERSISTENCE_ENCRYPTION_KEY_B64: keyBase64,
          SYNC_SERVER_ENCRYPTION_KEYRING_JSON: "",
          SYNC_SERVER_ENCRYPTION_KEYRING_PATH: "",
        },
        () => loadConfigFromEnv()
      ),
    /SYNC_SERVER_PERSISTENCE_ENCRYPTION_KEY_B64 must be a base64-encoded 32-byte/
  );
});

