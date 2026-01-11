import assert from "node:assert/strict";
import { Writable } from "node:stream";
import test from "node:test";

import { createLogger } from "../src/logger.js";

test("createLogger redacts internal admin headers + persistence encryption keys", () => {
  const chunks: string[] = [];
  const destination = new Writable({
    write(chunk, _encoding, callback) {
      chunks.push(chunk.toString());
      callback();
    },
  });

  const logger = createLogger("info", destination);

  const internalAdminToken = "token-should-not-appear";
  const encryptionKeyBase64 = "key-should-not-appear";
  const leveldbEncryptionKey = "leveldb-key-should-not-appear";
  const keyRingSecret = "keyring-should-not-appear";
  const authToken = "auth-token-should-not-appear";
  const jwtSecret = "jwt-secret-should-not-appear";
  const introspectionToken = "introspection-token-should-not-appear";

  logger.info(
    {
      req: {
        headers: {
          authorization: "Bearer abc123",
          "x-internal-admin-token": internalAdminToken,
          "x-sync-server-admin-token": internalAdminToken,
        },
      },
      auth: { token: authToken, secret: jwtSecret },
      introspection: { token: introspectionToken },
      internalAdminToken,
      persistence: {
        encryption: { keyBase64: encryptionKeyBase64, keyRing: { secret: keyRingSecret } },
        leveldbEncryption: { key: leveldbEncryptionKey, strict: true },
      },
      config: {
        persistence: {
          encryption: { keyRing: { secret: keyRingSecret } },
          leveldbEncryption: { key: leveldbEncryptionKey, strict: true },
        },
        auth: { token: authToken, secret: jwtSecret },
        introspection: { token: introspectionToken },
        internalAdminToken,
      },
    },
    "redaction_test"
  );

  const lines = chunks.join("").split("\n").filter(Boolean);
  assert.equal(lines.length, 1);

  const raw = lines[0]!;
  assert.ok(!raw.includes(internalAdminToken));
  assert.ok(!raw.includes(encryptionKeyBase64));
  assert.ok(!raw.includes(leveldbEncryptionKey));
  assert.ok(!raw.includes(keyRingSecret));
  assert.ok(!raw.includes(authToken));
  assert.ok(!raw.includes(jwtSecret));
  assert.ok(!raw.includes(introspectionToken));

  const parsed = JSON.parse(raw) as any;
  assert.ok(!("authorization" in (parsed.req?.headers ?? {})));
  assert.ok(!("x-internal-admin-token" in (parsed.req?.headers ?? {})));
  assert.ok(!("x-sync-server-admin-token" in (parsed.req?.headers ?? {})));
  assert.ok(!("token" in (parsed.auth ?? {})));
  assert.ok(!("secret" in (parsed.auth ?? {})));
  assert.ok(!("token" in (parsed.introspection ?? {})));
  assert.ok(!("internalAdminToken" in parsed));
  assert.ok(!("keyBase64" in (parsed.persistence?.encryption ?? {})));
  assert.ok(!("keyRing" in (parsed.persistence?.encryption ?? {})));
  assert.ok(!("key" in (parsed.persistence?.leveldbEncryption ?? {})));
  assert.ok(!("keyRing" in (parsed.config?.persistence?.encryption ?? {})));
  assert.ok(!("key" in (parsed.config?.persistence?.leveldbEncryption ?? {})));
  assert.ok(!("token" in (parsed.config?.auth ?? {})));
  assert.ok(!("secret" in (parsed.config?.auth ?? {})));
  assert.ok(!("token" in (parsed.config?.introspection ?? {})));
  assert.ok(!("internalAdminToken" in (parsed.config ?? {})));
});
