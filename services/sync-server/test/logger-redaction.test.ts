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

  logger.info(
    {
      req: {
        headers: {
          authorization: "Bearer abc123",
          "x-internal-admin-token": internalAdminToken,
          "x-sync-server-admin-token": internalAdminToken,
        },
      },
      persistence: {
        encryption: { keyBase64: encryptionKeyBase64 },
        leveldbEncryption: { key: leveldbEncryptionKey, strict: true },
      },
      config: {
        persistence: { leveldbEncryption: { key: leveldbEncryptionKey, strict: true } },
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

  const parsed = JSON.parse(raw) as any;
  assert.ok(!("authorization" in (parsed.req?.headers ?? {})));
  assert.ok(!("x-internal-admin-token" in (parsed.req?.headers ?? {})));
  assert.ok(!("x-sync-server-admin-token" in (parsed.req?.headers ?? {})));
  assert.ok(!("keyBase64" in (parsed.persistence?.encryption ?? {})));
  assert.ok(!("key" in (parsed.persistence?.leveldbEncryption ?? {})));
  assert.ok(!("key" in (parsed.config?.persistence?.leveldbEncryption ?? {})));
});
