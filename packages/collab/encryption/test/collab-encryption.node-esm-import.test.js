import assert from "node:assert/strict";
import test from "node:test";

// Include an explicit `.ts` import specifier so the repo's node:test runner can
// automatically skip this suite when TypeScript execution isn't available.
import { decryptCellPlaintext as decryptFromTs, encryptCellPlaintext as encryptFromTs } from "../src/index.node.ts";

test("collab-encryption is importable under Node ESM when executing TS sources directly", async () => {
  const mod = await import("@formula/collab-encryption");

  assert.equal(typeof mod.encryptCellPlaintext, "function");
  assert.equal(typeof mod.decryptCellPlaintext, "function");
  assert.equal(typeof mod.isEncryptedCellPayload, "function");
  assert.equal(typeof encryptFromTs, "function");
  assert.equal(typeof decryptFromTs, "function");

  const keyBytes = new Uint8Array(32);
  keyBytes.fill(7);
  const key = { keyId: "k1", keyBytes };
  const context = { docId: "d1", sheetId: "Sheet1", row: 0, col: 0 };
  const plaintext = { value: "hello", formula: null };

  const encrypted = await mod.encryptCellPlaintext({ plaintext, key, context });
  assert.equal(mod.isEncryptedCellPayload(encrypted), true);

  const decrypted = await mod.decryptCellPlaintext({ encrypted, key, context });
  assert.deepEqual(decrypted, plaintext);
});
