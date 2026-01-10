import test from "node:test";
import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import { promises as fs } from "node:fs";

import { InMemoryKeychainProvider } from "../../../../../../packages/security/crypto/keychain/inMemoryKeychain.js";
import { DesktopEncryptedDocumentStore } from "../encryptedDocumentStore.js";

async function mkTempDir() {
  return fs.mkdtemp(path.join(os.tmpdir(), "formula-desktop-store-"));
}

test("desktop: enable encryption -> ciphertext on disk; disable -> plaintext on disk", async () => {
  const dir = await mkTempDir();
  const filePath = path.join(dir, "store.json");
  const keychain = new InMemoryKeychainProvider();

  const store = new DesktopEncryptedDocumentStore({ filePath, keychainProvider: keychain });

  await store.enableEncryption();
  await store.saveDocument("doc-1", { name: "Secret Document", value: 42 });

  const encryptedRaw = await fs.readFile(filePath, "utf8");
  assert.ok(!encryptedRaw.includes("Secret Document"), "expected ciphertext-only on disk");
  assert.ok(encryptedRaw.includes("\"ciphertext\""), "expected encrypted payload fields on disk");

  const storeReload = new DesktopEncryptedDocumentStore({ filePath, keychainProvider: keychain });
  const loadedEncrypted = await storeReload.loadDocument("doc-1");
  assert.deepEqual(loadedEncrypted, { name: "Secret Document", value: 42 });

  await storeReload.disableEncryption();

  assert.equal(
    await keychain.getSecret({ service: "formula.desktop", account: "storage-keyring" }),
    null,
    "expected encryption key to be removed from keychain on disable"
  );

  const plaintextRaw = await fs.readFile(filePath, "utf8");
  assert.ok(plaintextRaw.includes("Secret Document"), "expected plaintext after disabling encryption");

  const storePlain = new DesktopEncryptedDocumentStore({ filePath, keychainProvider: keychain });
  const loadedPlain = await storePlain.loadDocument("doc-1");
  assert.deepEqual(loadedPlain, { name: "Secret Document", value: 42 });
});
