import assert from "node:assert/strict";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";

import { isEncryptedFileBytes } from "../../../security/crypto/encryptedFile.js";

import { createNodeOAuth2Manager } from "../../src/oauth2/node.js";
import { normalizeScopes } from "../../src/oauth2/tokenStore.js";

test("createNodeOAuth2Manager uses a persistent encrypted token store", async () => {
  const tmpDir = await mkdtemp(path.join(os.tmpdir(), "pq-oauth2-node-"));
  try {
    const filePath = path.join(tmpDir, "oauthTokens.bin");
    const now = () => 1_000;

    const { manager, tokenStore } = createNodeOAuth2Manager({ filePath, now, fetch: async () => new Response("{}", { status: 500 }) });
    manager.registerProvider({ id: "example", clientId: "client", tokenEndpoint: "https://auth.example/token" });

    const { scopesHash, scopes } = normalizeScopes(["read"]);
    await tokenStore.set(
      { providerId: "example", scopesHash },
      { providerId: "example", scopesHash, scopes, refreshToken: "refresh-1" },
    );

    const bytes = await readFile(filePath);
    assert.ok(isEncryptedFileBytes(bytes), "token store file should be encrypted");

    const { tokenStore: tokenStore2 } = createNodeOAuth2Manager({ filePath, now, fetch: async () => new Response("{}", { status: 500 }) });
    const entry = await tokenStore2.get({ providerId: "example", scopesHash });
    assert.equal(entry?.refreshToken, "refresh-1");
  } finally {
    await rm(tmpDir, { recursive: true, force: true });
  }
});

