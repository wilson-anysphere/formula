import test from "node:test";
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import crypto from "node:crypto";
import zlib from "node:zlib";

// Basic smoke test to ensure the CLI can verify v1 packages when provided
// a detached signature (used during the migration period).

test("extension CLI verifies v1 package with detached signature", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-extension-cli-"));
  try {
    const pkgPath = path.join(tmpRoot, "pkg.fextpkg");
    const pubPath = path.join(tmpRoot, "pub.pem");
    const privPath = path.join(tmpRoot, "priv.pem");

    const { publicKey, privateKey } = crypto.generateKeyPairSync("ed25519");
    await fs.writeFile(pubPath, publicKey.export({ type: "spki", format: "pem" }));
    await fs.writeFile(privPath, privateKey.export({ type: "pkcs8", format: "pem" }));

    const bundle = {
      format: "formula-extension-package",
      formatVersion: 1,
      createdAt: "2020-01-01T00:00:00.000Z",
      manifest: { name: "x", publisher: "p", version: "1.0.0", main: "./dist/extension.js", engines: { formula: "^1.0.0" } },
      files: [],
    };
    const gz = zlib.gzipSync(Buffer.from(JSON.stringify(bundle), "utf8"));
    await fs.writeFile(pkgPath, gz);

    const signature = crypto.sign(null, gz, privateKey).toString("base64");

    const proc = spawnSync(process.execPath, ["scripts/extension.mjs", "verify", pkgPath, "--pubkey", pubPath, "--signature", signature], {
      cwd: path.resolve("."),
      encoding: "utf8",
    });

    assert.equal(proc.status, 0, proc.stderr || proc.stdout);
    const out = JSON.parse(proc.stdout);
    assert.equal(out.ok, true);
    assert.equal(out.formatVersion, 1);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

