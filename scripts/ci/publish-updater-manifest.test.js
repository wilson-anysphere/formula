import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { generateKeyPairSync } from "node:crypto";
import { mkdirSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const scriptPath = path.join(repoRoot, "scripts", "ci", "publish-updater-manifest.mjs");

function makeTempDir() {
  const dir = path.join(
    os.tmpdir(),
    `formula-updater-publish-${Date.now()}-${Math.random().toString(16).slice(2)}`,
  );
  mkdirSync(dir, { recursive: true });
  return dir;
}

/**
 * @param {{ cwd: string; args: string[]; env: Record<string, string | undefined> }}
 */
function run({ cwd, args, env }) {
  const proc = spawnSync(process.execPath, [scriptPath, ...args], {
    encoding: "utf8",
    cwd,
    env: { ...process.env, ...env },
  });
  if (proc.error) throw proc.error;
  return proc;
}

test("fails before uploading if TAURI_PRIVATE_KEY does not match the embedded updater pubkey", () => {
  const tmp = makeTempDir();
  const manifestsDir = path.join(tmp, "manifests");
  mkdirSync(manifestsDir, { recursive: true });

  // Minimal per-platform manifest produced by tauri-action (enough for publish-updater-manifest to merge).
  writeFileSync(
    path.join(manifestsDir, "linux.json"),
    JSON.stringify(
      {
        version: "0.1.0",
        platforms: {
          "linux-x86_64": {
            url: "https://example.com/Formula.AppImage",
            signature: "sig",
          },
        },
      },
      null,
      2,
    ),
  );

  // Generate a random Ed25519 private key that will NOT match the committed updater pubkey.
  const { privateKey } = generateKeyPairSync("ed25519");
  const der = privateKey.export({ format: "der", type: "pkcs8" });
  const tauriPrivateKey = Buffer.from(der).toString("base64");

  const proc = run({
    cwd: tmp,
    args: ["v0.1.0", manifestsDir],
    env: {
      // Provide placeholders so the script gets past required env checks; it should fail before any network call.
      GITHUB_REPOSITORY: "owner/repo",
      GITHUB_TOKEN: "dummy",
      TAURI_PRIVATE_KEY: tauriPrivateKey,
      TAURI_KEY_PASSWORD: "",
    },
  });

  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /does not verify latest\.json/i);

  // If this ever appears, the script got past the key mismatch check and attempted a network call.
  assert.doesNotMatch(proc.stderr, /api\\.github\\.com/i);
});
