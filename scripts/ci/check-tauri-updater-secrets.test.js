import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { generateKeyPairSync } from "node:crypto";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const scriptPath = path.join(repoRoot, "scripts", "ci", "check-tauri-updater-secrets.mjs");

/**
 * @param {Record<string, string | undefined>} env
 */
function run(env) {
  const proc = spawnSync(process.execPath, [scriptPath], {
    encoding: "utf8",
    cwd: repoRoot,
    env: {
      ...process.env,
      ...env,
    },
  });
  if (proc.error) throw proc.error;
  return proc;
}

test("fails when TAURI_PRIVATE_KEY is missing", () => {
  const proc = run({ TAURI_PRIVATE_KEY: undefined, TAURI_KEY_PASSWORD: undefined });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /Missing Tauri updater signing secrets/);
  assert.match(proc.stderr, /\bTAURI_PRIVATE_KEY\b/);
});

test("fails when TAURI_PRIVATE_KEY is set but not a supported format", () => {
  const proc = run({ TAURI_PRIVATE_KEY: "definitely-not-a-key", TAURI_KEY_PASSWORD: "pass" });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /Invalid TAURI_PRIVATE_KEY/);
});

test("passes with an unencrypted raw Ed25519 key and empty password", () => {
  const raw = Buffer.alloc(64, 1).toString("base64").replace(/=+$/, ""); // unpadded base64
  const proc = run({ TAURI_PRIVATE_KEY: raw, TAURI_KEY_PASSWORD: "" });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /preflight passed/i);
});

test("passes with an unencrypted PKCS#8 DER key and empty password", () => {
  const { privateKey } = generateKeyPairSync("ed25519");
  const der = privateKey.export({ format: "der", type: "pkcs8" });
  const proc = run({ TAURI_PRIVATE_KEY: Buffer.from(der).toString("base64"), TAURI_KEY_PASSWORD: "" });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /preflight passed/i);
});

test("requires TAURI_KEY_PASSWORD for encrypted private keys", () => {
  const { privateKey } = generateKeyPairSync("ed25519");
  const pem = privateKey.export({
    format: "pem",
    type: "pkcs8",
    cipher: "aes-256-cbc",
    passphrase: "pass",
  });

  {
    const proc = run({ TAURI_PRIVATE_KEY: String(pem), TAURI_KEY_PASSWORD: "" });
    assert.notEqual(proc.status, 0);
    assert.match(proc.stderr, /\bTAURI_KEY_PASSWORD\b/);
  }

  {
    const proc = run({ TAURI_PRIVATE_KEY: String(pem), TAURI_KEY_PASSWORD: "pass" });
    assert.equal(proc.status, 0, proc.stderr);
    assert.match(proc.stdout, /preflight passed/i);
  }
});

test("passes with base64-encoded unencrypted PEM and empty password", () => {
  const { privateKey } = generateKeyPairSync("ed25519");
  const pem = privateKey.export({ format: "pem", type: "pkcs8" });
  const b64 = Buffer.from(String(pem), "utf8").toString("base64");
  const proc = run({ TAURI_PRIVATE_KEY: b64, TAURI_KEY_PASSWORD: "" });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /preflight passed/i);
});

test("requires TAURI_KEY_PASSWORD for encrypted PKCS#8 DER keys", () => {
  const { privateKey } = generateKeyPairSync("ed25519");
  const der = privateKey.export({
    format: "der",
    type: "pkcs8",
    cipher: "aes-256-cbc",
    passphrase: "pass",
  });
  const b64 = Buffer.from(der).toString("base64");

  {
    const proc = run({ TAURI_PRIVATE_KEY: b64, TAURI_KEY_PASSWORD: "" });
    assert.notEqual(proc.status, 0);
    assert.match(proc.stderr, /\bTAURI_KEY_PASSWORD\b/);
  }

  {
    const proc = run({ TAURI_PRIVATE_KEY: b64, TAURI_KEY_PASSWORD: "pass" });
    assert.equal(proc.status, 0, proc.stderr);
    assert.match(proc.stdout, /preflight passed/i);
  }
});
