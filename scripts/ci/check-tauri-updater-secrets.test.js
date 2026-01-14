import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { generateKeyPairSync } from "node:crypto";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const scriptPath = path.join(repoRoot, "scripts", "ci", "check-tauri-updater-secrets.mjs");

const configPath = path.join(repoRoot, "apps", "desktop", "src-tauri", "tauri.conf.json");
const tauriConfig = JSON.parse(readFileSync(configPath, "utf8"));
const updaterPubkey = tauriConfig?.plugins?.updater?.pubkey;
assert.equal(typeof updaterPubkey, "string");

function updaterKeyIdBytes() {
  const decoded = Buffer.from(updaterPubkey, "base64").toString("utf8");
  const lines = decoded
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);
  const payload = lines.find((line) => !line.toLowerCase().startsWith("untrusted comment:"));
  assert.ok(payload, "expected minisign payload line in updater public key");
  const binary = Buffer.from(payload, "base64");
  assert.equal(binary.slice(0, 2).toString("ascii"), "Ed");
  return binary.subarray(2, 10);
}

const updaterKeyId = updaterKeyIdBytes();

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

function fakeMinisignSecretKey({ encrypted, keyId = updaterKeyId }) {
  // minisign keys are binary payloads prefixed with "Ed" + 8-byte key id.
  const header = Buffer.from([0x45, 0x64]); // "Ed"
  const secretPayload = Buffer.alloc(encrypted ? 140 : 64, 0x22);
  const binary = Buffer.concat([header, keyId, secretPayload]);
  const payloadLine = binary.toString("base64").replace(/=+$/, "");
  // Match minisign's displayed key ID format (big-endian hex).
  const keyIdHex = Buffer.from(keyId).reverse().toString("hex").toUpperCase();
  const keyFile = `untrusted comment: minisign secret key: ${keyIdHex}\n${payloadLine}\n`;

  // `cargo tauri signer generate` prints base64 strings that decode to minisign key files.
  return Buffer.from(keyFile, "utf8").toString("base64");
}

test("passes with an unencrypted minisign secret key and empty password", () => {
  const key = fakeMinisignSecretKey({ encrypted: false });
  const proc = run({ TAURI_PRIVATE_KEY: key, TAURI_KEY_PASSWORD: "" });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /preflight passed/i);
});

test("fails when minisign TAURI_PRIVATE_KEY does not match plugins.updater.pubkey", () => {
  const wrongKeyId = Buffer.alloc(8, 0x33);
  const key = fakeMinisignSecretKey({ encrypted: false, keyId: wrongKeyId });
  const proc = run({ TAURI_PRIVATE_KEY: key, TAURI_KEY_PASSWORD: "" });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /Tauri updater key mismatch/i);
});

test("fails for encrypted minisign secret keys (unsupported by release tooling)", () => {
  const key = fakeMinisignSecretKey({ encrypted: true });
  {
    const proc = run({ TAURI_PRIVATE_KEY: key, TAURI_KEY_PASSWORD: "" });
    assert.notEqual(proc.status, 0);
    assert.match(proc.stderr, /Encrypted minisign TAURI_PRIVATE_KEY is not supported/);
  }
  {
    const proc = run({ TAURI_PRIVATE_KEY: key, TAURI_KEY_PASSWORD: "pass" });
    assert.notEqual(proc.status, 0);
    assert.match(proc.stderr, /Encrypted minisign TAURI_PRIVATE_KEY is not supported/);
  }
});

test("fails when TAURI_PRIVATE_KEY is a minisign public key (copied from tauri.conf.json)", () => {
  const proc = run({ TAURI_PRIVATE_KEY: updaterPubkey, TAURI_KEY_PASSWORD: "pass" });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /minisign \*public\* key/i);
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
