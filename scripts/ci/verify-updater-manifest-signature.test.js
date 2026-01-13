import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { generateKeyPairSync, sign } from "node:crypto";
import { mkdirSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const scriptPath = path.join(repoRoot, "scripts", "ci", "verify-updater-manifest-signature.mjs");

/**
 * Extract the raw 32-byte Ed25519 public key from a Node KeyObject.
 * @param {import("node:crypto").KeyObject} publicKey
 */
function rawEd25519PublicKey(publicKey) {
  const spki = publicKey.export({ format: "der", type: "spki" });
  const prefix = Buffer.from("302a300506032b6570032100", "hex");
  assert.ok(
    Buffer.from(spki).subarray(0, prefix.length).equals(prefix),
    `Unexpected Ed25519 SPKI prefix: ${Buffer.from(spki).toString("hex")}`,
  );
  return Buffer.from(spki).subarray(prefix.length);
}

/**
 * Run the verifier script.
 * @param {{ configPath: string; latestJsonPath: string; latestSigPath: string }}
 */
function run({ configPath, latestJsonPath, latestSigPath }) {
  const proc = spawnSync(process.execPath, [scriptPath, latestJsonPath, latestSigPath], {
    encoding: "utf8",
    cwd: repoRoot,
    env: {
      ...process.env,
      FORMULA_TAURI_CONF_PATH: configPath,
    },
  });
  if (proc.error) throw proc.error;
  return proc;
}

function makeTempDir() {
  const dir = path.join(os.tmpdir(), `formula-updater-sig-${Date.now()}-${Math.random().toString(16).slice(2)}`);
  mkdirSync(dir, { recursive: true });
  return dir;
}

test("prints \"signature OK\" when latest.json.sig matches latest.json (raw pubkey + raw signature)", () => {
  const tmp = makeTempDir();
  const { publicKey, privateKey } = generateKeyPairSync("ed25519");
  const rawPubkey = rawEd25519PublicKey(publicKey);

  const latestJsonPath = path.join(tmp, "latest.json");
  const latestSigPath = path.join(tmp, "latest.json.sig");
  const configPath = path.join(tmp, "tauri.conf.json");

  const latestJsonBytes = Buffer.from(JSON.stringify({ version: "0.1.0", platforms: {} }), "utf8");
  const signature = sign(null, latestJsonBytes, privateKey);

  writeFileSync(latestJsonPath, latestJsonBytes);
  writeFileSync(latestSigPath, Buffer.from(signature).toString("base64"));
  writeFileSync(
    configPath,
    JSON.stringify({ plugins: { updater: { pubkey: rawPubkey.toString("base64") } } }, null, 2),
  );

  const proc = run({ configPath, latestJsonPath, latestSigPath });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /signature OK/);
});

test("fails when latest.json.sig does not match latest.json", () => {
  const tmp = makeTempDir();
  const { publicKey, privateKey } = generateKeyPairSync("ed25519");
  const rawPubkey = rawEd25519PublicKey(publicKey);

  const latestJsonPath = path.join(tmp, "latest.json");
  const latestSigPath = path.join(tmp, "latest.json.sig");
  const configPath = path.join(tmp, "tauri.conf.json");

  const latestJsonBytes = Buffer.from(JSON.stringify({ version: "0.1.0", platforms: {} }), "utf8");
  const signature = sign(null, Buffer.from("different-bytes"), privateKey);

  writeFileSync(latestJsonPath, latestJsonBytes);
  writeFileSync(latestSigPath, Buffer.from(signature).toString("base64"));
  writeFileSync(
    configPath,
    JSON.stringify({ plugins: { updater: { pubkey: rawPubkey.toString("base64") } } }, null, 2),
  );

  const proc = run({ configPath, latestJsonPath, latestSigPath });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /signature mismatch/i);
});

test("fails when plugins.updater.pubkey is still set to the placeholder value", () => {
  const tmp = makeTempDir();
  const latestJsonPath = path.join(tmp, "latest.json");
  const latestSigPath = path.join(tmp, "latest.json.sig");
  const configPath = path.join(tmp, "tauri.conf.json");

  writeFileSync(latestJsonPath, "{}");
  writeFileSync(latestSigPath, "AA==");
  writeFileSync(
    configPath,
    JSON.stringify(
      { plugins: { updater: { pubkey: "REPLACE_WITH_TAURI_UPDATER_PUBLIC_KEY" } } },
      null,
      2,
    ),
  );

  const proc = run({ configPath, latestJsonPath, latestSigPath });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /placeholder/i);
});

test("supports minisign key/signature structures (\"Ed\" + keyId + bytes)", () => {
  const tmp = makeTempDir();
  const { publicKey, privateKey } = generateKeyPairSync("ed25519");
  const rawPubkey = rawEd25519PublicKey(publicKey);

  // minisign key id is metadata; for testing we can use any 8 bytes as long as they match.
  const keyId = Buffer.from([1, 2, 3, 4, 5, 6, 7, 8]);
  const keyIdHex = Buffer.from(keyId).reverse().toString("hex").toUpperCase();

  const pubPayload = Buffer.concat([Buffer.from([0x45, 0x64]), keyId, rawPubkey]); // "Ed" + keyId + pubkey
  const pubPayloadLine = pubPayload.toString("base64").replace(/=+$/, ""); // mimic tauri output (unpadded)
  const minisignPubkeyFile = `untrusted comment: minisign public key: ${keyIdHex}\n${pubPayloadLine}\n`;
  const tauriPubkey = Buffer.from(minisignPubkeyFile, "utf8").toString("base64");

  const latestJsonPath = path.join(tmp, "latest.json");
  const latestSigPath = path.join(tmp, "latest.json.sig");
  const configPath = path.join(tmp, "tauri.conf.json");

  const latestJsonBytes = Buffer.from(JSON.stringify({ version: "0.1.0", platforms: {} }), "utf8");
  const signature = sign(null, latestJsonBytes, privateKey);
  const minisignSig = Buffer.concat([Buffer.from([0x45, 0x64]), keyId, Buffer.from(signature)]);

  writeFileSync(latestJsonPath, latestJsonBytes);
  writeFileSync(latestSigPath, minisignSig.toString("base64").replace(/=+$/, ""));
  writeFileSync(
    configPath,
    JSON.stringify({ plugins: { updater: { pubkey: tauriPubkey } } }, null, 2),
  );

  const proc = run({ configPath, latestJsonPath, latestSigPath });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /signature OK/);
});

