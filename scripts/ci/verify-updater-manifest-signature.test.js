import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { generateKeyPairSync, sign } from "node:crypto";
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
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
function run({ configPath, latestJsonPath, latestSigPath, env = {} }) {
  const proc = spawnSync(process.execPath, [scriptPath, latestJsonPath, latestSigPath], {
    encoding: "utf8",
    cwd: repoRoot,
    env: {
      ...process.env,
      FORMULA_TAURI_CONF_PATH: configPath,
      ...env,
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

test("appends a short note to GITHUB_STEP_SUMMARY when set", () => {
  const tmp = makeTempDir();
  const { publicKey, privateKey } = generateKeyPairSync("ed25519");
  const rawPubkey = rawEd25519PublicKey(publicKey);

  const latestJsonPath = path.join(tmp, "latest.json");
  const latestSigPath = path.join(tmp, "latest.json.sig");
  const configPath = path.join(tmp, "tauri.conf.json");
  const stepSummaryPath = path.join(tmp, "step-summary.md");

  const latestJsonBytes = Buffer.from(JSON.stringify({ version: "0.1.0", platforms: {} }), "utf8");
  const signature = sign(null, latestJsonBytes, privateKey);

  writeFileSync(latestJsonPath, latestJsonBytes);
  writeFileSync(latestSigPath, Buffer.from(signature).toString("base64"));
  writeFileSync(
    configPath,
    JSON.stringify({ plugins: { updater: { pubkey: rawPubkey.toString("base64") } } }, null, 2),
  );
  writeFileSync(stepSummaryPath, "## Updater manifest validation\n\n", "utf8");

  const proc = run({
    configPath,
    latestJsonPath,
    latestSigPath,
    env: { GITHUB_STEP_SUMMARY: stepSummaryPath },
  });
  assert.equal(proc.status, 0, proc.stderr);
  const summary = readFileSync(stepSummaryPath, "utf8");
  assert.match(summary, /Manifest signature: OK/);
  assert.match(summary, /Updater pubkey: raw/);
});

test("supports base64url for both pubkey and signature", () => {
  const tmp = makeTempDir();
  const { publicKey, privateKey } = generateKeyPairSync("ed25519");
  const rawPubkey = rawEd25519PublicKey(publicKey);

  const latestJsonPath = path.join(tmp, "latest.json");
  const latestSigPath = path.join(tmp, "latest.json.sig");
  const configPath = path.join(tmp, "tauri.conf.json");

  const latestJsonBytes = Buffer.from(JSON.stringify({ version: "0.1.0", platforms: {} }), "utf8");
  const signature = sign(null, latestJsonBytes, privateKey);

  const toBase64Url = (b64) => b64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");

  writeFileSync(latestJsonPath, latestJsonBytes);
  writeFileSync(latestSigPath, toBase64Url(Buffer.from(signature).toString("base64")));
  writeFileSync(
    configPath,
    JSON.stringify(
      { plugins: { updater: { pubkey: toBase64Url(rawPubkey.toString("base64")) } } },
      null,
      2,
    ),
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

test("supports minisign signature *files* (comment + payload lines)", () => {
  const tmp = makeTempDir();
  const { publicKey, privateKey } = generateKeyPairSync("ed25519");
  const rawPubkey = rawEd25519PublicKey(publicKey);

  // minisign key id is metadata; for testing we can use any 8 bytes as long as they match.
  const keyId = Buffer.from([8, 7, 6, 5, 4, 3, 2, 1]);
  const keyIdHex = Buffer.from(keyId).reverse().toString("hex").toUpperCase();

  const pubPayload = Buffer.concat([Buffer.from([0x45, 0x64]), keyId, rawPubkey]); // "Ed" + keyId + pubkey
  const pubPayloadLine = pubPayload.toString("base64").replace(/=+$/, "");
  const minisignPubkeyFile = `untrusted comment: minisign public key: ${keyIdHex}\n${pubPayloadLine}\n`;
  const tauriPubkey = Buffer.from(minisignPubkeyFile, "utf8").toString("base64");

  const latestJsonPath = path.join(tmp, "latest.json");
  const latestSigPath = path.join(tmp, "latest.json.sig");
  const configPath = path.join(tmp, "tauri.conf.json");

  const latestJsonBytes = Buffer.from(JSON.stringify({ version: "0.1.0", platforms: {} }), "utf8");
  const signature = sign(null, latestJsonBytes, privateKey);
  const minisignSigPayload = Buffer.concat([Buffer.from([0x45, 0x64]), keyId, Buffer.from(signature)]);
  const sigPayloadLine = minisignSigPayload.toString("base64").replace(/=+$/, "");

  const minisignSigFile = `untrusted comment: signature from minisign secret key\n${sigPayloadLine}\ntrusted comment: timestamp: 0\nAAAA\n`;

  writeFileSync(latestJsonPath, latestJsonBytes);
  writeFileSync(latestSigPath, minisignSigFile, "utf8");
  writeFileSync(configPath, JSON.stringify({ plugins: { updater: { pubkey: tauriPubkey } } }, null, 2));

  const proc = run({ configPath, latestJsonPath, latestSigPath });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /signature OK/);
});

test("fails when minisign signature comment key id does not match payload key id", () => {
  const tmp = makeTempDir();
  const { publicKey, privateKey } = generateKeyPairSync("ed25519");
  const rawPubkey = rawEd25519PublicKey(publicKey);

  const pubKeyId = Buffer.from([9, 9, 9, 9, 9, 9, 9, 9]);
  const sigKeyId = Buffer.from([7, 7, 7, 7, 7, 7, 7, 7]);

  const pubKeyIdHex = Buffer.from(pubKeyId).reverse().toString("hex").toUpperCase();
  const sigKeyIdHex = Buffer.from(sigKeyId).reverse().toString("hex").toUpperCase();

  const pubPayload = Buffer.concat([Buffer.from([0x45, 0x64]), pubKeyId, rawPubkey]);
  const pubPayloadLine = pubPayload.toString("base64").replace(/=+$/, "");
  const minisignPubkeyFile = `untrusted comment: minisign public key: ${pubKeyIdHex}\n${pubPayloadLine}\n`;
  const tauriPubkey = Buffer.from(minisignPubkeyFile, "utf8").toString("base64");

  const latestJsonPath = path.join(tmp, "latest.json");
  const latestSigPath = path.join(tmp, "latest.json.sig");
  const configPath = path.join(tmp, "tauri.conf.json");

  const latestJsonBytes = Buffer.from(JSON.stringify({ version: "0.1.0", platforms: {} }), "utf8");
  const signature = sign(null, latestJsonBytes, privateKey);

  const minisignSigPayload = Buffer.concat([Buffer.from([0x45, 0x64]), sigKeyId, Buffer.from(signature)]);
  const sigPayloadLine = minisignSigPayload.toString("base64").replace(/=+$/, "");
  const minisignSigFile = `untrusted comment: minisign signature: ${pubKeyIdHex}\n${sigPayloadLine}\n`;

  writeFileSync(latestJsonPath, latestJsonBytes);
  writeFileSync(latestSigPath, minisignSigFile, "utf8");
  writeFileSync(configPath, JSON.stringify({ plugins: { updater: { pubkey: tauriPubkey } } }, null, 2));

  const proc = run({ configPath, latestJsonPath, latestSigPath });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /comment key id/i);
  assert.match(proc.stderr, new RegExp(pubKeyIdHex));
  assert.match(proc.stderr, new RegExp(sigKeyIdHex));
});

test("fails when minisign pubkey comment key id does not match payload key id", () => {
  const tmp = makeTempDir();
  const { publicKey, privateKey } = generateKeyPairSync("ed25519");
  const rawPubkey = rawEd25519PublicKey(publicKey);

  const keyId = Buffer.from([1, 1, 1, 1, 1, 1, 1, 1]);
  const keyIdHexWrong = "0000000000000000";

  const pubPayload = Buffer.concat([Buffer.from([0x45, 0x64]), keyId, rawPubkey]);
  const pubPayloadLine = pubPayload.toString("base64").replace(/=+$/, "");
  const minisignPubkeyFile = `untrusted comment: minisign public key: ${keyIdHexWrong}\n${pubPayloadLine}\n`;
  const tauriPubkey = Buffer.from(minisignPubkeyFile, "utf8").toString("base64");

  const latestJsonPath = path.join(tmp, "latest.json");
  const latestSigPath = path.join(tmp, "latest.json.sig");
  const configPath = path.join(tmp, "tauri.conf.json");

  const latestJsonBytes = Buffer.from(JSON.stringify({ version: "0.1.0", platforms: {} }), "utf8");
  const signature = sign(null, latestJsonBytes, privateKey);

  writeFileSync(latestJsonPath, latestJsonBytes);
  writeFileSync(latestSigPath, Buffer.from(signature).toString("base64"));
  writeFileSync(configPath, JSON.stringify({ plugins: { updater: { pubkey: tauriPubkey } } }, null, 2));

  const proc = run({ configPath, latestJsonPath, latestSigPath });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /comment key id/i);
});

test("fails when minisign signature key id does not match the pubkey key id", () => {
  const tmp = makeTempDir();
  const { publicKey, privateKey } = generateKeyPairSync("ed25519");
  const rawPubkey = rawEd25519PublicKey(publicKey);

  const pubKeyId = Buffer.from([1, 1, 1, 1, 1, 1, 1, 1]);
  const sigKeyId = Buffer.from([2, 2, 2, 2, 2, 2, 2, 2]);

  const pubKeyIdHex = Buffer.from(pubKeyId).reverse().toString("hex").toUpperCase();
  const sigKeyIdHex = Buffer.from(sigKeyId).reverse().toString("hex").toUpperCase();

  const pubPayload = Buffer.concat([Buffer.from([0x45, 0x64]), pubKeyId, rawPubkey]); // "Ed" + keyId + pubkey
  const pubPayloadLine = pubPayload.toString("base64").replace(/=+$/, "");
  const minisignPubkeyFile = `untrusted comment: minisign public key: ${pubKeyIdHex}\n${pubPayloadLine}\n`;
  const tauriPubkey = Buffer.from(minisignPubkeyFile, "utf8").toString("base64");

  const latestJsonPath = path.join(tmp, "latest.json");
  const latestSigPath = path.join(tmp, "latest.json.sig");
  const configPath = path.join(tmp, "tauri.conf.json");

  const latestJsonBytes = Buffer.from(JSON.stringify({ version: "0.1.0", platforms: {} }), "utf8");
  const signature = sign(null, latestJsonBytes, privateKey);

  // Signature payload uses a different key id than the pubkey. (The signature bytes themselves are
  // valid, but the metadata mismatch should fail fast.)
  const minisignSigPayload = Buffer.concat([Buffer.from([0x45, 0x64]), sigKeyId, Buffer.from(signature)]);
  writeFileSync(latestJsonPath, latestJsonBytes);
  writeFileSync(latestSigPath, minisignSigPayload.toString("base64").replace(/=+$/, ""));
  writeFileSync(
    configPath,
    JSON.stringify({ plugins: { updater: { pubkey: tauriPubkey } } }, null, 2),
  );

  const proc = run({ configPath, latestJsonPath, latestSigPath });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /key id mismatch/i);
  assert.match(proc.stderr, new RegExp(sigKeyIdHex));
  assert.match(proc.stderr, new RegExp(pubKeyIdHex));
});

test("fails with a clear error when latest.json.sig is not valid base64", () => {
  const tmp = makeTempDir();
  const { publicKey } = generateKeyPairSync("ed25519");
  const rawPubkey = rawEd25519PublicKey(publicKey);

  const latestJsonPath = path.join(tmp, "latest.json");
  const latestSigPath = path.join(tmp, "latest.json.sig");
  const configPath = path.join(tmp, "tauri.conf.json");

  writeFileSync(latestJsonPath, "{}");
  writeFileSync(latestSigPath, "definitely-not-base64");
  writeFileSync(
    configPath,
    JSON.stringify({ plugins: { updater: { pubkey: rawPubkey.toString("base64") } } }, null, 2),
  );

  const proc = run({ configPath, latestJsonPath, latestSigPath });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /not valid base64/i);
});
