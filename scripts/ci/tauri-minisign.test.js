import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";
import {
  ed25519PublicKeyFromRaw,
  parseTauriUpdaterPubkey,
  parseTauriUpdaterSignature,
} from "./tauri-minisign.mjs";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");

test("parses the repo's updater pubkey (tauri.conf.json) minisign format", () => {
  const cfgPath = path.join(repoRoot, "apps", "desktop", "src-tauri", "tauri.conf.json");
  const cfg = JSON.parse(fs.readFileSync(cfgPath, "utf8"));
  const pubkeyBase64 = cfg?.plugins?.updater?.pubkey;
  assert.equal(typeof pubkeyBase64, "string");

  const parsed = parseTauriUpdaterPubkey(pubkeyBase64);
  assert.equal(parsed.format, "minisign");
  assert.equal(parsed.publicKeyBytes.length, 32);
  // This is stable as long as we don't rotate updater keys.
  assert.equal(
    parsed.publicKeyBytes.toString("hex"),
    "6ead221cc4c18737a3861ee9db5bff3f48b6b77ccd00ab281dc7a09b5a7972fe",
  );
  assert.equal(parsed.keyId, "86D6B6E8B99EE0D1");
});

test("parses supported signature formats (raw, minisign payload, minisign text)", () => {
  const sig = Buffer.alloc(64, 7);

  // Raw base64 signature (64 bytes).
  {
    const raw = sig.toString("base64");
    const parsed = parseTauriUpdaterSignature(raw);
    assert.equal(parsed.format, "raw");
    assert.deepEqual(parsed.signatureBytes, sig);
    assert.equal(parsed.keyId, null);
  }

  // Minisign payload (74 bytes: "Ed" + keyid_le + signature).
  const keyIdLe = Buffer.from("0102030405060708", "hex");
  const minisignPayload = Buffer.concat([Buffer.from([0x45, 0x64]), keyIdLe, sig]);
  assert.equal(minisignPayload.length, 74);
  const minisignB64 = minisignPayload.toString("base64");

  {
    const parsed = parseTauriUpdaterSignature(minisignB64);
    assert.equal(parsed.format, "minisign");
    assert.deepEqual(parsed.signatureBytes, sig);
    assert.equal(parsed.keyId, "0807060504030201");
  }

  // Minisign 2-line signature file.
  {
    const text = `untrusted comment: signature from minisign secret key\n${minisignB64}\n`;
    const parsed = parseTauriUpdaterSignature(text);
    assert.equal(parsed.format, "minisign");
    assert.deepEqual(parsed.signatureBytes, sig);
    assert.equal(parsed.keyId, "0807060504030201");
  }

  // Minisign signature file comment contains an explicit key id (validate + return).
  {
    const text = `untrusted comment: minisign signature: 0807060504030201\n${minisignB64}\n`;
    const parsed = parseTauriUpdaterSignature(text);
    assert.equal(parsed.format, "minisign");
    assert.deepEqual(parsed.signatureBytes, sig);
    assert.equal(parsed.keyId, "0807060504030201");
  }

  // Minisign signature file comment key id mismatch should be rejected.
  {
    const text = `untrusted comment: minisign signature: 0000000000000000\n${minisignB64}\n`;
    assert.throws(() => parseTauriUpdaterSignature(text), /comment key id/i);
  }

  // Raw signature bytes with a minisign-style comment that includes a key id.
  {
    const rawB64 = sig.toString("base64");
    const text = `untrusted comment: minisign signature: 0807060504030201\n${rawB64}\n`;
    const parsed = parseTauriUpdaterSignature(text);
    assert.equal(parsed.format, "raw");
    assert.deepEqual(parsed.signatureBytes, sig);
    assert.equal(parsed.keyId, "0807060504030201");
  }
});

test("parses pubkey formats (raw Ed25519 bytes, minisign payload bytes)", () => {
  // Raw 32-byte Ed25519 public key (base64 encoded).
  {
    const raw = Buffer.alloc(32, 9);
    const parsed = parseTauriUpdaterPubkey(raw.toString("base64"));
    assert.equal(parsed.format, "raw");
    assert.equal(parsed.keyId, null);
    assert.deepEqual(parsed.publicKeyBytes, raw);
  }

  // Minisign payload (42 bytes: "Ed" + keyid_le + pubkey), base64 encoded.
  {
    const keyIdLe = Buffer.from("0102030405060708", "hex");
    const pubkey = Buffer.alloc(32, 3);
    const payload = Buffer.concat([Buffer.from([0x45, 0x64]), keyIdLe, pubkey]);
    const parsed = parseTauriUpdaterPubkey(payload.toString("base64"));
    assert.equal(parsed.format, "minisign");
    assert.equal(parsed.keyId, "0807060504030201");
    assert.deepEqual(parsed.publicKeyBytes, pubkey);
  }
});

test("rejects minisign pubkey text when comment key id does not match payload", () => {
  const keyIdLe = Buffer.from("1122334455667788", "hex");
  const pubkeyPayload = Buffer.concat([Buffer.from([0x45, 0x64]), keyIdLe, Buffer.alloc(32, 1)]);
  const wrongKeyId = "0000000000000000";
  const pubkeyText = `untrusted comment: minisign public key: ${wrongKeyId}\n${pubkeyPayload.toString("base64")}\n`;
  const pubkeyBase64 = Buffer.from(pubkeyText, "utf8").toString("base64");
  assert.throws(() => parseTauriUpdaterPubkey(pubkeyBase64), /comment key id/i);
});

test("end-to-end verify works with minisign pubkey + multiple signature formats", () => {
  const { publicKey, privateKey } = crypto.generateKeyPairSync("ed25519");
  const message = Buffer.from("formula updater manifest", "utf8");
  const signature = crypto.sign(null, message, privateKey);
  assert.equal(signature.length, 64);

  // Extract raw 32-byte Ed25519 public key from SPKI (it ends with the raw bytes).
  const spki = /** @type {Buffer} */ (publicKey.export({ format: "der", type: "spki" }));
  const rawPubkey = spki.subarray(spki.length - 32);
  assert.equal(rawPubkey.length, 32);

  const keyIdLe = Buffer.from("1122334455667788", "hex");
  const keyIdHex = Buffer.from(keyIdLe).reverse().toString("hex").toUpperCase();

  // Build a Tauri/minisign public key string: base64(minisign text block).
  const pubkeyPayload = Buffer.concat([Buffer.from([0x45, 0x64]), keyIdLe, rawPubkey]);
  assert.equal(pubkeyPayload.length, 42);
  const pubkeyText = `untrusted comment: minisign public key: ${keyIdHex}\n${pubkeyPayload.toString("base64")}\n`;
  const tauriPubkey = Buffer.from(pubkeyText, "utf8").toString("base64");

  const parsedPubkey = parseTauriUpdaterPubkey(tauriPubkey);
  assert.equal(parsedPubkey.format, "minisign");
  assert.deepEqual(parsedPubkey.publicKeyBytes, rawPubkey);
  assert.equal(parsedPubkey.keyId, keyIdHex);

  const nodePublicKey = ed25519PublicKeyFromRaw(parsedPubkey.publicKeyBytes);

  /** @type {Array<{ name: string; sigText: string }>} */
  const sigCases = [
    { name: "raw base64", sigText: signature.toString("base64") },
    {
      name: "minisign payload base64",
      sigText: Buffer.concat([Buffer.from([0x45, 0x64]), keyIdLe, signature]).toString("base64"),
    },
    {
      name: "minisign text file",
      sigText: `untrusted comment: signature from minisign secret key\n${Buffer.concat([Buffer.from([0x45, 0x64]), keyIdLe, signature]).toString("base64")}\n`,
    },
  ];

  for (const { name, sigText } of sigCases) {
    const parsedSig = parseTauriUpdaterSignature(sigText);
    const ok = crypto.verify(null, message, nodePublicKey, parsedSig.signatureBytes);
    assert.equal(ok, true, `expected signature verification to succeed for case: ${name}`);
  }
});
