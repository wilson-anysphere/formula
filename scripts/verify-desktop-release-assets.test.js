import assert from "node:assert/strict";
import crypto from "node:crypto";
import test from "node:test";
import {
  ActionableError,
  filenameFromUrl,
  validateLatestJson,
  verifyUpdaterManifestSignature,
} from "./verify-desktop-release-assets.mjs";

function assetMap(names) {
  return new Map(names.map((name) => [name, { name }]));
}

test("filenameFromUrl extracts decoded filename and strips query", () => {
  assert.equal(
    filenameFromUrl("https://example.com/download/My%20File.exe?foo=1#bar"),
    "My File.exe",
  );
});

test("validateLatestJson passes for a minimal manifest (version normalization + required OS keys)", () => {
  const manifest = {
    version: "v0.1.0",
    platforms: {
      "linux-x86_64": { url: "https://example.com/Formula_0.1.0_x86_64.AppImage", signature: "sig" },
      "linux-aarch64": { url: "https://example.com/Formula_0.1.0_arm64.AppImage", signature: "sig" },
      "windows-x86_64": { url: "https://example.com/Formula_0.1.0_x64.msi", signature: "sig" },
      "windows-aarch64": { url: "https://example.com/Formula_0.1.0_arm64.msi", signature: "sig" },
      "darwin-universal": { url: "https://example.com/Formula_0.1.0.app.tar.gz", signature: "sig" },
    },
  };

  const assets = assetMap([
    "Formula_0.1.0_x86_64.AppImage",
    "Formula_0.1.0_arm64.AppImage",
    "Formula_0.1.0_x64.msi",
    "Formula_0.1.0_arm64.msi",
    "Formula_0.1.0.app.tar.gz",
  ]);

  assert.doesNotThrow(() => validateLatestJson(manifest, "0.1.0", assets));
});

test("validateLatestJson finds a nested platforms map", () => {
  const manifest = {
    version: "0.1.0",
    data: {
      platforms: {
        "linux-x86_64": { url: "https://example.com/Formula.AppImage", signature: "sig" },
        "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "sig" },
        "windows-x86_64": { url: "https://example.com/Formula_x64.msi", signature: "sig" },
        "windows-aarch64": { url: "https://example.com/Formula_arm64.msi", signature: "sig" },
        "darwin-universal": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
      },
    },
  };

  const assets = assetMap([
    "Formula.AppImage",
    "Formula_arm64.AppImage",
    "Formula_x64.msi",
    "Formula_arm64.msi",
    "Formula.app.tar.gz",
  ]);
  assert.doesNotThrow(() => validateLatestJson(manifest, "0.1.0", assets));
});

test("validateLatestJson fails when a required OS is missing", () => {
  const manifest = {
    version: "0.1.0",
    platforms: {
      "linux-x86_64": { url: "https://example.com/Formula.AppImage", signature: "sig" },
      "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "sig" },
      "windows-x86_64": { url: "https://example.com/Formula_x64.msi", signature: "sig" },
      "windows-aarch64": { url: "https://example.com/Formula_arm64.msi", signature: "sig" },
    },
  };

  const assets = assetMap([
    "Formula.AppImage",
    "Formula_arm64.AppImage",
    "Formula_x64.msi",
    "Formula_arm64.msi",
  ]);
  assert.throws(
    () => validateLatestJson(manifest, "0.1.0", assets),
    (err) => err instanceof ActionableError && err.message.includes("missing required key \"darwin-universal\""),
  );
});

test("validateLatestJson fails when a platform URL references a missing asset", () => {
  const manifest = {
    version: "0.1.0",
    platforms: {
      "linux-x86_64": { url: "https://example.com/Formula.AppImage", signature: "sig" },
      "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "sig" },
      "windows-x86_64": { url: "https://example.com/Formula_x64.msi", signature: "sig" },
      "windows-aarch64": { url: "https://example.com/Formula_arm64.msi", signature: "sig" },
      "darwin-universal": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
    },
  };

  const assets = assetMap([
    "Formula.AppImage",
    "Formula_arm64.AppImage",
    "Formula_x64.msi",
    "Formula_arm64.msi",
  ]);
  assert.throws(
    () => validateLatestJson(manifest, "0.1.0", assets),
    (err) => err instanceof ActionableError && err.message.includes("but that asset is not present"),
  );
});

test("validateLatestJson accepts missing inline signature when a sibling .sig asset exists", () => {
  const manifest = {
    version: "0.1.0",
    platforms: {
      "linux-x86_64": { url: "https://example.com/Formula.AppImage", signature: "" },
      "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "" },
      "windows-x86_64": { url: "https://example.com/Formula_x64.msi", signature: "" },
      "windows-aarch64": { url: "https://example.com/Formula_arm64.msi", signature: "" },
      "darwin-universal": { url: "https://example.com/Formula.app.tar.gz", signature: "" },
    },
  };

  const assets = assetMap([
    "Formula.AppImage",
    "Formula.AppImage.sig",
    "Formula_arm64.AppImage",
    "Formula_arm64.AppImage.sig",
    "Formula_x64.msi",
    "Formula_x64.msi.sig",
    "Formula_arm64.msi",
    "Formula_arm64.msi.sig",
    "Formula.app.tar.gz",
    "Formula.app.tar.gz.sig",
  ]);

  assert.doesNotThrow(() => validateLatestJson(manifest, "0.1.0", assets));
});

test("validateLatestJson fails when both inline signature and sibling .sig are missing", () => {
  const manifest = {
    version: "0.1.0",
    platforms: {
      "linux-x86_64": { url: "https://example.com/Formula.AppImage", signature: "" },
      "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "sig" },
      "windows-x86_64": { url: "https://example.com/Formula_x64.msi", signature: "sig" },
      "windows-aarch64": { url: "https://example.com/Formula_arm64.msi", signature: "sig" },
      "darwin-universal": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
    },
  };

  const assets = assetMap([
    "Formula.AppImage",
    "Formula_arm64.AppImage",
    "Formula_x64.msi",
    "Formula_arm64.msi",
    "Formula.app.tar.gz",
  ]);
  assert.throws(
    () => validateLatestJson(manifest, "0.1.0", assets),
    (err) => err instanceof ActionableError && err.message.includes("missing a non-empty \"signature\""),
  );
});

test("verifyUpdaterManifestSignature verifies latest.json.sig against latest.json with minisign pubkey", () => {
  const { publicKey, privateKey } = crypto.generateKeyPairSync("ed25519");
  const latestJsonBytes = Buffer.from(JSON.stringify({ version: "0.1.0", platforms: {} }) + "\n", "utf8");
  const signature = crypto.sign(null, latestJsonBytes, privateKey);
  const latestSigText = `${signature.toString("base64")}\n`;

  // Build a Tauri/minisign pubkey string: base64(minisign text block).
  const spki = /** @type {Buffer} */ (publicKey.export({ format: "der", type: "spki" }));
  const rawPubkey = spki.subarray(spki.length - 32);
  const keyIdLe = Buffer.from("0102030405060708", "hex");
  const keyIdHex = Buffer.from(keyIdLe).reverse().toString("hex").toUpperCase();
  const pubPayload = Buffer.concat([Buffer.from([0x45, 0x64]), keyIdLe, rawPubkey]);
  const pubkeyText = `untrusted comment: minisign public key: ${keyIdHex}\n${pubPayload.toString("base64")}\n`;
  const pubkeyBase64 = Buffer.from(pubkeyText, "utf8").toString("base64");

  assert.doesNotThrow(() => verifyUpdaterManifestSignature(latestJsonBytes, latestSigText, pubkeyBase64));
});

test("verifyUpdaterManifestSignature fails on key id mismatch (minisign payload)", () => {
  const { publicKey, privateKey } = crypto.generateKeyPairSync("ed25519");
  const latestJsonBytes = Buffer.from(JSON.stringify({ version: "0.1.0", platforms: {} }) + "\n", "utf8");
  const signature = crypto.sign(null, latestJsonBytes, privateKey);

  const spki = /** @type {Buffer} */ (publicKey.export({ format: "der", type: "spki" }));
  const rawPubkey = spki.subarray(spki.length - 32);

  const pubKeyIdLe = Buffer.from("0102030405060708", "hex");
  const pubKeyIdHex = Buffer.from(pubKeyIdLe).reverse().toString("hex").toUpperCase();
  const pubPayload = Buffer.concat([Buffer.from([0x45, 0x64]), pubKeyIdLe, rawPubkey]);
  const pubkeyText = `untrusted comment: minisign public key: ${pubKeyIdHex}\n${pubPayload.toString("base64")}\n`;
  const pubkeyBase64 = Buffer.from(pubkeyText, "utf8").toString("base64");

  const sigKeyIdLe = Buffer.from("1111111111111111", "hex");
  const minisignSigPayload = Buffer.concat([Buffer.from([0x45, 0x64]), sigKeyIdLe, Buffer.from(signature)]);
  const latestSigText = `${minisignSigPayload.toString("base64")}\n`;

  assert.throws(
    () => verifyUpdaterManifestSignature(latestJsonBytes, latestSigText, pubkeyBase64),
    (err) => err instanceof ActionableError && /key id mismatch/i.test(err.message),
  );
});
