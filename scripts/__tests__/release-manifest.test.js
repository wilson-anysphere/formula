import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

// Note: this file uses the `.test.js` extension so it is picked up by `pnpm test:node`
// (scripts/run-node-tests.mjs collects `*.test.js` suites).

import crypto from "node:crypto";

import {
  mergeTauriUpdaterManifests,
  normalizeVersion as normalizeVersionForMerge,
} from "../merge-tauri-updater-manifests.mjs";
import { verifyTauriManifestSignature } from "../tauri-updater-manifest.mjs";
import {
  ActionableError,
  filenameFromUrl,
  findPlatformsObject,
  isPrimaryBundleAssetName,
  normalizeVersion as normalizeVersionForVerify,
  validateLatestJson,
} from "../verify-desktop-release-assets.mjs";
import { validatePlatformEntries } from "../ci/validate-updater-manifest.mjs";
import {
  ed25519PrivateKeyFromSeed,
  ed25519PublicKeyFromRaw,
  parseTauriUpdaterPubkey,
  parseTauriUpdaterSignature,
} from "../ci/tauri-minisign.mjs";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const fixtureDir = path.join(repoRoot, "scripts", "__fixtures__", "tauri-updater");

async function readJsonFixture(name) {
  const full = path.join(fixtureDir, name);
  const text = await readFile(full, "utf8");
  return JSON.parse(text);
}

async function readTextFixture(name) {
  const full = path.join(fixtureDir, name);
  return await readFile(full, "utf8");
}

test("normalizeVersion strips refs/tags + leading v", () => {
  assert.equal(normalizeVersionForVerify("v0.1.0"), "0.1.0");
  assert.equal(normalizeVersionForVerify("0.1.0"), "0.1.0");

  // Keep behavior consistent across scripts.
  assert.equal(normalizeVersionForMerge("v0.1.0"), "0.1.0");
  assert.equal(normalizeVersionForMerge("refs/tags/v0.1.0"), "0.1.0");
});

async function runCheckDesktopVersion({ tag, tauriVersion, cargoVersion }) {
  const tmpDir = await mkdtemp(path.join(os.tmpdir(), "formula-desktop-version-"));
  try {
    const scriptsDir = path.join(tmpDir, "scripts");
    const srcTauriDir = path.join(tmpDir, "apps", "desktop", "src-tauri");

    await mkdir(scriptsDir, { recursive: true });
    await mkdir(srcTauriDir, { recursive: true });

    const scriptText = await readFile(path.join(repoRoot, "scripts", "check-desktop-version.mjs"), "utf8");
    await writeFile(path.join(scriptsDir, "check-desktop-version.mjs"), scriptText, "utf8");

    await writeFile(
      path.join(srcTauriDir, "tauri.conf.json"),
      `${JSON.stringify({ version: tauriVersion }, null, 2)}\n`,
      "utf8",
    );
    await writeFile(
      path.join(srcTauriDir, "Cargo.toml"),
      [
        "[package]",
        'name = "formula-desktop-tauri"',
        `version = "${cargoVersion}"`,
        'edition = "2021"',
        "",
      ].join("\n"),
      "utf8",
    );

    return spawnSync(
      process.execPath,
      [path.join(scriptsDir, "check-desktop-version.mjs"), tag],
      { cwd: tmpDir, encoding: "utf8" },
    );
  } finally {
    await rm(tmpDir, { recursive: true, force: true });
  }
}

test("check-desktop-version.mjs passes when tag matches tauri.conf.json and Cargo.toml", async () => {
  const proc = await runCheckDesktopVersion({
    tag: "v1.2.3",
    tauriVersion: "1.2.3",
    cargoVersion: "1.2.3",
  });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /passed/i);
});

test("check-desktop-version.mjs accepts refs/tags/* inputs", async () => {
  const proc = await runCheckDesktopVersion({
    tag: "refs/tags/v1.2.3",
    tauriVersion: "1.2.3",
    cargoVersion: "1.2.3",
  });
  assert.equal(proc.status, 0, proc.stderr);
});

test("check-desktop-version.mjs lists only tauri.conf.json when it is the mismatch", async () => {
  const proc = await runCheckDesktopVersion({
    tag: "v1.2.3",
    tauriVersion: "1.2.2",
    cargoVersion: "1.2.3",
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /Desktop version mismatch detected/i);
  assert.match(proc.stderr, /Bump "version" in apps\/desktop\/src-tauri\/tauri\.conf\.json/i);
  assert.doesNotMatch(proc.stderr, /Bump \[package\]\.version in apps\/desktop\/src-tauri\/Cargo\.toml/i);
});

test("check-desktop-version.mjs lists only Cargo.toml when it is the mismatch", async () => {
  const proc = await runCheckDesktopVersion({
    tag: "v1.2.3",
    tauriVersion: "1.2.3",
    cargoVersion: "1.2.2",
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /Desktop version mismatch detected/i);
  assert.match(proc.stderr, /Bump \[package\]\.version in apps\/desktop\/src-tauri\/Cargo\.toml/i);
  assert.doesNotMatch(proc.stderr, /Bump "version" in apps\/desktop\/src-tauri\/tauri\.conf\.json/i);
});

test("verify-desktop-release-assets helpers: filenameFromUrl extracts decoded filenames", () => {
  assert.equal(
    filenameFromUrl(
      "https://example.invalid/download/Formula%20Desktop_0.1.0_x64_en-US.msi?foo=bar",
    ),
    "Formula Desktop_0.1.0_x64_en-US.msi",
  );
  assert.equal(filenameFromUrl("https://example.invalid/download/latest.json#fragment"), "latest.json");
  assert.equal(filenameFromUrl("not-a-url/Formula_0.1.0_x64_en-US.msi?x=y"), "Formula_0.1.0_x64_en-US.msi");
});

test("verify-desktop-release-assets helpers: findPlatformsObject locates nested platforms maps", () => {
  const nested = {
    foo: { bar: { platforms: { "windows-x86_64": { url: "https://example.invalid/x.msi" } } } },
  };
  const found = findPlatformsObject(nested);
  assert.ok(found, "expected to find platforms object");
  assert.deepEqual(found?.path, ["foo", "bar", "platforms"]);
  assert.ok("windows-x86_64" in found.platforms);
});

test("verify-desktop-release-assets helpers: isPrimaryBundleAssetName matches expected bundle extensions", () => {
  assert.equal(isPrimaryBundleAssetName("Formula_0.1.0_x64_en-US.msi"), true);
  assert.equal(isPrimaryBundleAssetName("Formula_0.1.0_x64_en-US.exe"), true);
  assert.equal(isPrimaryBundleAssetName("Formula_0.1.0_universal.dmg"), true);
  assert.equal(isPrimaryBundleAssetName("Formula_0.1.0.AppImage"), true);
  assert.equal(isPrimaryBundleAssetName("Formula_0.1.0_amd64.deb"), true);
  assert.equal(isPrimaryBundleAssetName("latest.json"), false);
  assert.equal(isPrimaryBundleAssetName("latest.json.sig"), false);
});

test("mergeTauriUpdaterManifests unions platform keys and normalizes version", async () => {
  const a = await readJsonFixture("latest.partial.a.json");
  const b = await readJsonFixture("latest.partial.b.json");

  const merged = mergeTauriUpdaterManifests([a, b]);
  assert.equal(merged.version, "0.1.0");

  const keys = Object.keys(merged.platforms).sort();
  assert.deepEqual(keys, [
    "darwin-aarch64",
    "darwin-universal",
    "darwin-x86_64",
    "linux-aarch64",
    "linux-x86_64",
    "windows-aarch64",
    "windows-x86_64",
  ]);
});

test("mergeTauriUpdaterManifests fails on conflicting duplicate platform entries", async () => {
  const a = await readJsonFixture("latest.conflict.a.json");
  const b = await readJsonFixture("latest.conflict.b.json");

  assert.throws(() => mergeTauriUpdaterManifests([a, b]), /Conflicting platform entry/);
});

test("mergeTauriUpdaterManifests fails when versions do not match (after normalization)", async () => {
  const a = await readJsonFixture("latest.partial.a.json");
  const b = await readJsonFixture("latest.wrong-version.json");

  assert.throws(() => mergeTauriUpdaterManifests([a, b]), /version mismatch/i);
});

test("mergeTauriUpdaterManifests rejects malformed inputs (empty list / missing fields)", () => {
  assert.throws(() => mergeTauriUpdaterManifests([]), /Expected one or more/i);
  assert.throws(() => mergeTauriUpdaterManifests([/** @type {any} */ ({})]), /version/i);
  assert.throws(
    () => mergeTauriUpdaterManifests([/** @type {any} */ ({ version: "0.1.0" })]),
    /platforms/i,
  );
  assert.throws(
    () => mergeTauriUpdaterManifests([/** @type {any} */ ({ version: "0.1.0", platforms: {} }), null]),
    /Manifest\[1\] is not an object/i,
  );
});

test("verify-desktop-release-assets: validateLatestJson fails when a platform URL references a missing release asset", async () => {
  const manifest = await readJsonFixture("latest.multi-platform.json");
  const assetsByName = assetsMapFromManifest(manifest);

  const missingAsset = filenameFromUrl(manifest.platforms["windows-aarch64"].url);
  assert.ok(missingAsset);
  assetsByName.delete(missingAsset);

  assert.throws(
    () => validateLatestJson(manifest, "0.1.0", assetsByName),
    (err) => {
      assert.ok(err instanceof ActionableError);
      assert.ok(
        err.message.includes(missingAsset),
        `expected error message to mention missing asset ${missingAsset}, got:\n${err.message}`,
      );
      return true;
    },
  );
});

test("verify-desktop-release-assets: validateLatestJson requires per-platform signatures (inline or <asset>.sig)", async () => {
  const manifest = await readJsonFixture("latest.multi-platform.json");
  const assetsByName = assetsMapFromManifest(manifest);

  // Remove an inline signature without providing a corresponding `.sig` asset.
  const target = "linux-aarch64";
  const assetName = filenameFromUrl(manifest.platforms[target].url);
  assert.ok(assetName);
  delete manifest.platforms[target].signature;

  assert.throws(
    () => validateLatestJson(manifest, "0.1.0", assetsByName),
    (err) => {
      assert.ok(err instanceof ActionableError);
      assert.match(err.message, new RegExp(`${assetName}\\.sig`));
      return true;
    },
  );
});

test("verify-desktop-release-assets: validateLatestJson accepts signature-only-as-asset (<bundle>.sig) when inline signatures are absent", async () => {
  const manifest = await readJsonFixture("latest.multi-platform.json");
  const assetsByName = assetsMapFromManifest(manifest);

  for (const entry of Object.values(manifest.platforms)) {
    if (!entry || typeof entry !== "object") continue;
    delete entry.signature;
  }

  // Provide `.sig` assets for every referenced bundle.
  for (const name of [...assetsByName.keys()]) {
    assetsByName.set(`${name}.sig`, { name: `${name}.sig` });
  }

  // Should not throw.
  validateLatestJson(manifest, "0.1.0", assetsByName);
});

test("verify-desktop-release-assets: validateLatestJson accepts a complete multi-platform manifest fixture", async () => {
  const manifest = await readJsonFixture("latest.multi-platform.json");
  const assetsByName = assetsMapFromManifest(manifest);

  validateLatestJson(manifest, "0.1.0", assetsByName);
});

test("verify-desktop-release-assets: validateLatestJson rejects raw Windows .exe updater entries by default", async () => {
  const manifest = await readJsonFixture("latest.multi-platform.json");
  manifest.platforms["windows-x86_64"].url = "https://example.invalid/download/v0.1.0/formula-desktop_0.1.0_x64_en-US.exe";
  manifest.platforms["windows-aarch64"].url =
    "https://example.invalid/download/v0.1.0/formula-desktop_0.1.0_arm64_en-US.exe";
  const assetsByName = assetsMapFromManifest(manifest);

  assert.throws(
    () => validateLatestJson(manifest, "0.1.0", assetsByName),
    (err) => {
      assert.ok(err instanceof ActionableError);
      assert.match(err.message, /allow-windows-exe/i);
      return true;
    },
  );
});

test("verify-desktop-release-assets: validateLatestJson allows raw Windows .exe updater entries when --allow-windows-exe is set", async () => {
  const manifest = await readJsonFixture("latest.multi-platform.json");
  manifest.platforms["windows-x86_64"].url = "https://example.invalid/download/v0.1.0/formula-desktop_0.1.0_x64_en-US.exe";
  manifest.platforms["windows-aarch64"].url =
    "https://example.invalid/download/v0.1.0/formula-desktop_0.1.0_arm64_en-US.exe";
  const assetsByName = assetsMapFromManifest(manifest);

  validateLatestJson(manifest, "0.1.0", assetsByName, { allowWindowsExe: true });
});

test("verify-desktop-release-assets: validateLatestJson enforces per-OS updater artifact types (no DMG/DEB)", async () => {
  const manifest = await readJsonFixture("latest.wrong-artifacts.json");
  const assetsByName = assetsMapFromManifest(manifest);

  assert.throws(
    () => validateLatestJson(manifest, "0.1.0", assetsByName),
    (err) => {
      assert.ok(err instanceof ActionableError);
      assert.ok(
        err.message.includes(".dmg") || err.message.includes(".deb"),
        `expected error message to mention .dmg or .deb, got:\n${err.message}`,
      );
      return true;
    },
  );
});

test("verify-desktop-release-assets: validateLatestJson fails on missing required platforms", async () => {
  const manifest = await readJsonFixture("latest.missing-platforms.json");
  const assetsByName = assetsMapFromManifest(manifest);

  assert.throws(
    () => validateLatestJson(manifest, "0.1.0", assetsByName),
    (err) => {
      assert.ok(err instanceof ActionableError);
      assert.match(err.message, /missing required key/i);
      return true;
    },
  );
});

test("verify-desktop-release-assets: validateLatestJson fails on wrong version", async () => {
  const manifest = await readJsonFixture("latest.wrong-version.json");
  const assetsByName = assetsMapFromManifest(manifest);

  assert.throws(
    () => validateLatestJson(manifest, "0.1.0", assetsByName),
    (err) => {
      assert.ok(err instanceof ActionableError);
      assert.match(err.message, /version mismatch/i);
      return true;
    },
  );
});

test("validate-updater-manifest: validatePlatformEntries enforces per-OS artifact rules (no DMG/DEB)", async () => {
  const manifest = await readJsonFixture("latest.wrong-artifacts.json");

  const assetNames = new Set(getAssetNamesFromPlatforms(manifest.platforms));
  const { errors } = validatePlatformEntries({ platforms: manifest.platforms, assetNames });

  assert.ok(errors.length > 0, "expected validation errors");
  assert.match(errors.join("\n"), /(dmg|\\.deb)/i);
});

test("tauri-minisign: verifies latest.json.sig with a test Ed25519 keypair", async () => {
  const manifestText = await readTextFixture("latest.multi-platform.json");
  const manifestBytes = Buffer.from(manifestText, "utf8");
  const signatureText = (await readTextFixture("latest.multi-platform.json.sig")).trim();
  const keypair = await readJsonFixture("test-keypair.json");

  // Deterministic signing from a fixed seed (RFC 8032 test vector).
  const seed = Buffer.from(keypair.privateKeySeedBase64, "base64");
  const privateKey = ed25519PrivateKeyFromSeed(seed);
  const computedSig = crypto.sign(null, manifestBytes, privateKey);
  assert.equal(computedSig.toString("base64"), signatureText);

  const parsedSig = parseTauriUpdaterSignature(signatureText, "latest.multi-platform.json.sig");
  assert.equal(parsedSig.signatureBytes.length, 64);

  const parsedPub = parseTauriUpdaterPubkey(keypair.publicKeyBase64);
  const publicKey = ed25519PublicKeyFromRaw(parsedPub.publicKeyBytes);

  assert.equal(crypto.verify(null, manifestBytes, publicKey, parsedSig.signatureBytes), true);

  // Wrong public key should fail.
  const wrongPub = ed25519PublicKeyFromRaw(Buffer.alloc(32));
  assert.equal(crypto.verify(null, manifestBytes, wrongPub, parsedSig.signatureBytes), false);
});

test("tauri-updater-manifest: verifyTauriManifestSignature supports minisign key/signature formats", async () => {
  const manifestText = await readTextFixture("latest.multi-platform.json");
  const keypair = await readJsonFixture("test-keypair.json");

  const seed = Buffer.from(keypair.privateKeySeedBase64, "base64");
  const privateKey = ed25519PrivateKeyFromSeed(seed);
  const signatureBytes = crypto.sign(null, Buffer.from(manifestText, "utf8"), privateKey);

  const rawPubkey = Buffer.from(keypair.publicKeyBase64, "base64");

  // Synthetic minisign key id for the test (little-endian bytes in payload; printed key id is big-endian hex).
  const keyIdLe = Buffer.from([1, 2, 3, 4, 5, 6, 7, 8]);
  const keyIdHex = Buffer.from(keyIdLe).reverse().toString("hex").toUpperCase();

  const pubPayload = Buffer.concat([Buffer.from([0x45, 0x64]), keyIdLe, rawPubkey]); // "Ed" + keyId + pubkey
  const pubkeyText = `untrusted comment: minisign public key: ${keyIdHex}\n${pubPayload.toString("base64")}\n`;
  const pubkeyBase64 = Buffer.from(pubkeyText, "utf8").toString("base64");

  // 1) Raw base64 signature (64 bytes)
  assert.equal(
    verifyTauriManifestSignature(manifestText, signatureBytes.toString("base64"), pubkeyBase64),
    true,
  );

  // 2) Minisign payload (base64 of 74 bytes)
  const sigPayload = Buffer.concat([Buffer.from([0x45, 0x64]), keyIdLe, signatureBytes]); // "Ed" + keyId + sig
  assert.equal(
    verifyTauriManifestSignature(manifestText, sigPayload.toString("base64"), pubkeyBase64),
    true,
  );

  // 3) Minisign signature file (comment + payload line)
  const sigFileText = `untrusted comment: minisign signature: ${keyIdHex}\n${sigPayload.toString("base64")}\n`;
  assert.equal(verifyTauriManifestSignature(manifestText, sigFileText, pubkeyBase64), true);
});

test("tauri-updater-manifest: verifyTauriManifestSignature returns false for a wrong signature", async () => {
  const manifestText = await readTextFixture("latest.multi-platform.json");
  const keypair = await readJsonFixture("test-keypair.json");

  const seed = Buffer.from(keypair.privateKeySeedBase64, "base64");
  const privateKey = ed25519PrivateKeyFromSeed(seed);
  const signatureBytes = crypto.sign(null, Buffer.from(manifestText, "utf8"), privateKey);

  // Flip a byte to invalidate.
  signatureBytes[0] ^= 0xff;

  const ok = verifyTauriManifestSignature(manifestText, signatureBytes.toString("base64"), keypair.publicKeyBase64);
  assert.equal(ok, false);
});

test("tauri-updater-manifest: verifyTauriManifestSignature fails fast on minisign key id mismatch", async () => {
  const manifestText = await readTextFixture("latest.multi-platform.json");
  const keypair = await readJsonFixture("test-keypair.json");

  const seed = Buffer.from(keypair.privateKeySeedBase64, "base64");
  const privateKey = ed25519PrivateKeyFromSeed(seed);
  const signatureBytes = crypto.sign(null, Buffer.from(manifestText, "utf8"), privateKey);

  const rawPubkey = Buffer.from(keypair.publicKeyBase64, "base64");
  const pubKeyIdLe = Buffer.from([1, 2, 3, 4, 5, 6, 7, 8]);
  const pubKeyIdHex = Buffer.from(pubKeyIdLe).reverse().toString("hex").toUpperCase();
  const pubPayload = Buffer.concat([Buffer.from([0x45, 0x64]), pubKeyIdLe, rawPubkey]);
  const pubkeyText = `untrusted comment: minisign public key: ${pubKeyIdHex}\n${pubPayload.toString("base64")}\n`;
  const pubkeyBase64 = Buffer.from(pubkeyText, "utf8").toString("base64");

  // Signature uses a different key id.
  const sigKeyIdLe = Buffer.from([9, 9, 9, 9, 9, 9, 9, 9]);
  const sigPayload = Buffer.concat([Buffer.from([0x45, 0x64]), sigKeyIdLe, signatureBytes]);

  assert.throws(
    () => verifyTauriManifestSignature(manifestText, sigPayload.toString("base64"), pubkeyBase64),
    /key id mismatch/i,
  );
});

test("verify-updater-manifest-signature.mjs verifies latest.json.sig against a test tauri.conf.json pubkey", async () => {
  const keypair = await readJsonFixture("test-keypair.json");

  const tmpDir = await mkdtemp(path.join(os.tmpdir(), "formula-updater-sigtest-"));
  const tmpConfigPath = path.join(tmpDir, "tauri.conf.json");
  await writeFile(
    tmpConfigPath,
    `${JSON.stringify({ plugins: { updater: { pubkey: keypair.publicKeyBase64 } } }, null, 2)}\n`,
    "utf8",
  );

  const latestJsonPath = path.join(fixtureDir, "latest.multi-platform.json");
  const latestSigPath = path.join(fixtureDir, "latest.multi-platform.json.sig");

  const child = spawnSync(
    process.execPath,
    [path.join(repoRoot, "scripts", "ci", "verify-updater-manifest-signature.mjs"), latestJsonPath, latestSigPath],
    {
      cwd: repoRoot,
      env: { ...process.env, FORMULA_TAURI_CONF_PATH: tmpConfigPath },
      encoding: "utf8",
    },
  );

  assert.equal(
    child.status,
    0,
    `verify-updater-manifest-signature.mjs failed (exit ${child.status})\nstdout:\n${child.stdout}\nstderr:\n${child.stderr}`,
  );
  assert.match(child.stdout, /signature OK/i);
});

test("merge-tauri-updater-manifests.mjs CLI merges manifests and writes output JSON", async () => {
  const tmpDir = await mkdtemp(path.join(os.tmpdir(), "formula-updater-merge-"));
  const outPath = path.join(tmpDir, "merged.json");

  const child = spawnSync(
    process.execPath,
    [
      path.join(repoRoot, "scripts", "merge-tauri-updater-manifests.mjs"),
      "--out",
      outPath,
      path.join("scripts", "__fixtures__", "tauri-updater", "latest.partial.a.json"),
      path.join("scripts", "__fixtures__", "tauri-updater", "latest.partial.b.json"),
    ],
    { cwd: repoRoot, encoding: "utf8" },
  );

  assert.equal(
    child.status,
    0,
    `merge-tauri-updater-manifests.mjs failed (exit ${child.status})\nstdout:\n${child.stdout}\nstderr:\n${child.stderr}`,
  );

  const mergedText = await readFile(outPath, "utf8");
  const merged = JSON.parse(mergedText);
  assert.equal(merged.version, "0.1.0");
  assert.ok(merged.platforms && typeof merged.platforms === "object");
  assert.ok("windows-x86_64" in merged.platforms);
  assert.ok("windows-aarch64" in merged.platforms);
  assert.ok("linux-x86_64" in merged.platforms);
  assert.ok("linux-aarch64" in merged.platforms);
});

function assetsMapFromManifest(manifest) {
  const names = getAssetNamesFromPlatforms(manifest?.platforms);
  return new Map(names.map((name) => [name, { name }]));
}

function getAssetNamesFromPlatforms(platforms) {
  if (!platforms || typeof platforms !== "object") return [];
  return Object.values(platforms)
    .map((entry) => (entry && typeof entry === "object" ? entry.url : null))
    .filter((url) => typeof url === "string")
    .map((url) => new URL(url).pathname.split("/").filter(Boolean).pop())
    .filter((name) => typeof name === "string" && name.length > 0);
}
