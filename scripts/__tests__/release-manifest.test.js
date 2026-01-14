import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFile } from "node:fs/promises";
import { mkdtemp, writeFile } from "node:fs/promises";
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
import {
  ActionableError,
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
