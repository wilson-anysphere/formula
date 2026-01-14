import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { generateKeyPairSync, verify } from "node:crypto";
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
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

function makeSigningEnv(tmpDir) {
  const { privateKey, publicKey } = generateKeyPairSync("ed25519");

  const pkcs8Der = privateKey.export({ format: "der", type: "pkcs8" });
  const tauriPrivateKey = Buffer.from(pkcs8Der).toString("base64");

  const spkiDer = publicKey.export({ format: "der", type: "spki" });
  const spkiPrefix = Buffer.from("302a300506032b6570032100", "hex");
  assert.equal(spkiDer.length, spkiPrefix.length + 32);
  assert.ok(spkiDer.subarray(0, spkiPrefix.length).equals(spkiPrefix));
  const rawPubkey = spkiDer.subarray(spkiPrefix.length);
  const tauriUpdaterPubkey = Buffer.from(rawPubkey).toString("base64");

  const tauriConfPath = path.join(tmpDir, "tauri.conf.json");
  writeFileSync(
    tauriConfPath,
    JSON.stringify(
      {
        plugins: {
          updater: {
            pubkey: tauriUpdaterPubkey,
          },
        },
      },
      null,
      2,
    ),
  );

  return {
    publicKey,
    env: {
      TAURI_PRIVATE_KEY: tauriPrivateKey,
      TAURI_KEY_PASSWORD: "",
      FORMULA_TAURI_CONF_PATH: tauriConfPath,
    },
  };
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

  // Minimal multi-platform manifest (enough for publish-updater-manifest to validate required runtime targets).
  writeFileSync(
    path.join(manifestsDir, "linux.json"),
    JSON.stringify(
      {
        version: "0.1.0",
        platforms: {
          "darwin-x86_64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
          "darwin-aarch64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
          "windows-x86_64": { url: "https://example.com/Formula_x64.msi", signature: "sig" },
          "windows-aarch64": { url: "https://example.com/Formula_arm64.msi", signature: "sig" },
          "linux-x86_64": { url: "https://example.com/Formula_x86_64.AppImage", signature: "sig" },
          "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "sig" },
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

test("dry-run merges manifests deterministically and produces a verifiable signature", () => {
  const tmp = makeTempDir();
  const manifestsDir = path.join(tmp, "manifests");
  mkdirSync(manifestsDir, { recursive: true });

  const { publicKey, env } = makeSigningEnv(tmp);

  writeFileSync(
    path.join(manifestsDir, "a.json"),
    JSON.stringify(
      {
        version: "0.1.0",
        notes: "Hello\r\nWorld",
        pub_date: "2020-01-01T00:00:00Z",
        custom: { z: 1, a: 2 },
        platforms: {
          "linux-x86_64": {
            url: "https://example.com/Formula.AppImage",
            signature: "sig-linux",
            foo: "bar",
            nested: { b: 2, a: 1 },
          },
        },
      },
      null,
      2,
    ),
  );

  writeFileSync(
    path.join(manifestsDir, "b.json"),
    JSON.stringify(
      {
        version: "0.1.0",
        notes: "Ignored",
        pub_date: "2021-01-01T00:00:00Z",
        // Same value as in a.json (but with key order swapped) should be accepted.
        custom: { a: 2, z: 1 },
        // Field present only in this manifest should be included in the output.
        other: { b: 2, a: 1 },
        platforms: {
          "darwin-aarch64": {
            url: "https://example.com/Formula.app.tar.gz",
            signature: "sig-macos",
            size: 123,
          },
          // Same platform entry duplicated across manifests should be accepted if identical.
          "linux-x86_64": {
            url: "https://example.com/Formula.AppImage",
            signature: "sig-linux",
            foo: "bar",
            nested: { a: 1, b: 2 },
          },
        },
      },
      null,
      2,
    ),
  );

  const proc = run({
    cwd: tmp,
    args: ["--dry-run", "v0.1.0", manifestsDir],
    env,
  });

  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /dry run enabled/i);
  assert.match(proc.stdout, /notes\/pub_date sourced from .*a\.json/i);

  const latestJsonPath = path.join(tmp, "latest.json");
  const latestSigPath = path.join(tmp, "latest.json.sig");
  const latestJsonBytes = readFileSync(latestJsonPath);
  const latest = JSON.parse(latestJsonBytes.toString("utf8"));

  assert.equal(latest.version, "0.1.0");
  assert.equal(latest.notes, "Hello\nWorld");
  assert.equal(latest.pub_date, "2020-01-01T00:00:00Z");

  // Deterministic key ordering.
  assert.deepEqual(Object.keys(latest.platforms), ["darwin-aarch64", "linux-x86_64"]);
  assert.deepEqual(Object.keys(latest.custom), ["a", "z"]);
  assert.deepEqual(Object.keys(latest.other), ["a", "b"]);
  assert.deepEqual(Object.keys(latest.platforms["linux-x86_64"]), ["url", "signature", "foo", "nested"]);
  assert.deepEqual(Object.keys(latest.platforms["linux-x86_64"].nested), ["a", "b"]);

  const sig = Buffer.from(readFileSync(latestSigPath, "utf8").trim(), "base64");
  assert.equal(sig.length, 64);
  assert.equal(verify(null, latestJsonBytes, publicKey, sig), true);
});

test("fails when a runtime updater entry references an installer artifact (Linux .deb)", () => {
  const tmp = makeTempDir();
  const manifestsDir = path.join(tmp, "manifests");
  mkdirSync(manifestsDir, { recursive: true });
 
  const { env } = makeSigningEnv(tmp);
 
  writeFileSync(
    path.join(manifestsDir, "linux.json"),
    JSON.stringify(
      {
        version: "0.1.0",
        platforms: {
          "linux-x86_64": {
            url: "https://example.com/formula_0.1.0_amd64.deb",
            signature: "sig",
          },
        },
      },
      null,
      2,
    ),
  );
 
  const proc = run({
    cwd: tmp,
    args: ["--dry-run", "v0.1.0", manifestsDir],
    env,
  });
 
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /Linux updater artifact must be an AppImage bundle/i);
});

test("fails when required runtime platform keys are missing (non-dry-run)", () => {
  const tmp = makeTempDir();
  const manifestsDir = path.join(tmp, "manifests");
  mkdirSync(manifestsDir, { recursive: true });

  writeFileSync(
    path.join(manifestsDir, "partial.json"),
    JSON.stringify(
      {
        version: "0.1.0",
        platforms: {
          // Valid updater artifact type for Linux, but missing other required platform keys.
          "linux-x86_64": { url: "https://example.com/Formula_x86_64.AppImage", signature: "sig" },
        },
      },
      null,
      2,
    ),
  );

  const proc = run({
    cwd: tmp,
    args: ["v0.1.0", manifestsDir],
    env: {
      // Provide placeholders so the script gets past required env checks; it should fail before any network call.
      GITHUB_REPOSITORY: "owner/repo",
      GITHUB_TOKEN: "dummy",
    },
  });

  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /missing required platform/i);
  assert.match(proc.stderr, /darwin-x86_64/i);
  // If this ever appears, the script got past manifest validation and attempted a network call.
  assert.doesNotMatch(proc.stderr, /api\\.github\\.com/i);
});

test("fails when Windows runtime updater asset name is missing an arch token (non-dry-run)", () => {
  const tmp = makeTempDir();
  const manifestsDir = path.join(tmp, "manifests");
  mkdirSync(manifestsDir, { recursive: true });

  writeFileSync(
    path.join(manifestsDir, "bad-windows.json"),
    JSON.stringify(
      {
        version: "0.1.0",
        platforms: {
          "darwin-x86_64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
          "darwin-aarch64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
          // Missing x64 token in filename should be rejected by the stricter validator.
          "windows-x86_64": { url: "https://example.com/Formula.msi", signature: "sig" },
          "windows-aarch64": { url: "https://example.com/Formula_arm64.msi", signature: "sig" },
          "linux-x86_64": { url: "https://example.com/Formula_x86_64.AppImage", signature: "sig" },
          "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "sig" },
        },
      },
      null,
      2,
    ),
  );

  const proc = run({
    cwd: tmp,
    args: ["v0.1.0", manifestsDir],
    env: {
      // Provide placeholders so the script gets past required env checks if it reaches them.
      GITHUB_REPOSITORY: "owner/repo",
      GITHUB_TOKEN: "dummy",
    },
  });

  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /Invalid Windows updater asset naming/i);
  assert.match(proc.stderr, /arch token/i);
  // If this ever appears, the script got past manifest validation and attempted a network call.
  assert.doesNotMatch(proc.stderr, /api\\.github\\.com/i);
});

test("fails when required runtime targets collide on the same updater URL (non-dry-run)", () => {
  const tmp = makeTempDir();
  const manifestsDir = path.join(tmp, "manifests");
  mkdirSync(manifestsDir, { recursive: true });

  const sharedUrl = "https://example.com/Formula_x64.msi";

  writeFileSync(
    path.join(manifestsDir, "collision.json"),
    JSON.stringify(
      {
        version: "0.1.0",
        platforms: {
          "darwin-x86_64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
          "darwin-aarch64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
          // Collision: Windows x64 + arm64 point to the same updater URL/asset.
          "windows-x86_64": { url: sharedUrl, signature: "sig" },
          "windows-aarch64": { url: sharedUrl, signature: "sig" },
          "linux-x86_64": { url: "https://example.com/Formula_x86_64.AppImage", signature: "sig" },
          "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "sig" },
        },
      },
      null,
      2,
    ),
  );

  const proc = run({
    cwd: tmp,
    args: ["v0.1.0", manifestsDir],
    env: {
      // Provide placeholders so the script gets past required env checks if it reaches them.
      GITHUB_REPOSITORY: "owner/repo",
      GITHUB_TOKEN: "dummy",
    },
  });

  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /Duplicate updater URLs across required targets/i);
  // If this ever appears, the script got past manifest validation and attempted a network call.
  assert.doesNotMatch(proc.stderr, /api\\.github\\.com/i);
});

test("fails loudly on conflicting top-level manifest fields", () => {
  const tmp = makeTempDir();
  const manifestsDir = path.join(tmp, "manifests");
  mkdirSync(manifestsDir, { recursive: true });

  const { env } = makeSigningEnv(tmp);

  writeFileSync(
    path.join(manifestsDir, "a.json"),
    JSON.stringify(
      {
        version: "0.1.0",
        channel: "stable",
        platforms: {
          "linux-x86_64": { url: "https://example.com/a", signature: "sig" },
        },
      },
      null,
      2,
    ),
  );
  writeFileSync(
    path.join(manifestsDir, "b.json"),
    JSON.stringify(
      {
        version: "0.1.0",
        channel: "beta",
        platforms: {
          "darwin-aarch64": { url: "https://example.com/b", signature: "sig" },
        },
      },
      null,
      2,
    ),
  );

  const proc = run({
    cwd: tmp,
    args: ["--dry-run", "v0.1.0", manifestsDir],
    env,
  });

  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /Conflicting top-level manifest field/i);
  assert.match(proc.stderr, /channel/i);
});

test("fails loudly on conflicting platform entries", () => {
  const tmp = makeTempDir();
  const manifestsDir = path.join(tmp, "manifests");
  mkdirSync(manifestsDir, { recursive: true });

  const { env } = makeSigningEnv(tmp);

  writeFileSync(
    path.join(manifestsDir, "a.json"),
    JSON.stringify(
      {
        version: "0.1.0",
        platforms: {
          "linux-x86_64": { url: "https://example.com/a", signature: "sig" },
        },
      },
      null,
      2,
    ),
  );
  writeFileSync(
    path.join(manifestsDir, "b.json"),
    JSON.stringify(
      {
        version: "0.1.0",
        platforms: {
          "linux-x86_64": { url: "https://example.com/b", signature: "sig" },
        },
      },
      null,
      2,
    ),
  );

  const proc = run({
    cwd: tmp,
    args: ["--dry-run", "v0.1.0", manifestsDir],
    env,
  });

  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /Conflicting platform entry/i);
  assert.match(proc.stderr, /linux-x86_64/i);
});
