import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const scriptPath = path.join(repoRoot, "scripts", "check-updater-config.mjs");

function runWithConfig(config) {
  const tmpRoot = path.join(repoRoot, ".tmp");
  mkdirSync(tmpRoot, { recursive: true });
  const dir = mkdtempSync(path.join(tmpRoot, "check-updater-config-"));
  const configPath = path.join(dir, "tauri.conf.json");
  writeFileSync(configPath, `${JSON.stringify(config)}\n`, "utf8");

  const proc = spawnSync(process.execPath, [scriptPath], {
    encoding: "utf8",
    cwd: repoRoot,
    env: {
      ...process.env,
      FORMULA_TAURI_CONF_PATH: configPath,
    },
  });
  if (proc.error) throw proc.error;
  rmSync(dir, { recursive: true, force: true });
  return proc;
}

function fakeMinisignPublicKey() {
  // minisign public key payload: "Ed" + keyId (8 bytes) + pubkey (32 bytes)
  const header = Buffer.from([0x45, 0x64]); // "Ed"
  const keyId = Buffer.alloc(8, 0x11);
  const pub = Buffer.alloc(32, 0x22);
  const binary = Buffer.concat([header, keyId, pub]);
  const payload = binary.toString("base64").replace(/=+$/, "");
  const keyIdHex = Buffer.from(keyId).reverse().toString("hex").toUpperCase();
  const keyFile = `untrusted comment: minisign public key: ${keyIdHex}\n${payload}\n`;
  return Buffer.from(keyFile, "utf8").toString("base64");
}

test("passes when updater is active and pubkey looks like a minisign public key", () => {
  const config = {
    plugins: {
      updater: {
        active: true,
        dialog: false,
        endpoints: ["https://github.com/example-org/example-repo/releases/latest/download/latest.json"],
        pubkey: fakeMinisignPublicKey(),
      },
    },
  };
  const proc = runWithConfig(config);
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /preflight passed/i);
});

test("fails when updater is active but pubkey is not a minisign public key", () => {
  const config = {
    plugins: {
      updater: {
        active: true,
        dialog: false,
        endpoints: ["https://github.com/example-org/example-repo/releases/latest/download/latest.json"],
        pubkey: "not-a-key",
      },
    },
  };
  const proc = runWithConfig(config);
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /plugins\.updater\.pubkey/i);
  assert.match(proc.stderr, /minisign public key/i);
});
