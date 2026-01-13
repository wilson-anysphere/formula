import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const scriptPath = path.join(repoRoot, "scripts", "ci", "prepare-tauri-signing-config.mjs");

/**
 * @param {string} filePath
 */
function readJson(filePath) {
  return JSON.parse(fs.readFileSync(filePath, "utf8"));
}

/**
 * @param {string} filePath
 */
function readText(filePath) {
  try {
    return fs.readFileSync(filePath, "utf8");
  } catch {
    return "";
  }
}

/**
 * @param {Record<string, string | undefined>} env
 * @param {any} config
 */
function runWithConfig(env, config) {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "formula-tauri-signing-"));
  const configPath = path.join(tmpDir, "tauri.conf.json");
  const envFile = path.join(tmpDir, "github_env");

  fs.writeFileSync(configPath, `${JSON.stringify(config, null, 2)}\n`, "utf8");

  const proc = spawnSync(process.execPath, [scriptPath], {
    encoding: "utf8",
    cwd: repoRoot,
    env: {
      ...process.env,
      ...env,
      FORMULA_TAURI_CONF_PATH: configPath,
      GITHUB_ENV: envFile,
    },
  });
  if (proc.error) throw proc.error;

  return {
    proc,
    configPath,
    envFile,
    tmpDir,
    config: readJson(configPath),
    githubEnv: readText(envFile),
  };
}

test("disables macOS code signing in the config when secrets are missing", () => {
  const minimalConfig = { bundle: { macOS: {}, windows: { certificateThumbprint: null } } };
  const { proc, config, githubEnv } = runWithConfig({ RUNNER_OS: "macOS" }, minimalConfig);
  assert.equal(proc.status, 0, proc.stderr);
  assert.equal(config.bundle.macOS.signingIdentity, null);
  assert.equal(githubEnv.trim(), "");
});

test("enables macOS code signing with an explicit identity when secrets are present", () => {
  const minimalConfig = { bundle: { macOS: {}, windows: { certificateThumbprint: null } } };
  const identity = "Developer ID Application: Example Co (TEAMID)";
  const { proc, config, githubEnv } = runWithConfig(
    {
      RUNNER_OS: "macOS",
      APPLE_CERTIFICATE: "base64p12",
      APPLE_CERTIFICATE_PASSWORD: "pw",
      APPLE_SIGNING_IDENTITY: identity,
    },
    minimalConfig,
  );
  assert.equal(proc.status, 0, proc.stderr);
  assert.equal(config.bundle.macOS.signingIdentity, identity);
  assert.match(githubEnv, /\bAPPLE_CERTIFICATE<</);
  assert.match(githubEnv, /\bAPPLE_SIGNING_IDENTITY<</);
});

test("disables macOS code signing when certificate secrets are present but signing identity is missing", () => {
  const minimalConfig = { bundle: { macOS: {}, windows: { certificateThumbprint: null } } };
  const { proc, config, githubEnv } = runWithConfig(
    {
      RUNNER_OS: "macOS",
      APPLE_CERTIFICATE: "base64p12",
      APPLE_CERTIFICATE_PASSWORD: "pw",
    },
    minimalConfig,
  );
  assert.equal(proc.status, 0, proc.stderr);
  assert.equal(config.bundle.macOS.signingIdentity, null);
  assert.equal(githubEnv.trim(), "");
});

test("exports notarization env vars only when all notarization credentials are present", () => {
  const minimalConfig = { bundle: { macOS: {}, windows: { certificateThumbprint: null } } };
  const identity = "Developer ID Application: Example Co (TEAMID)";
  const { proc, githubEnv } = runWithConfig(
    {
      RUNNER_OS: "macOS",
      APPLE_CERTIFICATE: "base64p12",
      APPLE_CERTIFICATE_PASSWORD: "pw",
      APPLE_SIGNING_IDENTITY: identity,
      APPLE_ID: "user@example.com",
      APPLE_PASSWORD: "app-specific-password",
      APPLE_TEAM_ID: "TEAMID",
    },
    minimalConfig,
  );
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(githubEnv, /\bAPPLE_ID<</);
  assert.match(githubEnv, /\bAPPLE_PASSWORD<</);
  assert.match(githubEnv, /\bAPPLE_TEAM_ID<</);
});

test("fails fast when FORMULA_REQUIRE_CODESIGN=1 and required secrets are missing", () => {
  const minimalConfig = { bundle: { macOS: {} } };
  const { proc, config } = runWithConfig(
    {
      RUNNER_OS: "macOS",
      FORMULA_REQUIRE_CODESIGN: "1",
      APPLE_CERTIFICATE: "",
      APPLE_CERTIFICATE_PASSWORD: "",
      APPLE_SIGNING_IDENTITY: "",
      APPLE_ID: "",
      APPLE_PASSWORD: "",
      APPLE_TEAM_ID: "",
    },
    minimalConfig,
  );
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /Code signing is required/i);
  // In enforcement mode we should fail before mutating the config.
  assert.equal(config.bundle.macOS.signingIdentity, undefined);
});

test("disables Windows code signing in the config when secrets are missing", () => {
  const cfg = { bundle: { windows: { certificateThumbprint: "ABCDEF" } } };
  const { proc, config, githubEnv } = runWithConfig({ RUNNER_OS: "Windows" }, cfg);
  assert.equal(proc.status, 0, proc.stderr);
  assert.equal(config.bundle.windows.certificateThumbprint, null);
  assert.equal(githubEnv.trim(), "");
});

test("exports Windows signing secrets when configured", () => {
  const cfg = { bundle: { windows: { certificateThumbprint: null } } };
  const { proc, config, githubEnv } = runWithConfig(
    {
      RUNNER_OS: "Windows",
      WINDOWS_CERTIFICATE: "base64pfx",
      WINDOWS_CERTIFICATE_PASSWORD: "pw",
    },
    cfg,
  );
  assert.equal(proc.status, 0, proc.stderr);
  assert.equal(config.bundle.windows.certificateThumbprint, null);
  assert.match(githubEnv, /\bWINDOWS_CERTIFICATE<</);
  assert.match(githubEnv, /\bWINDOWS_CERTIFICATE_PASSWORD<</);
});

test("fails fast when FORMULA_REQUIRE_CODESIGN=1 on Windows and required secrets are missing", () => {
  const cfg = { bundle: { windows: { certificateThumbprint: "ABCDEF" } } };
  const { proc, config } = runWithConfig(
    {
      RUNNER_OS: "Windows",
      FORMULA_REQUIRE_CODESIGN: "1",
      WINDOWS_CERTIFICATE: "",
      WINDOWS_CERTIFICATE_PASSWORD: "",
    },
    cfg,
  );
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /Code signing is required/i);
  // In enforcement mode we should fail before mutating the config.
  assert.equal(config.bundle.windows.certificateThumbprint, "ABCDEF");
});
