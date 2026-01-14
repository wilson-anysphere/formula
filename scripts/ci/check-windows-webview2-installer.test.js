import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { stripPythonComments } from "../../apps/desktop/test/sourceTextUtils.js";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const scriptPath = path.join(repoRoot, "scripts", "ci", "check-windows-webview2-installer.py");

/**
 * Best-effort detection for a runnable Python executable (for CI + local dev).
 */
function detectPythonExecutable() {
  /** @type {string[]} */
  const candidates = [];
  const envPython = process.env.PYTHON?.trim() ?? "";
  if (envPython) candidates.push(envPython);

  // Prefer python3 on Unix, python on Windows.
  candidates.push(process.platform === "win32" ? "python" : "python3");
  candidates.push("python3");
  candidates.push("python");

  for (const py of candidates) {
    const probe = spawnSync(py, ["--version"], { encoding: "utf8" });
    if (probe.error && probe.error.code === "ENOENT") continue;
    if (probe.status === 0) return py;
  }

  return null;
}

const pythonExe = detectPythonExecutable();
const hasPython = Boolean(pythonExe);

test("check-windows-webview2-installer bounds fallback src-tauri discovery (perf guardrail)", () => {
  const src = stripPythonComments(readFileSync(scriptPath, "utf8"));
  assert.ok(
    src.includes("max_depth = 8") && src.includes("if depth >= max_depth"),
    "Expected check-windows-webview2-installer.py to bound os.walk repo discovery with max_depth.",
  );
});

/**
 * @param {{ installerBytes: Buffer; webviewInstallMode?: string; cwd?: string }} opts
 */
function run(opts) {
  assert.ok(pythonExe, "python executable not found");
  const tmpdir = mkdtempSync(path.join(os.tmpdir(), "formula-webview2-installer-"));
  const targetDir = path.join(tmpdir, "target");
  const installerDir = path.join(targetDir, "release", "bundle", "nsis");
  mkdirSync(installerDir, { recursive: true });

  const installerPath = path.join(installerDir, "FormulaInstaller.exe");
  writeFileSync(installerPath, opts.installerBytes);

  const configPath = path.join(tmpdir, "tauri.conf.json");
  writeFileSync(
    configPath,
    JSON.stringify(
      { bundle: { windows: { webviewInstallMode: opts.webviewInstallMode ?? "downloadBootstrapper" } } },
      null,
      2,
    ),
    "utf8",
  );

  const proc = spawnSync(pythonExe, [scriptPath], {
    cwd: opts.cwd ?? repoRoot,
    encoding: "utf8",
    env: {
      ...process.env,
      // Point the verifier at our temp bundle output so the test does not depend on any
      // real build artifacts being present.
      CARGO_TARGET_DIR: targetDir,
      FORMULA_TAURI_CONF_PATH: configPath,
    },
  });

  rmSync(tmpdir, { recursive: true, force: true });
  if (proc.error) throw proc.error;
  return proc;
}

test("passes when installer contains WebView2 marker string", { skip: !hasPython }, () => {
  const proc = run({
    installerBytes: Buffer.from("...MicrosoftEdgeWebView2Setup.exe...", "utf8"),
  });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout + proc.stderr, /webview2-check: OK/i);
});

test("fails when installer contains no WebView2 markers", { skip: !hasPython }, () => {
  const proc = run({
    installerBytes: Buffer.from("no markers here", "utf8"),
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /no WebView2 installer markers found|does not appear to bundle\/reference/i);
});

test("fails when webviewInstallMode is set to skip (via FORMULA_TAURI_CONF_PATH)", { skip: !hasPython }, () => {
  const proc = run({
    installerBytes: Buffer.from("...MicrosoftEdgeWebView2Setup.exe...", "utf8"),
    webviewInstallMode: "skip",
  });
  assert.equal(proc.status, 2);
  assert.match(proc.stderr, /webviewInstallMode.*skip/i);
});

test("can be invoked from a non-repo-root cwd (repo root derived from script path)", { skip: !hasPython }, () => {
  const proc = run({
    installerBytes: Buffer.from("...MicrosoftEdgeWebView2Setup.exe...", "utf8"),
    cwd: path.join(repoRoot, "apps", "desktop"),
  });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout + proc.stderr, /webview2-check: OK/i);
});
