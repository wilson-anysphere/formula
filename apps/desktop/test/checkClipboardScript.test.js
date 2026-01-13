import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("pnpm check:clipboard script builds via cargo_agent and targets the desktop tauri package", () => {
  const scriptPath = path.join(__dirname, "..", "scripts", "check-clipboard.mjs");
  const src = fs.readFileSync(scriptPath, "utf8");

  // Must target the workspace package (not the lib crate name).
  assert.match(
    src,
    /const\s+DESKTOP_TAURI_PACKAGE\s*=\s*["']formula-desktop-tauri["']\s*;/,
    "expected clipboard smoke-check script to set DESKTOP_TAURI_PACKAGE to 'formula-desktop-tauri'",
  );

  // Must build via the repo's agent-safe wrapper by default.
  assert.match(
    src,
    /run\(\s*["']bash["']\s*,\s*\[\s*["']scripts\/cargo_agent\.sh["']/,
    "expected clipboard smoke-check script to invoke scripts/cargo_agent.sh",
  );

  // Must pass the correct build selectors.
  for (const token of [
    "--features",
    "desktop",
    "--bin",
    "formula-desktop",
    "--release",
    "--clipboard-smoke-check",
  ]) {
    assert.ok(
      src.includes(token),
      `expected clipboard smoke-check script to include build/run arg: ${token}`,
    );
  }

  // Should ensure the expected release binary exists after the build.
  assert.match(
    src,
    /existsSync\(\s*binary\s*\)/,
    "expected clipboard smoke-check script to verify the built binary exists",
  );

  // If Cargo output is redirected via CARGO_TARGET_DIR, the script should still find the binary.
  assert.match(
    src,
    /process\.env\.CARGO_TARGET_DIR/,
    "expected clipboard smoke-check script to respect CARGO_TARGET_DIR when locating the built binary",
  );

  // Must build the desktop frontend (Vite).
  assert.match(
    src,
    /run\(\s*["']pnpm["']\s*,\s*\[[^\]]*["']build["']/,
    "expected clipboard smoke-check script to run `pnpm build` for the desktop frontend",
  );

  // On Linux we need to support headless CI, so ensure the script is aware of xvfb-run-safe.sh.
  assert.match(
    src,
    /process\.platform\s*===\s*["']linux["'][^]*xvfb-run-safe\.sh/,
    "expected clipboard smoke-check script to reference scripts/xvfb-run-safe.sh for Linux/headless runs",
  );

  // Windows compatibility: allow falling back to `cargo` when `bash` isn't available.
  assert.match(
    src,
    /process\.platform\s*!==\s*["']win32["'][^]*run\(\s*["']bash["']/,
    "expected clipboard smoke-check script to use bash+cargo_agent.sh by default on non-Windows platforms",
  );
  assert.match(
    src,
    /err\.code\s*===\s*["']ENOENT["'][^]*run\(\s*["']cargo["']\s*,\s*cargoArgs/,
    "expected clipboard smoke-check script to fall back to `cargo` on Windows when bash is unavailable",
  );
});

