import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("pnpm check:coi script builds via cargo_agent and targets the desktop tauri package", () => {
  const scriptPath = path.join(__dirname, "..", "scripts", "check-cross-origin-isolation.mjs");
  const src = fs.readFileSync(scriptPath, "utf8");

  // Must target the workspace package (not the lib crate name).
  assert.match(
    src,
    /formula-desktop-tauri/,
    "expected COI smoke-check script to reference the Cargo package 'formula-desktop-tauri'",
  );

  // Must build via the repo's agent-safe wrapper by default.
  assert.match(
    src,
    /scripts\/cargo_agent\.sh/,
    "expected COI smoke-check script to invoke scripts/cargo_agent.sh",
  );

  // Must pass the correct build selectors.
  for (const token of [
    "--features",
    "desktop",
    "--bin",
    "formula-desktop",
    "--release",
    "--cross-origin-isolation-check",
  ]) {
    assert.ok(
      src.includes(token),
      `expected COI smoke-check script to include build/run arg: ${token}`,
    );
  }

  // Should ensure the expected release binary exists after the build.
  assert.match(
    src,
    /existsSync\(\s*binary\s*\)/,
    "expected COI smoke-check script to verify the built binary exists",
  );

  // On Linux we need to support headless CI, so ensure the script is aware of xvfb-run-safe.sh.
  assert.match(
    src,
    /xvfb-run-safe\.sh/,
    "expected COI smoke-check script to reference scripts/xvfb-run-safe.sh for Linux/headless runs",
  );
});

