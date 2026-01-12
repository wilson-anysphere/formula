import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("desktop docs use the correct Cargo package name for desktop shell builds", () => {
  const repoRoot = path.join(__dirname, "..", "..", "..");

  const desktopShellDoc = fs.readFileSync(path.join(repoRoot, "docs", "11-desktop-shell.md"), "utf8");
  assert.match(
    desktopShellDoc,
    /\bcargo_agent\.sh\s+test\s+-p\s+desktop\b/,
    "expected docs/11-desktop-shell.md to use -p desktop in cargo_agent test example",
  );
  assert.match(
    desktopShellDoc,
    /\bcargo_agent\.sh\s+check\s+-p\s+desktop\b[^\\n]*--features\s+desktop\b/,
    "expected docs/11-desktop-shell.md to use -p desktop in cargo_agent check example",
  );
  assert.ok(
    !/\bcargo_agent\.sh\s+test\s+-p\s+formula-desktop-tauri\b/.test(desktopShellDoc),
    "docs/11-desktop-shell.md should not suggest `-p formula-desktop-tauri` (use -p desktop via cargo_agent.sh remapping)",
  );
  assert.ok(
    !/\bcargo_agent\.sh\s+check\s+-p\s+formula-desktop-tauri\b/.test(desktopShellDoc),
    "docs/11-desktop-shell.md should not suggest `-p formula-desktop-tauri` (use -p desktop via cargo_agent.sh remapping)",
  );

  const platformDoc = fs.readFileSync(path.join(repoRoot, "instructions", "platform.md"), "utf8");
  assert.match(
    platformDoc,
    /\bcargo_agent\.sh\s+check\s+-p\s+desktop\b[^\\n]*--features\s+desktop\b[^\\n]*--lib\b/,
    "expected instructions/platform.md to use -p desktop in cargo_agent check example",
  );
  assert.ok(
    !/\bcargo_agent\.sh\s+check\s+-p\s+formula-desktop-tauri\b/.test(platformDoc),
    "instructions/platform.md should not suggest `-p formula-desktop-tauri` (use -p desktop via cargo_agent.sh remapping)",
  );

  const desktopReadme = fs.readFileSync(path.join(repoRoot, "apps", "desktop", "README.md"), "utf8");
  assert.match(
    desktopReadme,
    /\bcargo_agent\.sh\s+build\s+-p\s+formula-desktop-tauri\b[^\\n]*--features\s+desktop\b[^\\n]*--bin\s+formula-desktop\b[^\\n]*--release\b/,
    "expected apps/desktop/README.md to use -p formula-desktop-tauri in manual build instructions",
  );
  assert.ok(
    !/\bcargo_agent\.sh\s+build\s+-p\s+desktop\b/.test(desktopReadme),
    "apps/desktop/README.md should not suggest `-p desktop` (workspace package is formula-desktop-tauri)",
  );
});
