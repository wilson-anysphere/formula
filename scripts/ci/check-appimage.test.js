import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const scriptPath = path.join(repoRoot, "scripts", "ci", "check-appimage.sh");

const hasBash = (() => {
  if (process.platform === "win32") return false;
  const probe = spawnSync("bash", ["-lc", "exit 0"], { stdio: "ignore" });
  return probe.status === 0;
})();

test("check-appimage: --help prints usage and mentions FORMULA_TAURI_CONF_PATH", { skip: !hasBash }, () => {
  const proc = spawnSync("bash", [scriptPath, "--help"], { cwd: repoRoot, encoding: "utf8" });
  if (proc.error) throw proc.error;
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /check-appimage\.sh/i);
  assert.match(proc.stdout, /FORMULA_TAURI_CONF_PATH/);
});

test("check-appimage avoids unbounded find scans when directory args are provided (perf guardrail)", () => {
  const raw = readFileSync(scriptPath, "utf8");
  // Historical versions used `find "$arg" -type f -name '*.AppImage'` which can be
  // extremely slow when callers pass a Cargo `target/` directory.
  assert.ok(
    !raw.includes(`find "$arg" -type f -name '*.AppImage' -print0`),
    "Expected check-appimage.sh to avoid unbounded `find \"$arg\" -type f -name '*.AppImage'` scans.",
  );
});
