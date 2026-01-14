import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const scriptPath = path.join(repoRoot, "scripts", "ci", "linux-package-install-smoke.sh");

function run(args) {
  const proc = spawnSync("bash", [scriptPath, ...args], {
    cwd: repoRoot,
    encoding: "utf8",
  });
  if (proc.error) throw proc.error;
  return proc;
}

test("linux-package-install-smoke: --help prints expected usage + env vars", () => {
  const proc = run(["--help"]);
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /linux-package-install-smoke\.sh \[deb\|rpm\|all\]/);
  assert.match(proc.stdout, /CARGO_TARGET_DIR/);
  assert.match(proc.stdout, /DOCKER_PLATFORM/);
  assert.match(proc.stdout, /FORMULA_DEB_SMOKE_IMAGE/);
  assert.match(proc.stdout, /FORMULA_RPM_SMOKE_IMAGE/);
});

test("linux-package-install-smoke: invalid arg exits with usage (status 2)", () => {
  const proc = run(["not-a-real-subcommand"]);
  assert.equal(proc.status, 2);
  assert.match(proc.stderr, /usage:/i);
});

test("linux-package-install-smoke: uses token-based Exec= matching (avoid substring grep)", () => {
  const script = readFileSync(scriptPath, "utf8");
  assert.match(script, /target the expected executable/i);
  assert.doesNotMatch(script, /Exec referencing/i);
  // Historical implementation used `grep -rlE` to find a desktop file by substring.
  assert.doesNotMatch(script, /grep -rlE/);
});

test("linux-package-install-smoke: uses identifier-based Parquet shared-mime-info path", () => {
  const script = readFileSync(scriptPath, "utf8");
  assert.match(script, /FORMULA_TAURI_IDENTIFIER/);
  assert.ok(
    script.includes('mime_xml="/usr/share/mime/packages/${ident}.xml"'),
    "Expected linux-package-install-smoke.sh to build the shared-mime-info XML path from the Tauri identifier (mime_xml=/usr/share/mime/packages/${ident}.xml).",
  );
  // Avoid hardcoding the current identifier filename so we don't regress if identifier changes.
  assert.doesNotMatch(script, /app\\.formula\\.desktop\\.xml/);
});
