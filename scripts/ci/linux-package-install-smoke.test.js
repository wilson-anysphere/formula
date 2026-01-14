import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { chmodSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";
import { tmpdir } from "node:os";

import { stripHashComments } from "../../apps/desktop/test/sourceTextUtils.js";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const scriptPath = path.join(repoRoot, "scripts", "ci", "linux-package-install-smoke.sh");
const scriptContents = stripHashComments(readFileSync(scriptPath, "utf8"));

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
  assert.match(proc.stdout, /FORMULA_TAURI_CONF_PATH/);
});

test("linux-package-install-smoke: invalid arg exits with usage (status 2)", () => {
  const proc = run(["not-a-real-subcommand"]);
  assert.equal(proc.status, 2);
  assert.match(proc.stderr, /usage:/i);
});

test("linux-package-install-smoke: uses token-based Exec= matching (avoid substring grep)", () => {
  assert.match(scriptContents, /target the expected executable/i);
  assert.doesNotMatch(scriptContents, /Exec referencing/i);
  // Historical implementation used `grep -rlE` to find a desktop file by substring.
  assert.doesNotMatch(scriptContents, /grep -rlE/);
});

test("linux-package-install-smoke: uses identifier-based Parquet shared-mime-info path", () => {
  assert.match(scriptContents, /FORMULA_TAURI_IDENTIFIER/);
  assert.ok(
    scriptContents.includes('mime_xml="/usr/share/mime/packages/${ident}.xml"'),
    "Expected linux-package-install-smoke.sh to build the shared-mime-info XML path from the Tauri identifier (mime_xml=/usr/share/mime/packages/${ident}.xml).",
  );
  // Avoid hardcoding the current identifier filename so we don't regress if identifier changes.
  assert.doesNotMatch(scriptContents, /app\\.formula\\.desktop\\.xml/);
});

test("linux-package-install-smoke: rejects path separators in the Tauri identifier", () => {
  assert.match(scriptContents, /invalid tauri identifier \(contains path separators\)/i);
});

test("linux-package-install-smoke: validates URL scheme handler(s) from tauri.conf.json (no hardcoded formula)", () => {
  assert.match(scriptContents, /FORMULA_DEEP_LINK_SCHEMES/);
  // Validate that the script doesn't hardcode only the `formula` scheme in its desktop entry checks.
  assert.doesNotMatch(scriptContents, /x-scheme-handler\/formula/);
});

test("linux-package-install-smoke: can print --help without a working python3/node (sed fallback for identifier)", () => {
  const tmp = mkdtempSync(path.join(tmpdir(), "formula-linux-smoke-help-"));
  try {
    const binDir = path.join(tmp, "bin");
    mkdirSync(binDir, { recursive: true });
    // Stub python3/node so the script cannot JSON-parse via either runtime; it should fall
    // back to `sed` for reading `identifier` (needed when Parquet association is configured).
    writeFileSync(
      path.join(binDir, "python3"),
      `#!/usr/bin/env bash\nset -euo pipefail\ncat >/dev/null || true\nexit 0\n`,
      "utf8",
    );
    chmodSync(path.join(binDir, "python3"), 0o755);
    writeFileSync(
      path.join(binDir, "node"),
      `#!/usr/bin/env bash\nset -euo pipefail\ncat >/dev/null || true\nexit 0\n`,
      "utf8",
    );
    chmodSync(path.join(binDir, "node"), 0o755);

    const proc = spawnSync("bash", [scriptPath, "--help"], {
      cwd: repoRoot,
      encoding: "utf8",
      env: {
        ...process.env,
        PATH: `${binDir}:${process.env.PATH}`,
      },
    });
    if (proc.error) throw proc.error;
    assert.equal(proc.status, 0, proc.stderr);
    assert.match(proc.stdout, /usage:/i);
  } finally {
    rmSync(tmp, { recursive: true, force: true });
  }
});
