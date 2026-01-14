import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { chmodSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const runLimitedPath = path.join(repoRoot, "scripts", "run_limited.sh");
const xvfbSafePath = path.join(repoRoot, "scripts", "xvfb-run-safe.sh");

const bashProbe = spawnSync("bash", ["--version"], { encoding: "utf8" });
const hasBash = !bashProbe.error && bashProbe.status === 0;

function writeExecutable(dir, name, contents) {
  const outPath = path.join(dir, name);
  writeFileSync(outPath, contents, "utf8");
  chmodSync(outPath, 0o755);
  return outPath;
}

function makeFakeTools() {
  const tmpDir = mkdtempSync(path.join(os.tmpdir(), "formula-rustup-toolchain-cleared-"));
  const binDir = path.join(tmpDir, "bin");
  mkdirSync(binDir, { recursive: true });

  writeExecutable(
    binDir,
    "rustc",
    `#!/usr/bin/env bash
echo "RUSTUP_TOOLCHAIN=\${RUSTUP_TOOLCHAIN-}"
`,
  );

  writeExecutable(
    binDir,
    "xdpyinfo",
    `#!/usr/bin/env bash
exit 0
`,
  );

  return { tmpDir, binDir };
}

test("run_limited clears RUSTUP_TOOLCHAIN for rustc invocations", { skip: !hasBash }, () => {
  const { tmpDir, binDir } = makeFakeTools();
  try {
    const proc = spawnSync("bash", [runLimitedPath, "--no-as", "--no-cpu", "--", "rustc"], {
      cwd: repoRoot,
      encoding: "utf8",
      env: {
        ...process.env,
        RUSTUP_TOOLCHAIN: "stable",
        PATH: `${binDir}${path.delimiter}${process.env.PATH ?? ""}`,
      },
    });
    assert.equal(proc.status, 0, proc.stderr);
    assert.equal(proc.stdout.trim(), "RUSTUP_TOOLCHAIN=");
  } finally {
    rmSync(tmpDir, { recursive: true, force: true });
  }
});

test("xvfb-run-safe clears RUSTUP_TOOLCHAIN for rustc invocations", { skip: !hasBash }, () => {
  const { tmpDir, binDir } = makeFakeTools();
  try {
    const proc = spawnSync("bash", [xvfbSafePath, "rustc"], {
      cwd: repoRoot,
      encoding: "utf8",
      env: {
        ...process.env,
        // Ensure xvfb-run-safe doesn't try to launch Xvfb; it will see xdpyinfo succeed and
        // exec the wrapped command directly.
        DISPLAY: ":99",
        CI: "",
        RUSTUP_TOOLCHAIN: "stable",
        PATH: `${binDir}${path.delimiter}${process.env.PATH ?? ""}`,
      },
    });
    assert.equal(proc.status, 0, proc.stderr);
    assert.equal(proc.stdout.trim(), "RUSTUP_TOOLCHAIN=");
  } finally {
    rmSync(tmpDir, { recursive: true, force: true });
  }
});

