import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const sourceScriptPath = path.join(repoRoot, "scripts", "ci", "check-rust-toolchain-pins.sh");
const scriptContents = readFileSync(sourceScriptPath, "utf8");

const bashProbe = spawnSync("bash", ["--version"], { encoding: "utf8" });
const hasBash = !bashProbe.error && bashProbe.status === 0;

const gitProbe = spawnSync("git", ["--version"], { encoding: "utf8" });
const hasGit = !gitProbe.error && gitProbe.status === 0;

const canRun = hasBash && hasGit;

/**
 * Runs the rust toolchain pin guard in a temporary git repo.
 * @param {Record<string, string>} files
 */
function run(files) {
  const tmpdir = mkdtempSync(path.join(os.tmpdir(), "formula-rust-toolchain-pins-"));
  try {
    let proc = spawnSync("git", ["init"], { cwd: tmpdir, encoding: "utf8" });
    assert.equal(proc.status, 0, proc.stderr);

    mkdirSync(path.join(tmpdir, ".github", "workflows"), { recursive: true });
    mkdirSync(path.join(tmpdir, "scripts", "ci"), { recursive: true });

    const tmpScriptPath = path.join(tmpdir, "scripts", "ci", "check-rust-toolchain-pins.sh");
    writeFileSync(tmpScriptPath, scriptContents, "utf8");

    for (const [name, content] of Object.entries(files)) {
      const filePath = path.join(tmpdir, name);
      mkdirSync(path.dirname(filePath), { recursive: true });
      writeFileSync(filePath, `${content}\n`, "utf8");
      proc = spawnSync("git", ["add", name], { cwd: tmpdir, encoding: "utf8" });
      assert.equal(proc.status, 0, proc.stderr);
    }

    return spawnSync("bash", [tmpScriptPath], { cwd: tmpdir, encoding: "utf8" });
  } finally {
    rmSync(tmpdir, { recursive: true, force: true });
  }
}

test("passes when workflows match rust-toolchain.toml", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: dtolnay/rust-toolchain@v1
        with:
          toolchain: 1.92.0
`,
  });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /Rust toolchain pins match/i);
});

test("passes when run script only mentions cargo in a comment", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - name: Script mentions cargo in a comment only
        run: |
          # cargo test --locked
          echo ok
`,
  });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /Rust toolchain pins match/i);
});

test("passes when a non-run block scalar mentions cargo (restore-keys)", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: actions/cache@v4
        with:
          restore-keys: |2
            cargo-\${{ runner.os }}-
      - run: echo ok
`,
  });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /Rust toolchain pins match/i);
});

test("fails when cargo is invoked inside a folded run block scalar without installing the toolchain", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - name: Run cargo in a folded scalar
        run: >-
          cargo test --locked
`,
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /before installing the pinned toolchain/i);
});

test("fails when cargo is invoked inside an indented run block scalar without installing the toolchain", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - name: Run cargo in an indented scalar
        run: |2
          cargo test --locked
`,
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /before installing the pinned toolchain/i);
});

test("passes when a github-script block mentions cargo in a string", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: actions/github-script@v7
        with:
          script: |
            console.log("cargo test --locked");
`,
  });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /Rust toolchain pins match/i);
});

test("fails when workflow toolchain differs from rust-toolchain.toml", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: dtolnay/rust-toolchain@v1
        with:
          toolchain: 1.91.0
`,
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /pin mismatch/i);
});

test("fails when dtolnay/rust-toolchain step is missing toolchain input", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: dtolnay/rust-toolchain@v1
        with:
          targets: wasm32-unknown-unknown
`,
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /missing.*toolchain/i);
});

test("fails when workflow uses cargo but does not install pinned toolchain", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - name: Run cargo without installing toolchain
        run: cargo test --locked
`,
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /before installing the pinned toolchain/i);
});

test("fails when workflow uses cargo in a quoted run line without installing pinned toolchain", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - name: Run cargo with YAML quoting
        run: "cargo test --locked"
`,
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /before installing the pinned toolchain/i);
});

test("passes when workflow installs toolchain before running cargo in a quoted run line", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: dtolnay/rust-toolchain@v1
        with:
          toolchain: 1.92.0
      - name: Run cargo with YAML quoting
        run: "cargo test --locked"
`,
  });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /Rust toolchain pins match/i);
});

test("fails when workflow runs the Cargo.lock reproducibility script without installing pinned toolchain", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  preflight:
    runs-on: ubuntu-24.04
    steps:
      - name: Check Cargo.lock
        run: bash scripts/ci/check-cargo-lock-reproducible.sh
`,
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /before installing the pinned toolchain/i);
});

test("passes when workflow installs toolchain before running Cargo.lock reproducibility script", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  preflight:
    runs-on: ubuntu-24.04
    steps:
      - uses: dtolnay/rust-toolchain@v1
        with:
          toolchain: 1.92.0
      - name: Check Cargo.lock
        run: bash scripts/ci/check-cargo-lock-reproducible.sh
`,
  });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /Rust toolchain pins match/i);
});

test("fails when workflow runs check-tauri-permissions.mjs without installing pinned toolchain", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  validate:
    runs-on: ubuntu-24.04
    steps:
      - run: node scripts/check-tauri-permissions.mjs
`,
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /before installing the pinned toolchain/i);
});

test("passes when workflow installs toolchain before running check-tauri-permissions.mjs", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  validate:
    runs-on: ubuntu-24.04
    steps:
      - uses: dtolnay/rust-toolchain@v1
        with:
          toolchain: 1.92.0
      - run: node scripts/check-tauri-permissions.mjs
`,
  });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /Rust toolchain pins match/i);
});

test("fails when workflow runs build-wasm.mjs without installing pinned toolchain", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - run: node packages/engine/scripts/build-wasm.mjs
`,
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /before installing the pinned toolchain/i);
});

test("passes when workflow installs toolchain before running build-wasm.mjs", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: dtolnay/rust-toolchain@v1
        with:
          toolchain: 1.92.0
      - run: node packages/engine/scripts/build-wasm.mjs
`,
  });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /Rust toolchain pins match/i);
});

test("fails when workflow runs build-formula-wasm-node.mjs without installing pinned toolchain", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - run: node scripts/build-formula-wasm-node.mjs
`,
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /before installing the pinned toolchain/i);
});

test("passes when workflow installs toolchain before running build-formula-wasm-node.mjs", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: dtolnay/rust-toolchain@v1
        with:
          toolchain: 1.92.0
      - run: node scripts/build-formula-wasm-node.mjs
`,
  });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /Rust toolchain pins match/i);
});

test("fails when workflow runs excel-oracle compat_gate.py without installing pinned toolchain", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  excel:
    runs-on: ubuntu-24.04
    steps:
      - run: python tools/excel-oracle/compat_gate.py
`,
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /before installing the pinned toolchain/i);
});

test("passes when workflow installs toolchain before running excel-oracle compat_gate.py", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  excel:
    runs-on: ubuntu-24.04
    steps:
      - uses: dtolnay/rust-toolchain@v1
        with:
          toolchain: 1.92.0
      - run: python tools/excel-oracle/compat_gate.py
`,
  });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /Rust toolchain pins match/i);
});

test("fails when workflow runs corpus triage without installing pinned toolchain", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  corpus:
    runs-on: ubuntu-24.04
    steps:
      - run: python -m tools.corpus.triage --help
`,
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /before installing the pinned toolchain/i);
});

test("passes when workflow installs toolchain before running corpus triage", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  corpus:
    runs-on: ubuntu-24.04
    steps:
      - uses: dtolnay/rust-toolchain@v1
        with:
          toolchain: 1.92.0
      - run: python -m tools.corpus.triage --help
`,
  });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /Rust toolchain pins match/i);
});

test("fails when workflow runs scripts/security/ci.sh without installing pinned toolchain", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  security:
    runs-on: ubuntu-24.04
    steps:
      - run: ./scripts/security/ci.sh
`,
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /before installing the pinned toolchain/i);
});

test("passes when workflow installs toolchain before running scripts/security/ci.sh", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  security:
    runs-on: ubuntu-24.04
    steps:
      - uses: dtolnay/rust-toolchain@v1
        with:
          toolchain: 1.92.0
      - run: ./scripts/security/ci.sh
`,
  });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /Rust toolchain pins match/i);
});

test("fails when workflow runs desktop_size_report.py without installing pinned toolchain", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  desktop:
    runs-on: ubuntu-24.04
    steps:
      - run: python3 scripts/desktop_size_report.py --dist apps/desktop/dist --json-out desktop-size-report.json
`,
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /before installing the pinned toolchain/i);
});

test("passes when workflow installs toolchain before running desktop_size_report.py", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  desktop:
    runs-on: ubuntu-24.04
    steps:
      - uses: dtolnay/rust-toolchain@v1
        with:
          toolchain: 1.92.0
      - run: python3 scripts/desktop_size_report.py --dist apps/desktop/dist --json-out desktop-size-report.json
`,
  });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /Rust toolchain pins match/i);
});

test("fails when workflow runs desktop_binary_size_report.py without installing pinned toolchain", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  desktop:
    runs-on: ubuntu-24.04
    steps:
      - run: python3 scripts/desktop_binary_size_report.py --no-build --json-out desktop-binary-size-report.json
`,
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /before installing the pinned toolchain/i);
});

test("passes when workflow installs toolchain before running desktop_binary_size_report.py", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  desktop:
    runs-on: ubuntu-24.04
    steps:
      - uses: dtolnay/rust-toolchain@v1
        with:
          toolchain: 1.92.0
      - run: python3 scripts/desktop_binary_size_report.py --no-build --json-out desktop-binary-size-report.json
`,
  });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /Rust toolchain pins match/i);
});

test("fails when workflow installs Rust in one job but uses cargo in another", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  setup:
    runs-on: ubuntu-24.04
    steps:
      - uses: dtolnay/rust-toolchain@v1
        with:
          toolchain: 1.92.0
      - run: echo ok
  build:
    runs-on: ubuntu-24.04
    steps:
      - run: cargo test --locked
`,
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /before installing the pinned toolchain/i);
});

test("fails when a job installs Rust after running cargo", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - run: cargo test --locked
      - uses: dtolnay/rust-toolchain@v1
        with:
          toolchain: 1.92.0
`,
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /before installing the pinned toolchain/i);
});

test("passes when a job installs Rust before running cargo", { skip: !canRun }, () => {
  const proc = run({
    "rust-toolchain.toml": `
[toolchain]
channel = "1.92.0"
`,
    ".github/workflows/ci.yml": `
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: dtolnay/rust-toolchain@v1
        with:
          toolchain: 1.92.0
      - run: cargo test --locked
`,
  });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /Rust toolchain pins match/i);
});
