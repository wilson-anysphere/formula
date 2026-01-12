import test from "node:test";
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

const SCRIPT = path.resolve("scripts/check-cursor-ai-policy.mjs");

async function writeFixtureFile(root, relativePath, contents) {
  const fullPath = path.join(root, relativePath);
  await fs.mkdir(path.dirname(fullPath), { recursive: true });
  await fs.writeFile(fullPath, contents, "utf8");
}

function runPolicy(rootDir) {
  return spawnSync(process.execPath, [SCRIPT, "--root", rootDir], {
    encoding: "utf8",
    cwd: path.resolve("."),
  });
}

test("cursor AI policy guard passes on a clean fixture", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-pass-"));
  try {
    await writeFixtureFile(tmpRoot, "packages/example/src/index.js", 'export const answer = 42;\n');
    await writeFixtureFile(tmpRoot, "apps/example/src/main.ts", "export function main() { return 1; }\n");

    const proc = runPolicy(tmpRoot);
    assert.equal(proc.status, 0, proc.stderr || proc.stdout);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard fails when forbidden provider strings are present", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "packages/example/src/index.js", 'import OpenAI from "openai";\n');

    const proc = runPolicy(tmpRoot);
    assert.notEqual(proc.status, 0);
    assert.match(`${proc.stdout}\n${proc.stderr}`, /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard fails when OpenAI appears in non-test source files (even without imports)", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-openai-source-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "packages/example/src/index.js", 'const provider = "OpenAI";\n');

    const proc = runPolicy(tmpRoot);
    assert.notEqual(proc.status, 0);
    assert.match(`${proc.stdout}\n${proc.stderr}`, /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard fails when forbidden provider strings are present in root config files", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-root-config-fail-"));
  try {
    // Root-level config files (package.json, Cargo.toml, etc) are scanned because
    // adding forbidden dependencies there should fail fast.
    await writeFixtureFile(tmpRoot, "package.json", '{ "name": "example", "dependencies": { "openai": "0.0.0" } }\n');

    const proc = runPolicy(tmpRoot);
    assert.notEqual(proc.status, 0);
    assert.match(`${proc.stdout}\n${proc.stderr}`, /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans markdown readmes for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-readme-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "packages/example/README.md", "This should not mention OpenAI.\n");

    const proc = runPolicy(tmpRoot);
    assert.notEqual(proc.status, 0);
    assert.match(`${proc.stdout}\n${proc.stderr}`, /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans html files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-html-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "apps/example/index.html", "<!-- OpenAI -->\n");

    const proc = runPolicy(tmpRoot);
    assert.notEqual(proc.status, 0);
    assert.match(`${proc.stdout}\n${proc.stderr}`, /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans css files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-css-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "apps/example/src/styles.css", "/* OpenAI */\n");

    const proc = runPolicy(tmpRoot);
    assert.notEqual(proc.status, 0);
    assert.match(`${proc.stdout}\n${proc.stderr}`, /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans snapshot files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-snap-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "packages/example/src/__tests__/__snapshots__/thing.snap", "OpenAI\n");

    const proc = runPolicy(tmpRoot);
    assert.notEqual(proc.status, 0);
    assert.match(`${proc.stdout}\n${proc.stderr}`, /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans SQL files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-sql-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "services/example/migrations/0001_init.sql", "-- OpenAI\n");

    const proc = runPolicy(tmpRoot);
    assert.notEqual(proc.status, 0);
    assert.match(`${proc.stdout}\n${proc.stderr}`, /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans .txt files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-txt-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "packages/example/NOTICE.txt", "OpenAI\n");

    const proc = runPolicy(tmpRoot);
    assert.notEqual(proc.status, 0);
    assert.match(`${proc.stdout}\n${proc.stderr}`, /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans PowerShell scripts for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-ps1-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "tools/example/run.ps1", "Write-Host OpenAI\n");

    const proc = runPolicy(tmpRoot);
    assert.notEqual(proc.status, 0);
    assert.match(`${proc.stdout}\n${proc.stderr}`, /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans XML files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-xml-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "crates/example/tests/fixtures/example.xml", "<!-- OpenAI -->\n");

    const proc = runPolicy(tmpRoot);
    assert.notEqual(proc.status, 0);
    assert.match(`${proc.stdout}\n${proc.stderr}`, /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans TSV files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-tsv-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "crates/example/src/data.tsv", "OpenAI\n");

    const proc = runPolicy(tmpRoot);
    assert.notEqual(proc.status, 0);
    assert.match(`${proc.stdout}\n${proc.stderr}`, /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans lockfiles (Cargo.lock) for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-cargo-lock-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "Cargo.lock", '[[package]]\nname = "openai"\nversion = "0.0.0"\n');

    const proc = runPolicy(tmpRoot);
    assert.notEqual(proc.status, 0);
    assert.match(`${proc.stdout}\n${proc.stderr}`, /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans pnpm-lock.yaml for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-pnpm-lock-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "pnpm-lock.yaml", "openai: 0.0.0\n");

    const proc = runPolicy(tmpRoot);
    assert.notEqual(proc.status, 0);
    assert.match(`${proc.stdout}\n${proc.stderr}`, /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans scripts/ directory by default", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-scripts-dir-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "scripts/example.mjs", 'import OpenAI from "openai";\n');

    const proc = runPolicy(tmpRoot);
    assert.notEqual(proc.status, 0);
    assert.match(`${proc.stdout}\n${proc.stderr}`, /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans shared/ directory by default", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-shared-dir-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "shared/example.js", 'const x = "OpenAI";\n');

    const proc = runPolicy(tmpRoot);
    assert.notEqual(proc.status, 0);
    assert.match(`${proc.stdout}\n${proc.stderr}`, /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans extensions/ directory by default", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-extensions-dir-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "extensions/example/dist/extension.js", 'const x = "OpenAI";\n');

    const proc = runPolicy(tmpRoot);
    assert.notEqual(proc.status, 0);
    assert.match(`${proc.stdout}\n${proc.stderr}`, /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans python/ directory by default", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-python-dir-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "python/example.py", 'provider = "OpenAI"\n');

    const proc = runPolicy(tmpRoot);
    assert.notEqual(proc.status, 0);
    assert.match(`${proc.stdout}\n${proc.stderr}`, /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans .github/workflows by default", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-github-dir-fail-"));
  try {
    await writeFixtureFile(tmpRoot, ".github/workflows/ci.yml", 'name: "OpenAI"\n');

    const proc = runPolicy(tmpRoot);
    assert.notEqual(proc.status, 0);
    assert.match(`${proc.stdout}\n${proc.stderr}`, /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans .cargo/ config by default", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-cargo-dir-fail-"));
  try {
    await writeFixtureFile(tmpRoot, ".cargo/config.toml", 'openai = "1"\n');

    const proc = runPolicy(tmpRoot);
    assert.notEqual(proc.status, 0);
    assert.match(`${proc.stdout}\n${proc.stderr}`, /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans patches/ directory by default", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-patches-dir-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "patches/example.patch", "OpenAI\n");

    const proc = runPolicy(tmpRoot);
    assert.notEqual(proc.status, 0);
    assert.match(`${proc.stdout}\n${proc.stderr}`, /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard fails when forbidden strings appear in unrelated unit tests", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-test-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "packages/example/src/something.test.js", 'const x = "anthropic";\n');

    const proc = runPolicy(tmpRoot);
    assert.notEqual(proc.status, 0);
    assert.match(`${proc.stdout}\n${proc.stderr}`, /anthropic/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard fails when OpenAI appears in unrelated unit tests (even without imports)", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-openai-test-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "packages/example/src/something.test.js", 'const x = "OpenAI";\n');

    const proc = runPolicy(tmpRoot);
    assert.notEqual(proc.status, 0);
    assert.match(`${proc.stdout}\n${proc.stderr}`, /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans repo-level tests/ directory for unit tests", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-tests-dir-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "tests/unrelated.test.js", 'const x = "OpenAI";\n');

    const proc = runPolicy(tmpRoot);
    assert.notEqual(proc.status, 0);
    assert.match(`${proc.stdout}\n${proc.stderr}`, /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans repo-level test/ directory for unit tests", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-test-dir-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "test/unrelated.test.js", 'const x = "OpenAI";\n');

    const proc = runPolicy(tmpRoot);
    assert.notEqual(proc.status, 0);
    assert.match(`${proc.stdout}\n${proc.stderr}`, /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard fails when OpenAI appears in *.vitest.* unit tests", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-openai-vitest-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "packages/example/src/something.vitest.ts", 'const x = "OpenAI";\n');

    const proc = runPolicy(tmpRoot);
    assert.notEqual(proc.status, 0);
    assert.match(`${proc.stdout}\n${proc.stderr}`, /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard allows forbidden strings in the guard's own tests", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-test-pass-"));
  try {
    // This file name intentionally matches the allowlist rule for policy guard tests.
    await writeFixtureFile(
      tmpRoot,
      "packages/example/src/cursor-ai-policy.test.js",
      'const fixtures = ["openai", "anthropic", "ollama", "formula:openaiApiKey"];\n',
    );

    const proc = runPolicy(tmpRoot);
    assert.equal(proc.status, 0, proc.stderr || proc.stdout);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});
