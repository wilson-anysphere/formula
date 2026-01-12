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

