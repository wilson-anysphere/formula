import { spawn } from "node:child_process";
import { readdir } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

/**
 * Node's `--test` runner started detecting TypeScript test files (`*.test.ts`) once
 * TypeScript stripping support landed. `apps/desktop` uses Vitest for `.test.ts`
 * suites, while `test:node` is intended to run only `node:test` suites written in
 * JavaScript.
 *
 * Run an explicit list of test files so `pnpm -C apps/desktop test:node` stays
 * stable across Node.js versions.
 */

const testDir = path.normalize(fileURLToPath(new URL("../test/", import.meta.url)));
const clipboardTestDir = path.normalize(fileURLToPath(new URL("../src/clipboard/__tests__/", import.meta.url)));

/** @type {string[]} */
const files = [];
await collectTests(testDir, files);
await collectTests(clipboardTestDir, files);
files.sort((a, b) => a.localeCompare(b));

if (files.length === 0) {
  console.log("No node:test files found.");
  process.exit(0);
}

// Keep node:test parallelism conservative; some suites start background services.
const nodeArgs = ["--no-warnings", "--test-concurrency=2", "--test", ...files];
const child = spawn(process.execPath, nodeArgs, { stdio: "inherit" });
child.on("exit", (code, signal) => {
  if (signal) {
    console.error(`node:test exited with signal ${signal}`);
    process.exit(1);
  }
  process.exit(code ?? 1);
});

/**
 * @param {string} dir
 * @param {string[]} out
 * @returns {Promise<void>}
 */
async function collectTests(dir, out) {
  let entries;
  try {
    entries = await readdir(dir, { withFileTypes: true });
  } catch {
    return;
  }

  for (const entry of entries) {
    if (!entry.isFile()) continue;
    if (!entry.name.endsWith(".test.js")) continue;
    out.push(path.join(dir, entry.name));
  }
}
