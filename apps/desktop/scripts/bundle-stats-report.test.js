import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "..", "..");
const scriptPath = path.join(repoRoot, "apps", "desktop", "scripts", "bundle-stats-report.mjs");

/**
 * @param {string} statsPath
 * @param {string[]} [args]
 */
function run(statsPath, args = []) {
  const proc = spawnSync(process.execPath, [scriptPath, "--file", statsPath, ...args], {
    cwd: repoRoot,
    encoding: "utf8",
    env: process.env,
  });
  if (proc.error) throw proc.error;
  return proc;
}

/**
 * @param {any} json
 */
function writeFixture(json) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "formula-desktop-bundle-stats-"));
  const statsPath = path.join(root, "bundle-stats.json");
  fs.writeFileSync(statsPath, JSON.stringify(json, null, 2), "utf8");
  return { root, statsPath };
}

test("prints top chunks, groups, and modules", () => {
  const { root, statsPath } = writeFixture({
    version: 2,
    tree: {
      name: "root",
      children: [
        {
          name: "assets/index-AAA.js",
          children: [
            { name: "apps/desktop/src/main.ts", uid: "u1" },
            { name: "node_modules/foo/index.js", uid: "u2" },
          ],
        },
        {
          name: "assets/vendor.js",
          children: [{ name: "node_modules/bar/index.js", uid: "u3" }],
        },
      ],
    },
    nodeParts: {
      u1: { renderedLength: 1_000, gzipLength: 400, brotliLength: 300 },
      u2: { renderedLength: 2_000, gzipLength: 800, brotliLength: 600 },
      u3: { renderedLength: 500, gzipLength: 200, brotliLength: 150 },
    },
  });

  const proc = run(statsPath, ["--", "--top", "5"]);
  fs.rmSync(root, { recursive: true, force: true });

  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /Top chunks/i);
  assert.match(proc.stdout, /assets\/index-AAA\.js/);
  assert.match(proc.stdout, /assets\/vendor\.js/);
  assert.match(proc.stdout, /Focus chunk: assets\/index-AAA\.js/);
  assert.match(proc.stdout, /Top dependency groups/i);
  assert.match(proc.stdout, /\bapps\/desktop\b/);
  assert.match(proc.stdout, /\bfoo\b/);
  assert.match(proc.stdout, /Top modules/i);
  assert.match(proc.stdout, /\bapps\/desktop\/src\/main\.ts\b/);
});

test("can focus a chunk by name substring", () => {
  const { root, statsPath } = writeFixture({
    version: 2,
    tree: {
      name: "root",
      children: [
        {
          name: "assets/index-AAA.js",
          children: [{ name: "apps/desktop/src/main.ts", uid: "u1" }],
        },
        {
          name: "assets/vendor.js",
          children: [{ name: "node_modules/bar/index.js", uid: "u2" }],
        },
      ],
    },
    nodeParts: {
      u1: { renderedLength: 1_000, gzipLength: 400, brotliLength: 300 },
      u2: { renderedLength: 500, gzipLength: 200, brotliLength: 150 },
    },
  });

  const proc = run(statsPath, ["--chunk", "vendor"]);
  fs.rmSync(root, { recursive: true, force: true });

  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /Focus chunk: assets\/vendor\.js/);
});

