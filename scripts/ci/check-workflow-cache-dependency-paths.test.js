import assert from "node:assert/strict";
import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const workflowsDir = path.join(repoRoot, ".github", "workflows");

test("workflows avoid recursive pnpm-lock cache-dependency-path globs (perf guardrail)", () => {
  const bad = [];
  const entries = readdirSync(workflowsDir, { withFileTypes: true })
    .filter((ent) => ent.isFile() && (ent.name.endsWith(".yml") || ent.name.endsWith(".yaml")))
    .map((ent) => ent.name)
    .sort();

  // Guard against accidentally using `**/pnpm-lock.yaml` for setup-node cache discovery. The
  // actions glob implementation scans the entire repository, which can get slow once `target/` or
  // other build outputs exist.
  const badLine = /^\s*cache-dependency-path:\s*['"]?\*\*\/pnpm-lock\.yaml['"]?\s*(?:#.*)?$/;

  for (const name of entries) {
    const filePath = path.join(workflowsDir, name);
    const lines = readFileSync(filePath, "utf8").split(/\r?\n/);
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      if (badLine.test(line)) {
        bad.push(`${name}:${i + 1}:${line.trim()}`);
      }
    }
  }

  assert.deepEqual(
    bad,
    [],
    `Found recursive pnpm lock cache-dependency-path globs (use "pnpm-lock.yaml" instead):\n${bad.join("\n")}`,
  );
});

