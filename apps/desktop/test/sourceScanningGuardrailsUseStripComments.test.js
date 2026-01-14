import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { stripComments } from "./sourceTextUtils.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function collectNodeTestFiles(dir) {
  return fs
    .readdirSync(dir, { withFileTypes: true })
    .filter((entry) => entry.isFile() && entry.name.endsWith(".test.js"))
    .map((entry) => path.join(dir, entry.name))
    .sort((a, b) => a.localeCompare(b));
}

test("node:test source-scanning guardrails strip comments when scanning main.ts", () => {
  const files = collectNodeTestFiles(__dirname);
  /** @type {string[]} */
  const offenders = [];

  // Heuristic: any node:test suite that reads `src/main.ts` as text should use
  // `stripComments()` so commented-out wiring can't satisfy or fail assertions.
  //
  // (We intentionally do not enforce this for other source files because some tests
  // may rely on comment markers as slicing anchors.)
  const referencesMainTsRe = /\bmain\.ts\b/;
  const readsFileRe = /\breadFileSync\s*\(|\breadFile\s*\(/;
  const callsStripCommentsRe = /\bstripComments\s*\(/;

  for (const file of files) {
    const raw = fs.readFileSync(file, "utf8");
    // Strip comments before applying heuristics so commented-out code (including a commented-out
    // `stripComments(...)` call) cannot satisfy this guardrail.
    const text = stripComments(raw);
    if (!referencesMainTsRe.test(text)) continue;
    if (!readsFileRe.test(text)) continue;
    if (!callsStripCommentsRe.test(text)) {
      offenders.push(path.relative(__dirname, file));
    }
  }

  assert.deepEqual(
    offenders,
    [],
    `The following node:test suites read src/main.ts as text but do not use stripComments():\n${offenders
      .map((f) => `- ${f}`)
      .join("\n")}`,
  );
});
