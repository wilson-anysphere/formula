import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { stripHashComments } from "../apps/desktop/test/sourceTextUtils.js";

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");

const scripts = [
  "validate-linux-appimage.sh",
  "validate-linux-deb.sh",
  "validate-linux-rpm.sh",
].map((name) => join(repoRoot, "scripts", name));

test("Linux bundle validators do not hardcode x-scheme-handler/formula", () => {
  for (const scriptPath of scripts) {
    const contents = stripHashComments(readFileSync(scriptPath, "utf8"));
    assert.doesNotMatch(
      contents,
      /x-scheme-handler\/formula/,
      `Expected ${scriptPath} to derive scheme handlers from tauri.conf.json, not hardcode formula`,
    );
  }
});
