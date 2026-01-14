import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { stripComments } from "../apps/desktop/test/sourceTextUtils.js";

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");

test("check-desktop-version supports overriding tauri.conf.json path via FORMULA_TAURI_CONF_PATH", () => {
  const script = stripComments(readFileSync(join(repoRoot, "scripts", "check-desktop-version.mjs"), "utf8"));
  assert.match(script, /FORMULA_TAURI_CONF_PATH/);
});
