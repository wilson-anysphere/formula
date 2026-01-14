import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { EXPECTED_PLATFORM_KEYS } from "../ci/validate-updater-manifest.mjs";
import { stripHtmlComments } from "../../apps/desktop/test/sourceTextUtils.js";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const docPath = path.join(repoRoot, "docs", "desktop-updater-target-mapping.md");

test("docs/desktop-updater-target-mapping.md stays in sync with required updater platform keys", async () => {
  const text = stripHtmlComments(await readFile(docPath, "utf8"));
  for (const key of EXPECTED_PLATFORM_KEYS) {
    assert.ok(
      text.includes(`\`${key}\``),
      `Expected ${path.relative(repoRoot, docPath)} to mention required platform key: ${key}`,
    );
  }
});
