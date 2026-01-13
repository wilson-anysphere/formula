import test from "node:test";
import assert from "node:assert/strict";
import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

test("legacy formula-bar name-box implementation is removed", () => {
  // The name box UI is implemented in `apps/desktop/src/formula-bar/FormulaBarView.ts`.
  // We intentionally removed the older, unwired `name-box/` implementation to avoid
  // having multiple competing sources of truth.
  const legacyDir = path.join(repoRoot, "apps/desktop/src/formula-bar/name-box");
  assert.equal(existsSync(legacyDir), false, `Expected legacy name-box directory to be absent: ${legacyDir}`);
});

