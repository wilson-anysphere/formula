import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("Desktop main.ts does not special-case legacy AI ribbon ids (dispatch via CommandRegistry)", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = fs.readFileSync(mainPath, "utf8");

  const legacyIds = ["open-panel-ai-chat", "open-inline-ai-edit", "open-ai-panel"];
  for (const id of legacyIds) {
    assert.doesNotMatch(
      main,
      new RegExp(`\\\\bcase\\\\s+["']${id.replace(/[.*+?^${}()|[\\\\]\\\\]/g, "\\\\$&")}["']`),
      `Expected main.ts to avoid special-casing ${id}; ribbon should emit canonical CommandRegistry ids instead`,
    );
  }
});

