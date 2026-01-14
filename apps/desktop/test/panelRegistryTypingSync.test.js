import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { PanelIds } from "../src/panels/panelRegistry.js";
import { stripComments } from "./sourceTextUtils.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("panelRegistry.d.ts PanelIds stays in sync with runtime PanelIds", () => {
  const dtsPath = path.join(__dirname, "..", "src", "panels", "panelRegistry.d.ts");
  const dts = stripComments(fs.readFileSync(dtsPath, "utf8"));

  const dtsKeys = [...dts.matchAll(/^\s+([A-Z0-9_]+):\s*string;/gm)].map((m) => m[1]).sort();
  assert.ok(dtsKeys.length > 0, "Expected to find PanelIds keys in panelRegistry.d.ts");

  const runtimeKeys = Object.keys(PanelIds).sort();
  assert.deepEqual(
    dtsKeys,
    runtimeKeys,
    `Expected panelRegistry.d.ts PanelIds keys to match runtime exports.\n` +
      `d.ts: ${dtsKeys.join(", ")}\n` +
      `js:  ${runtimeKeys.join(", ")}`,
  );
});
