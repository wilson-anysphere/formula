import assert from "node:assert/strict";
import fs from "node:fs";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { stripComments } from "../../../apps/desktop/test/sourceTextUtils.js";

function read(relativeUrl) {
  return fs.readFileSync(fileURLToPath(new URL(relativeUrl, import.meta.url)), "utf8");
}

function assertBundlerFriendlyDtsSurface(fileLabel, content) {
  const source = stripComments(content);
  // TypeScript's `--moduleResolution Bundler` can be sensitive to explicit `.d.ts` specifiers.
  // Ensure these declaration entrypoints only reference the runtime `.js` module specifiers,
  // letting TS follow them to the corresponding declaration output.
  assert.equal(
    /(?:import|export)\s+[^;]*from\s+["'][^"']*\.d\.ts["']/.test(source),
    false,
    `Expected ${fileLabel} to avoid importing/re-exporting .d.ts specifiers`,
  );

  // Keep the entrypoint aligned with the runtime ESM surface (`src/{node,web}.js` both export from `./index.js`).
  assert.match(source, /export\s+\*\s+from\s+["']\.\/index\.js["'];/);

  // Ensure types referenced by declarations in this file are imported into scope (not just re-exported).
  assert.match(source, /import\s+type\s+\{\s*WorkbookLike\s*\}\s+from\s+["']\.\/index\.js["'];/);
}

test("scripting: node/web .d.ts entrypoints avoid .d.ts specifiers and re-export runtime surface", () => {
  assertBundlerFriendlyDtsSurface("packages/scripting/src/node.d.ts", read("../src/node.d.ts"));
  assertBundlerFriendlyDtsSurface("packages/scripting/src/web.d.ts", read("../src/web.d.ts"));
});
