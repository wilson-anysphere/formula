import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import { stripComments } from "../../../apps/desktop/test/sourceTextUtils.js";

function extractStarExports(code) {
  const exports = [];
  // Allow optional semicolons so the test stays robust across formatting changes.
  const re = /^\s*export\s+\*\s+from\s+["'](.+?)["']\s*;?\s*$/gm;
  for (let match; (match = re.exec(code)); ) {
    exports.push(match[1]);
  }
  return exports;
}

test("ai-context: src/index.d.ts mirrors src/index.js exports", async () => {
  const indexJsUrl = new URL("../src/index.js", import.meta.url);
  const indexDtsUrl = new URL("../src/index.d.ts", import.meta.url);

  const indexJs = stripComments(await readFile(indexJsUrl, "utf8"));
  const indexDts = stripComments(await readFile(indexDtsUrl, "utf8"));

  const jsExports = extractStarExports(indexJs);
  const dtsExports = extractStarExports(indexDts);

  assert.strictEqual(jsExports.length, new Set(jsExports).size, "src/index.js should not contain duplicate exports");
  assert.strictEqual(dtsExports.length, new Set(dtsExports).size, "src/index.d.ts should not contain duplicate exports");

  const sortedJs = [...new Set(jsExports)].sort();
  const sortedDts = [...new Set(dtsExports)].sort();
  assert.deepStrictEqual(sortedDts, sortedJs);

  // Ensure every exported runtime module has a sibling `.d.ts` file so TS consumers importing
  // from `src/index.js` see full types.
  for (const spec of jsExports) {
    if (!spec.endsWith(".js")) continue;
    const dtsUrl = new URL(spec.replace(/\.js$/, ".d.ts"), indexJsUrl);
    await readFile(dtsUrl, "utf8");
  }
});
