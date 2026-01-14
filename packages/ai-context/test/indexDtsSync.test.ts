import { readFile } from "node:fs/promises";

import { expect, test } from "vitest";

import { stripComments } from "../../../apps/desktop/test/sourceTextUtils.js";

function extractStarExports(code: string): string[] {
  const exports: string[] = [];
  // Allow optional semicolons so the test keeps working if we ever switch to a
  // semicolon-less formatting style.
  const re = /^\s*export\s+\*\s+from\s+["'](.+?)["']\s*;?\s*$/gm;
  for (let match; (match = re.exec(code)); ) {
    exports.push(match[1]!);
  }
  return exports;
}

test("src/index.d.ts mirrors src/index.js exports and all exported modules have .d.ts files", async () => {
  const indexJsUrl = new URL("../src/index.js", import.meta.url);
  const indexDtsUrl = new URL("../src/index.d.ts", import.meta.url);

  const indexJs = stripComments(await readFile(indexJsUrl, "utf8"));
  const indexDts = stripComments(await readFile(indexDtsUrl, "utf8"));

  const jsExports = extractStarExports(indexJs);
  const dtsExports = extractStarExports(indexDts);

  // Ensure no accidental duplicates (keeps the entrypoint deterministic).
  expect(jsExports.length).toBe(new Set(jsExports).size);
  expect(dtsExports.length).toBe(new Set(dtsExports).size);

  // Ensure TS entrypoint stays in sync with runtime entrypoint.
  expect(new Set(dtsExports)).toEqual(new Set(jsExports));

  // Ensure every exported runtime module has a sibling `.d.ts` file so TS consumers
  // importing from `src/index.js` see full types.
  for (const spec of jsExports) {
    if (!spec.endsWith(".js")) continue;
    const dtsUrl = new URL(spec.replace(/\.js$/, ".d.ts"), indexJsUrl);
    await readFile(dtsUrl, "utf8");
  }
});
