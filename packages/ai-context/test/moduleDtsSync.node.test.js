import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import { stripComments } from "../../../apps/desktop/test/sourceTextUtils.js";

function extractStarExports(code) {
  const exports = [];
  // Allow optional semicolons so this stays robust across formatting styles.
  const re = /^\s*export\s+\*\s+from\s+["'](.+?)["']\s*;?\s*$/gm;
  for (let match; (match = re.exec(code)); ) {
    exports.push(match[1]);
  }
  return exports;
}

function extractJsNamedExports(code) {
  const names = new Set();

  // export const NAME =
  for (const m of code.matchAll(/^\s*export\s+const\s+([A-Za-z0-9_]+)\b/gm)) names.add(m[1]);
  // export let/var NAME
  for (const m of code.matchAll(/^\s*export\s+(?:let|var)\s+([A-Za-z0-9_]+)\b/gm)) names.add(m[1]);
  // export function NAME / export async function NAME
  for (const m of code.matchAll(/^\s*export\s+(?:async\s+)?function\s+([A-Za-z0-9_]+)\b/gm)) names.add(m[1]);
  // export class NAME
  for (const m of code.matchAll(/^\s*export\s+class\s+([A-Za-z0-9_]+)\b/gm)) names.add(m[1]);

  // export { foo, bar as baz } [from "..."]
  for (const m of code.matchAll(/^\s*export\s*\{([^}]+)\}\s*(?:from\s+["'][^"']+["'])?\s*;?/gm)) {
    let list = m[1] ?? "";
    // Strip comments from the export list so formats like:
    //   export { foo, /* comment */ bar }
    // or:
    //   export { foo, // comment\n bar }
    // don't confuse the parser.
    list = stripComments(list);
    for (const part of list.split(",")) {
      const trimmed = part.trim();
      if (!trimmed) continue;
      // Handle `foo as bar`
      const asMatch = /\bas\b/i.exec(trimmed);
      if (asMatch) {
        const [, alias] = trimmed.split(/\bas\b/i).map((s) => s.trim());
        if (alias) names.add(alias);
      } else {
        names.add(trimmed);
      }
    }
  }

  return [...names];
}

function dtsDeclaresExport(dts, name) {
  const patterns = [
    new RegExp(`\\bexport\\s+(?:declare\\s+)?(?:const|let|var|function|class)\\s+${name}\\b`),
    // `export { foo }` or `export { foo as bar }`
    new RegExp(`\\bexport\\s*\\{[^}]*\\b${name}\\b[^}]*\\}`),
  ];
  return patterns.some((re) => re.test(dts));
}

test("ai-context: all index-exported modules have matching named exports in their .d.ts files", async () => {
  const indexJsUrl = new URL("../src/index.js", import.meta.url);
  const indexJs = stripComments(await readFile(indexJsUrl, "utf8"));

  const modules = extractStarExports(indexJs).filter((spec) => spec.endsWith(".js"));

  for (const spec of modules) {
    const jsUrl = new URL(spec, indexJsUrl);
    const dtsUrl = new URL(spec.replace(/\.js$/, ".d.ts"), indexJsUrl);

    const js = stripComments(await readFile(jsUrl, "utf8"));
    const dts = stripComments(await readFile(dtsUrl, "utf8"));

    const named = extractJsNamedExports(js);
    const missing = named.filter((name) => !dtsDeclaresExport(dts, name));

    assert.deepStrictEqual(missing, [], `${spec} exports missing from ${spec.replace(/\.js$/, ".d.ts")}`);
  }
});
