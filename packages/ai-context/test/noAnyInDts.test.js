import assert from "node:assert/strict";
import { readdir, readFile } from "node:fs/promises";
import { join } from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

async function collectDtsFiles(dir, out = []) {
  const entries = await readdir(dir, { withFileTypes: true });
  for (const entry of entries) {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) {
      await collectDtsFiles(full, out);
      continue;
    }
    if (!entry.isFile()) continue;
    if (!entry.name.endsWith(".d.ts")) continue;
    out.push(full);
  }
  return out;
}

function stripComments(code) {
  // This is a lightweight heuristic to avoid flagging the English word "any" in doc comments.
  // It is not a full TS lexer, but it is sufficient for `.d.ts` files in this package.
  return code.replace(/\/\*[\s\S]*?\*\//g, "").replace(/\/\/.*$/gm, "");
}

test("ai-context .d.ts files do not use the `any` type", async () => {
  const srcDir = fileURLToPath(new URL("../src", import.meta.url));
  const files = await collectDtsFiles(srcDir);

  const offenders = [];
  for (const file of files) {
    const text = stripComments(await readFile(file, "utf8"));
    // Use word boundaries so we don't flag ordinary words like "many".
    if (/\bany\b/.test(text)) offenders.push(file.slice(srcDir.length + 1));
  }

  offenders.sort();
  assert.deepEqual(offenders, []);
});
