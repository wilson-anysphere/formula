import assert from "node:assert/strict";
import test from "node:test";

import os from "node:os";
import path from "node:path";
import { builtinModules } from "node:module";
import { promises as fs } from "node:fs";
import { fileURLToPath } from "node:url";

import { build } from "esbuild";

function collectOutputImports(result) {
  const outputs = result.metafile?.outputs ?? {};
  return Object.values(outputs).flatMap((output) => output.imports ?? []);
}

function createNodeBuiltinsSet() {
  /** @type {Set<string>} */
  const builtins = new Set();
  for (const mod of builtinModules) {
    builtins.add(mod);
    if (mod.startsWith("node:")) {
      builtins.add(mod.slice("node:".length));
    } else {
      builtins.add(`node:${mod}`);
    }
  }
  return builtins;
}

test("ai-completion bundles for the browser without Node builtins", async () => {
  const here = path.dirname(fileURLToPath(import.meta.url));
  const pkgRoot = path.resolve(here, "..");

  const outdir = await fs.mkdtemp(path.join(os.tmpdir(), "ai-completion-esbuild-"));

  const result = await build({
    absWorkingDir: pkgRoot,
    bundle: true,
    format: "esm",
    metafile: true,
    outdir,
    platform: "browser",
    write: false,
    stdin: {
      sourcefile: "entry.js",
      resolveDir: pkgRoot,
      contents: `
        import { TabCompletionEngine } from "@formula/ai-completion";
        console.log(TabCompletionEngine);
      `,
    },
  });

  assert.ok(result.outputFiles?.length, "Expected esbuild to produce outputFiles (write:false)");

  const inputFiles = Object.keys(result.metafile?.inputs ?? {}).map((file) => file.replace(/\\/g, "/"));
  assert.ok(
    inputFiles.some((file) => file.endsWith("src/index.js")),
    `Expected the bundle to include ai-completion sources, got inputs:\n${inputFiles.join("\n")}`,
  );

  const nodeBuiltins = createNodeBuiltinsSet();
  const outputImports = collectOutputImports(result);
  const offending = outputImports.filter((imp) => nodeBuiltins.has(imp.path));
  assert.deepEqual(
    offending,
    [],
    `Expected browser bundle to avoid Node builtin imports, found: ${offending.map((imp) => imp.path).join(", ")}`,
  );
});
