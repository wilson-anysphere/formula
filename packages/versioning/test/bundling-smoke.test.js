import assert from "node:assert/strict";
import test from "node:test";

import os from "node:os";
import path from "node:path";
import { builtinModules } from "node:module";
import { promises as fs } from "node:fs";
import { fileURLToPath } from "node:url";

/** @type {import("esbuild").build | null} */
let build = null;
try {
  // `esbuild` is a dev dependency that may be missing in some lightweight
  // environments (e.g. sandboxes that run tests without installing node_modules).
  // Skip the bundling smoke test in that case.
  // eslint-disable-next-line node/no-unsupported-features/es-syntax
  ({ build } = await import("esbuild"));
} catch {
  build = null;
}

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

test("versioning diff helpers bundle for the browser without Node builtins", { skip: !build }, async () => {
  const here = path.dirname(fileURLToPath(import.meta.url));
  const pkgRoot = path.resolve(here, "..");
  const repoRoot = path.resolve(pkgRoot, "../..");

  const outdir = await fs.mkdtemp(path.join(os.tmpdir(), "versioning-esbuild-"));

  const result = await build({
    absWorkingDir: repoRoot,
    bundle: true,
    format: "esm",
    metafile: true,
    outdir,
    platform: "browser",
    write: false,
    stdin: {
      sourcefile: "entry.js",
      resolveDir: repoRoot,
      contents: `
        import { semanticDiff } from "./packages/versioning/src/diff/semanticDiff.js";
        import { diffYjsWorkbookSnapshots } from "./packages/versioning/src/yjs/diffWorkbookSnapshots.js";
        import { diffDocumentWorkbookSnapshots } from "./packages/versioning/src/document/diffWorkbookSnapshots.js";
        console.log(semanticDiff, diffYjsWorkbookSnapshots, diffDocumentWorkbookSnapshots);
      `,
    },
  });

  assert.ok(result.outputFiles?.length, "Expected esbuild to produce outputFiles (write:false)");

  const inputFiles = Object.keys(result.metafile?.inputs ?? {}).map((file) => file.replace(/\\/g, "/"));
  assert.ok(
    inputFiles.some((file) => file.endsWith("packages/versioning/src/diff/semanticDiff.js")),
    `Expected the bundle to include versioning sources, got inputs:\n${inputFiles.join("\n")}`,
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

