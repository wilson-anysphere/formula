import { describe, expect, it } from "vitest";

import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { promises as fs } from "node:fs";

import { build } from "esbuild";

describe("python-runtime browser bundling", () => {
  it("bundles the Pyodide entrypoints without pulling in Node-only modules", async () => {
    const here = path.dirname(fileURLToPath(import.meta.url));
    const pkgRoot = path.resolve(here, "..");

    const outdir = await fs.mkdtemp(path.join(os.tmpdir(), "python-runtime-esbuild-"));

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
          import { PyodideRuntime as FromRoot, formulaFiles } from "@formula/python-runtime";
          import { PyodideRuntime as FromSubpath } from "@formula/python-runtime/pyodide";
          console.log(FromRoot, FromSubpath, formulaFiles && Object.keys(formulaFiles).length);
        `,
      },
    });

    const inputFiles = Object.keys(result.metafile?.inputs ?? {});
    expect(inputFiles.some((file) => file.includes("native-python-runtime"))).toBe(false);
    expect(inputFiles.some((file) => file.includes("mock-workbook"))).toBe(false);
    expect(inputFiles.some((file) => file.includes("document-controller-bridge"))).toBe(false);

    const outputImports = Object.values(result.metafile?.outputs ?? {}).flatMap((output) => output.imports);
    expect(outputImports.some((imp) => imp.path.startsWith("node:"))).toBe(false);
  });
});

