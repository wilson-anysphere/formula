import { describe, expect, it } from "vitest";

import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { promises as fs } from "node:fs";

import { build } from "esbuild";

describe("ai-audit browser bundling", () => {
  it("bundles the browser entrypoints without pulling in sql.js or Node-only modules", async () => {
    const here = path.dirname(fileURLToPath(import.meta.url));
    const pkgRoot = path.resolve(here, "..");

    const outdir = await fs.mkdtemp(path.join(os.tmpdir(), "ai-audit-esbuild-"));

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
          import { AIAuditRecorder, LocalStorageAIAuditStore } from "@formula/ai-audit";
          import { LocalStorageAIAuditStore as FromBrowserSubpath } from "@formula/ai-audit/browser";
          console.log(AIAuditRecorder, LocalStorageAIAuditStore, FromBrowserSubpath);
        `
      }
    });

    const inputFiles = Object.keys(result.metafile?.inputs ?? {});
    expect(inputFiles.some((file) => file.includes("sql.js"))).toBe(false);
    expect(inputFiles.some((file) => file.includes("sqlite-store"))).toBe(false);
    expect(inputFiles.some((file) => file.includes("storage.node"))).toBe(false);

    const outputImports = Object.values(result.metafile?.outputs ?? {}).flatMap((output) => output.imports);
    expect(outputImports.some((imp) => imp.path.startsWith("node:"))).toBe(false);
  });
});

