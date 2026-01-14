import { describe, expect, it } from "vitest";

import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { promises as fs } from "node:fs";

import { build } from "esbuild";

type OutputImport = { path: string };
type MetafileOutput = { imports: OutputImport[] };

function collectOutputImports(result: { metafile?: { outputs?: Record<string, MetafileOutput> } }): OutputImport[] {
  const outputs = result.metafile?.outputs ?? ({} as Record<string, MetafileOutput>);
  return Object.values(outputs).flatMap((output) => output.imports);
}

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
           import { AIAuditRecorder, LocalStorageAIAuditStore, createDefaultAIAuditStore } from "@formula/ai-audit";
           import { BoundedAIAuditStore } from "@formula/ai-audit";
           import { LocalStorageAIAuditStore as FromBrowserSubpath } from "@formula/ai-audit/browser";
           console.log(AIAuditRecorder, LocalStorageAIAuditStore, FromBrowserSubpath, BoundedAIAuditStore, createDefaultAIAuditStore);
         `
       }
     });

    const inputFiles = Object.keys(result.metafile?.inputs ?? {});
    expect(inputFiles.some((file) => file.includes("sql.js"))).toBe(false);
    expect(inputFiles.some((file) => file.includes("sqlite-store"))).toBe(false);
    expect(inputFiles.some((file) => file.includes("storage.node"))).toBe(false);

    const outputImports = collectOutputImports(result);
    expect(outputImports.some((imp) => imp.path.startsWith("node:"))).toBe(false);
  });

  it("bundles the export entrypoint without pulling in sql.js or Node-only modules", async () => {
    const here = path.dirname(fileURLToPath(import.meta.url));
    const pkgRoot = path.resolve(here, "..");

    const outdir = await fs.mkdtemp(path.join(os.tmpdir(), "ai-audit-export-esbuild-"));

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
          import { serializeAuditEntries } from "@formula/ai-audit/export";
          console.log(serializeAuditEntries);
        `
      }
    });

    const inputFiles = Object.keys(result.metafile?.inputs ?? {});
    expect(inputFiles.some((file) => file.includes("sql.js"))).toBe(false);
    expect(inputFiles.some((file) => file.includes("sqlite-store"))).toBe(false);
    expect(inputFiles.some((file) => file.includes("storage.node"))).toBe(false);

    const outputImports = collectOutputImports(result);
    expect(outputImports.some((imp) => imp.path.startsWith("node:"))).toBe(false);
  });

  it("bundles the sqlite entrypoint for browser builds without Node builtins", async () => {
    const here = path.dirname(fileURLToPath(import.meta.url));
    const pkgRoot = path.resolve(here, "..");

    const outdir = await fs.mkdtemp(path.join(os.tmpdir(), "ai-audit-sqlite-esbuild-"));

    const result = await build({
      absWorkingDir: pkgRoot,
      bundle: true,
      format: "esm",
      metafile: true,
      outdir,
      platform: "browser",
      external: ["fs", "path", "crypto"],
      write: false,
      stdin: {
        sourcefile: "entry.js",
        resolveDir: pkgRoot,
        contents: `
          import { SqliteAIAuditStore } from "@formula/ai-audit/sqlite";
          import { LocalStorageBinaryStorage } from "@formula/ai-audit/browser";
          console.log(SqliteAIAuditStore, LocalStorageBinaryStorage);
        `
      }
    });

    const inputFiles = Object.keys(result.metafile?.inputs ?? {});
    expect(inputFiles.some((file) => file.includes("sql.js"))).toBe(true);
    expect(inputFiles.some((file) => file.includes("storage.node"))).toBe(false);

    const outputImports = collectOutputImports(result);
    expect(outputImports.some((imp) => imp.path.startsWith("node:"))).toBe(false);
  });
});
