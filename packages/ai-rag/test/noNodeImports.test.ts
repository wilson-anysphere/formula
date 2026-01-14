import { readFile, stat } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { expect, test } from "vitest";

import { stripComments } from "../../../apps/desktop/test/sourceTextUtils.js";

const browserSafeEntrypoints = [
  new URL("../src/index.js", import.meta.url),
  new URL("../src/embedding/hashEmbedder.js", import.meta.url),
  new URL("../src/store/binaryStorage.js", import.meta.url),
  new URL("../src/store/inMemoryVectorStore.js", import.meta.url),
  new URL("../src/store/jsonVectorStore.js", import.meta.url),
  new URL("../src/store/sqliteVectorStore.js", import.meta.url),
  new URL("../src/pipeline/indexWorkbook.js", import.meta.url),
  new URL("../src/workbook/chunkWorkbook.js", import.meta.url),
  new URL("../src/workbook/chunkToText.js", import.meta.url),
  new URL("../src/workbook/fromSpreadsheetApi.js", import.meta.url),
  new URL("../src/workbook/rect.js", import.meta.url),
  new URL("../src/utils/abort.js", import.meta.url),
  new URL("../src/utils/hash.js", import.meta.url),
  new URL("../src/retrieval/searchWorkbookRag.js", import.meta.url),
  new URL("../src/retrieval/ranking.js", import.meta.url),
  new URL("../src/retrieval/rankResults.js", import.meta.url),
  new URL("../../../apps/desktop/src/ai/rag/index.js", import.meta.url),
];

const aiRagSrcRoot = new URL("../src/", import.meta.url);
const aiRagSrcRootPath = fileURLToPath(aiRagSrcRoot);

function collectRelativeSpecifiers(code: string): string[] {
  const out: string[] = [];
  const importFromRe = /\b(?:import|export)\s+(type\s+)?[^"']*?\sfrom\s+["']([^"']+)["']/g;
  const sideEffectImportRe = /\bimport\s+["']([^"']+)["']/g;
  const dynamicImportRe = /\bimport\(\s*["']([^"']+)["']\s*\)/g;

  for (const match of code.matchAll(importFromRe)) {
    const typeOnly = Boolean(match[1]);
    const spec = match[2];
    if (!spec || typeOnly) continue;
    if (spec.startsWith("./") || spec.startsWith("../")) out.push(spec);
  }
  for (const match of code.matchAll(sideEffectImportRe)) {
    const spec = match[1];
    if (!spec) continue;
    if (spec.startsWith("./") || spec.startsWith("../")) out.push(spec);
  }
  for (const match of code.matchAll(dynamicImportRe)) {
    const spec = match[1];
    if (!spec) continue;
    if (spec.startsWith("./") || spec.startsWith("../")) out.push(spec);
  }

  return out;
}

async function resolveRelativeModule(fromPath: string, specifier: string): Promise<string | null> {
  const cleaned = specifier.split("?")[0]!.split("#")[0]!;
  const base = path.resolve(path.dirname(fromPath), cleaned);
  if (path.extname(base)) {
    try {
      const s = await stat(base);
      if (s.isFile()) return base;
    } catch {
      return null;
    }
    return null;
  }

  // ESM in this repo always uses explicit `.js` extensions, but keep a small fallback for robustness.
  const candidates = [`${base}.js`, `${base}.mjs`, `${base}.cjs`];
  for (const candidate of candidates) {
    try {
      const s = await stat(candidate);
      if (s.isFile()) return candidate;
    } catch {
      // continue
    }
  }
  return null;
}

async function scanTransitiveForNodeImports(entryPath: string): Promise<void> {
  const visited = new Set<string>();

  async function visit(filePath: string): Promise<void> {
    if (visited.has(filePath)) return;
    visited.add(filePath);

    const code = stripComments(await readFile(filePath, "utf8"));
    expect(code, `${filePath} should not statically import node:*`).not.toMatch(/from\s+["']node:/);
    expect(code, `${filePath} should not statically import node:*`).not.toMatch(/import\(\s*["']node:/);
    expect(code, `${filePath} should not statically import node:*`).not.toMatch(/\bimport\s+["']node:/);

    for (const specifier of collectRelativeSpecifiers(code)) {
      const resolved = await resolveRelativeModule(filePath, specifier);
      if (!resolved) continue;
      // Only traverse within ai-rag's src directory so this stays fast and deterministic.
      if (!resolved.startsWith(aiRagSrcRootPath)) continue;
      await visit(resolved);
    }
  }

  await visit(entryPath);
}

test("browser-safe entrypoints do not contain static node:* imports", async () => {
  for (const url of browserSafeEntrypoints) {
    const code = stripComments(await readFile(url, "utf8"));
    expect(code, `${url} should not statically import node:*`).not.toMatch(/from\s+["']node:/);
    expect(code, `${url} should not statically import node:*`).not.toMatch(/import\(\s*["']node:/);
    expect(code, `${url} should not statically import node:*`).not.toMatch(/\bimport\s+["']node:/);

    // For ai-rag's own modules, scan transitive relative imports to ensure dependencies remain browser-safe too.
    const entryPath = fileURLToPath(url);
    if (entryPath.startsWith(aiRagSrcRootPath)) {
      await scanTransitiveForNodeImports(entryPath);
    }
  }
});
