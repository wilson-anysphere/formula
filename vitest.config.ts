import { existsSync, readdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { defineConfig } from "vitest/config";

const MAX_VITEST_THREADS = 8;

const repoRoot = fileURLToPath(new URL(".", import.meta.url));

function hasPnpmDependency(pkgName: string): boolean {
  const pnpmDir = resolve(repoRoot, "node_modules", ".pnpm");
  if (!existsSync(pnpmDir)) return false;
  try {
    // pnpm virtual store directory names start with `${name}@...` when present.
    return readdirSync(pnpmDir).some((entry) => entry.startsWith(`${pkgName}@`));
  } catch {
    return false;
  }
}

// Workspace packages are normally resolved via pnpm's node_modules symlinks. Some CI/dev
// environments can run with stale node_modules (cached installs), which causes Vite/Vitest to fail
// to resolve newly-added workspace dependencies. Alias the collab workspace entrypoints directly so
// cross-package integration suites (apps/desktop) remain resilient.
const collabUndoEntry = resolve(repoRoot, "packages/collab/undo/index.js");
const collabCommentsEntry = resolve(repoRoot, "packages/collab/comments/src/index.ts");
const collabPresenceEntry = resolve(repoRoot, "packages/collab/presence/index.js");
const collabSessionEntry = resolve(repoRoot, "packages/collab/session/src/index.ts");
const collabVersioningEntry = resolve(repoRoot, "packages/collab/versioning/src/index.ts");
const collabPersistenceEntry = resolve(repoRoot, "packages/collab/persistence/src/index.ts");
const collabPersistenceIndexedDbEntry = resolve(repoRoot, "packages/collab/persistence/src/indexeddb.ts");
const collabYjsUtilsEntry = resolve(repoRoot, "packages/collab/yjs-utils/src/index.ts");
const collabEncryptedRangesEntry = resolve(repoRoot, "packages/collab/encrypted-ranges/src/index.ts");
const collabConflictsEntry = resolve(repoRoot, "packages/collab/conflicts/index.js");
const collabWorkbookEntry = resolve(repoRoot, "packages/collab/workbook/src/index.ts");
const collabEncryptionEntry = resolve(repoRoot, "packages/collab/encryption/src/index.ts");
const marketplaceSharedEntry = resolve(repoRoot, "shared");
const extensionMarketplaceEntry = resolve(repoRoot, "packages/extension-marketplace/src/index.ts");
const gridEntry = resolve(repoRoot, "packages/grid/src/index.ts");
const gridNodeEntry = resolve(repoRoot, "packages/grid/src/node.ts");
const fillEngineEntry = resolve(repoRoot, "packages/fill-engine/src/index.ts");
const textLayoutEntry = resolve(repoRoot, "packages/text-layout/src/index.js");
const textLayoutHarfBuzzEntry = resolve(repoRoot, "packages/text-layout/src/harfbuzz.js");
const aiAuditEntry = resolve(repoRoot, "packages/ai-audit/src/index.node.ts");
const aiAuditBrowserEntry = resolve(repoRoot, "packages/ai-audit/src/index.ts");
const aiAuditNodeEntry = resolve(repoRoot, "packages/ai-audit/src/index.node.ts");
const aiAuditSqliteEntry = resolve(repoRoot, "packages/ai-audit/src/sqlite.node.ts");
const aiAuditExportEntry = resolve(repoRoot, "packages/ai-audit/src/export.ts");
const aiCompletionEntry = resolve(repoRoot, "packages/ai-completion/src/index.js");
const auditCoreEntry = resolve(repoRoot, "packages/audit-core/index.js");
const extensionApiEntry = resolve(repoRoot, "packages/extension-api/index.mjs");
const pythonRuntimeEntry = resolve(repoRoot, "packages/python-runtime/src/index-node.js");
const pythonRuntimeDocumentControllerEntry = resolve(repoRoot, "packages/python-runtime/src/document-controller.js");
const pythonRuntimeNativeEntry = resolve(repoRoot, "packages/python-runtime/src/native.js");
const pythonRuntimePyodideEntry = resolve(repoRoot, "packages/python-runtime/src/pyodide.js");
const pythonRuntimeTestUtilsEntry = resolve(repoRoot, "packages/python-runtime/src/test-utils.js");
const powerQueryEntry = resolve(repoRoot, "packages/power-query/src/index.js");
const workbookBackendEntry = resolve(repoRoot, "packages/workbook-backend/src/index.ts");
const graphemeSplitterShimEntry = resolve(repoRoot, "scripts/vitest-shims/grapheme-splitter.ts");
const linebreakShimEntry = resolve(repoRoot, "scripts/vitest-shims/linebreak.ts");
const zodShimEntry = resolve(repoRoot, "scripts/vitest-shims/zod.ts");
const yWebsocketShimEntry = resolve(repoRoot, "scripts/vitest-shims/y-websocket.ts");
const graphemeSplitterPackageEntry = resolve(repoRoot, "node_modules/grapheme-splitter");
const graphemeSplitterTextLayoutEntry = resolve(repoRoot, "packages/text-layout/node_modules/grapheme-splitter");
const linebreakPackageEntry = resolve(repoRoot, "node_modules/linebreak");
const linebreakTextLayoutEntry = resolve(repoRoot, "packages/text-layout/node_modules/linebreak");
const zodPackageEntry = resolve(repoRoot, "node_modules/zod");
const zodAiToolsEntry = resolve(repoRoot, "packages/ai-tools/node_modules/zod");
const reactPackageEntry = resolve(repoRoot, "node_modules/react");
const reactDesktopEntry = resolve(repoRoot, "apps/desktop/node_modules/react");
const reactWebEntry = resolve(repoRoot, "apps/web/node_modules/react");
const reactGridEntry = resolve(repoRoot, "packages/grid/node_modules/react");
const yWebsocketPackageEntry = resolve(repoRoot, "node_modules/y-websocket");
const yWebsocketCollabSessionEntry = resolve(repoRoot, "packages/collab/session/node_modules/y-websocket");

function anyExists(paths: string[]): boolean {
  return paths.some((p) => existsSync(p));
}
const spreadsheetFrontendEntry = resolve(repoRoot, "packages/spreadsheet-frontend/src/index.ts");
const spreadsheetFrontendA1Entry = resolve(repoRoot, "packages/spreadsheet-frontend/src/a1.ts");
const spreadsheetFrontendCacheEntry = resolve(repoRoot, "packages/spreadsheet-frontend/src/cache.ts");
const spreadsheetFrontendGridEntry = resolve(repoRoot, "packages/spreadsheet-frontend/src/grid-provider.ts");
const spreadsheetFrontendTokenizerEntry = resolve(
  repoRoot,
  "packages/spreadsheet-frontend/src/formula/tokenizeFormula.ts"
);
const engineEntry = resolve(repoRoot, "packages/engine/src/index.ts");
const engineBackendFormulaEntry = resolve(repoRoot, "packages/engine/src/backend/formula.ts");

function resolveJsToTs() {
  return {
    name: "formula:resolve-js-to-ts",
    enforce: "pre" as const,
    /**
     * Workspace packages use ESM-style `.js` import specifiers in TS source
     * files (e.g. `import './types.js'` next to `types.ts`). TypeScript can
     * resolve these, but Vite/Vitest will not unless we explicitly map.
     *
     * When a relative `.js` import target doesn't exist, fall back to `.ts`/`.tsx`.
     */
    resolveId(source: string, importer?: string) {
      if (!importer) return null;
      if (!source.endsWith(".js")) return null;
      if (!(source.startsWith("./") || source.startsWith("../"))) return null;

      const importerPath = importer.split("?", 1)[0]!;
      const resolved = resolve(dirname(importerPath), source);
      if (existsSync(resolved)) return null;

      const ts = resolved.slice(0, -3) + ".ts";
      if (existsSync(ts)) return ts;

      const tsx = resolved.slice(0, -3) + ".tsx";
      if (existsSync(tsx)) return tsx;

      return null;
    },
  };
}

export default defineConfig({
  plugins: [resolveJsToTs()],
  resolve: {
    alias: [
      // Core UI workspace packages used heavily by desktop tests/benchmarks.
      // `@formula/grid`'s primary entrypoint re-exports React components (TSX). In some cached/stale
      // `node_modules` environments, React may be missing; prefer the Node-friendly entrypoint in
      // those cases so non-React desktop tests can still run.
      ...((anyExists([reactPackageEntry, reactDesktopEntry, reactWebEntry, reactGridEntry]) || hasPnpmDependency("react"))
        ? [{ find: /^@formula\/grid$/, replacement: gridEntry }]
        : [{ find: /^@formula\/grid$/, replacement: gridNodeEntry }]),
      { find: /^@formula\/grid\/node$/, replacement: gridNodeEntry },
      { find: /^@formula\/fill-engine$/, replacement: fillEngineEntry },
      { find: /^@formula\/text-layout$/, replacement: textLayoutEntry },
      { find: /^@formula\/text-layout\/harfbuzz$/, replacement: textLayoutHarfBuzzEntry },
      // AI workspace packages used by desktop tests.
      { find: /^@formula\/ai-audit$/, replacement: aiAuditEntry },
      { find: /^@formula\/ai-audit\/browser$/, replacement: aiAuditBrowserEntry },
      { find: /^@formula\/ai-audit\/node$/, replacement: aiAuditNodeEntry },
      { find: /^@formula\/ai-audit\/sqlite$/, replacement: aiAuditSqliteEntry },
      { find: /^@formula\/ai-audit\/export$/, replacement: aiAuditExportEntry },
      { find: /^@formula\/ai-completion$/, replacement: aiCompletionEntry },
      { find: /^@formula\/audit-core$/, replacement: auditCoreEntry },
      { find: /^@formula\/extension-api$/, replacement: extensionApiEntry },
      { find: /^@formula\/python-runtime$/, replacement: pythonRuntimeEntry },
      { find: /^@formula\/python-runtime\/document-controller$/, replacement: pythonRuntimeDocumentControllerEntry },
      { find: /^@formula\/python-runtime\/native$/, replacement: pythonRuntimeNativeEntry },
      { find: /^@formula\/python-runtime\/pyodide$/, replacement: pythonRuntimePyodideEntry },
      { find: /^@formula\/python-runtime\/test-utils$/, replacement: pythonRuntimeTestUtilsEntry },
      { find: /^@formula\/power-query$/, replacement: powerQueryEntry },
      { find: /^@formula\/workbook-backend$/, replacement: workbookBackendEntry },
      // Some cached/stale `node_modules` environments may be missing transitive dependencies of the
      // aliased workspace packages. Provide lightweight shims for pure-JS deps used by
      // `@formula/text-layout` so desktop tests can still run.
      //
      // Prefer the real dependency when it's present; fall back to shims only when the pnpm
      // workspace/cached install is missing the package.
      ...(!(anyExists([graphemeSplitterPackageEntry, graphemeSplitterTextLayoutEntry]) || hasPnpmDependency("grapheme-splitter"))
        ? [{ find: /^grapheme-splitter$/, replacement: graphemeSplitterShimEntry }]
        : []),
      ...(!(anyExists([linebreakPackageEntry, linebreakTextLayoutEntry]) || hasPnpmDependency("linebreak"))
        ? [{ find: /^linebreak$/, replacement: linebreakShimEntry }]
        : []),
      ...(!(anyExists([zodPackageEntry, zodAiToolsEntry]) || hasPnpmDependency("zod"))
        ? [{ find: /^zod$/, replacement: zodShimEntry }]
        : []),
      ...(!(anyExists([yWebsocketPackageEntry, yWebsocketCollabSessionEntry]) || hasPnpmDependency("y-websocket"))
        ? [{ find: /^y-websocket$/, replacement: yWebsocketShimEntry }]
        : []),
      // `@formula/engine` is imported by many desktop + shared packages. Alias it directly so Vitest
      // runs stay resilient in cached/stale `node_modules` environments that may be missing the
      // pnpm workspace link.
      { find: /^@formula\/engine$/, replacement: engineEntry },
      // Also alias the `backend/formula` subpath export which is used by some desktop tooling.
      { find: /^@formula\/engine\/backend\/formula$/, replacement: engineBackendFormulaEntry },
      // `@formula/spreadsheet-frontend` is imported by many desktop modules (A1 helpers, formula ref parsing).
      // Alias it directly so Vitest stays resilient in cached/stale `node_modules` environments.
      { find: /^@formula\/spreadsheet-frontend$/, replacement: spreadsheetFrontendEntry },
      { find: /^@formula\/spreadsheet-frontend\/a1$/, replacement: spreadsheetFrontendA1Entry },
      { find: /^@formula\/spreadsheet-frontend\/cache$/, replacement: spreadsheetFrontendCacheEntry },
      { find: /^@formula\/spreadsheet-frontend\/grid$/, replacement: spreadsheetFrontendGridEntry },
      { find: /^@formula\/spreadsheet-frontend\/formula\/tokenizeFormula$/, replacement: spreadsheetFrontendTokenizerEntry },
      { find: "@formula/extension-marketplace", replacement: extensionMarketplaceEntry },
      { find: "@formula/collab-comments", replacement: collabCommentsEntry },
      { find: "@formula/collab-undo", replacement: collabUndoEntry },
      { find: "@formula/collab-presence", replacement: collabPresenceEntry },
      { find: "@formula/collab-session", replacement: collabSessionEntry },
      { find: "@formula/collab-versioning", replacement: collabVersioningEntry },
      { find: "@formula/collab-encrypted-ranges", replacement: collabEncryptedRangesEntry },
      { find: "@formula/collab-persistence/indexeddb", replacement: collabPersistenceIndexedDbEntry },
      { find: /^@formula\/collab-persistence$/, replacement: collabPersistenceEntry },
      { find: "@formula/collab-conflicts", replacement: collabConflictsEntry },
      { find: "@formula/collab-workbook", replacement: collabWorkbookEntry },
      { find: "@formula/collab-encryption", replacement: collabEncryptionEntry },
      { find: "@formula/collab-yjs-utils", replacement: collabYjsUtilsEntry },
      // `@formula/spreadsheet-frontend/formula/tokenizeFormula` is a subpath export used by the
      // desktop formula bar highlight code. Alias it directly so Vitest stays resilient in
      // cached/stale `node_modules` environments that may not include the latest package exports.
      // (Note: prefer the regex alias above, but keep this explicit mapping for compatibility with
      // any older call sites that might include query params.)
      { find: "@formula/spreadsheet-frontend/formula/tokenizeFormula", replacement: spreadsheetFrontendTokenizerEntry },
      // `@formula/marketplace-shared` lives under the repo `shared/` directory. Like the collab
      // workspace aliases above, we keep an explicit mapping here so Vitest stays resilient in
      // cached/stale `node_modules` environments that may be missing the pnpm workspace link.
      { find: /^@formula\/marketplace-shared/, replacement: marketplaceSharedEntry }
    ],
  },
  test: {
    // The repo includes several integration-style suites (API, sandboxed runtimes,
    // wasm-backed rendering) that can exceed Vitest's default 10s hook timeout on
    // shared/contended runners.
    testTimeout: 30_000,
    hookTimeout: 30_000,
    // Keep parallelism bounded in high-core agent sandboxes to avoid exhausting
    // per-user process/thread limits (Node can abort if it fails to spawn its
    // worker threads).
    maxWorkers: 4,
    minWorkers: 1,
    include: [
      "packages/**/*.test.ts",
      "packages/**/*.test.tsx",
      "packages/**/*.vitest.ts",
      "packages/**/*.vitest.tsx",
      "apps/**/*.test.ts",
      "apps/**/*.test.tsx",
      "apps/**/*.vitest.ts",
      "apps/**/*.vitest.tsx",
      "services/api/src/__tests__/**/*.test.ts",
      "services/api/src/__tests__/**/*.vitest.ts",
    ],
    environment: "node",
    poolOptions: {
      // Vitest defaults to using a worker count based on CPU cores. In some
      // shared/contended environments (CI runners, sandboxes) `nproc` can be very
      // high even when the process is constrained by OS thread limits. Cap the
      // pool size so `vitest run <single test>` doesn't try to spin up hundreds of
      // worker threads (which can lead to flaky shutdowns).
      forks: {
        minForks: 1,
        maxForks: MAX_VITEST_THREADS,
      },
      threads: {
        minThreads: 1,
        maxThreads: MAX_VITEST_THREADS,
      },
    },
    setupFiles: ["./vitest.setup.ts"],
    globalSetup: "./scripts/vitest.global-setup.mjs",
  },
});
