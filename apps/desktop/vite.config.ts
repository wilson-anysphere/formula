import { existsSync, readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { defineConfig } from "vite";

const require = createRequire(import.meta.url);

const repoRoot = fileURLToPath(new URL("../..", import.meta.url));

const extensionApiEntry = fileURLToPath(new URL("../../packages/extension-api/index.mjs", import.meta.url));
const extensionMarketplaceEntry = fileURLToPath(
  new URL("../../packages/extension-marketplace/src/index.ts", import.meta.url),
);
const marketplaceSharedEntry = fileURLToPath(new URL("../../shared", import.meta.url));
const collabUndoEntry = fileURLToPath(new URL("../../packages/collab/undo/index.js", import.meta.url));
const collabCommentsEntry = fileURLToPath(new URL("../../packages/collab/comments/src/index.ts", import.meta.url));
const collabSessionEntry = fileURLToPath(new URL("../../packages/collab/session/src/index.ts", import.meta.url));
const collabEncryptionEntry = fileURLToPath(new URL("../../packages/collab/encryption/src/index.ts", import.meta.url));
const collabEncryptedRangesEntry = fileURLToPath(new URL("../../packages/collab/encrypted-ranges/src/index.ts", import.meta.url));
const collabWorkbookEntry = fileURLToPath(new URL("../../packages/collab/workbook/src/index.ts", import.meta.url));
const collabYjsUtilsEntry = fileURLToPath(new URL("../../packages/collab/yjs-utils/src/index.ts", import.meta.url));
const collabVersioningEntry = fileURLToPath(new URL("../../packages/collab/versioning/src/index.ts", import.meta.url));
const collabPersistenceEntry = fileURLToPath(new URL("../../packages/collab/persistence/src/index.ts", import.meta.url));
const collabPersistenceIndexedDbEntry = fileURLToPath(
  new URL("../../packages/collab/persistence/src/indexeddb.ts", import.meta.url),
);
const spreadsheetFrontendTokenizerEntry = fileURLToPath(
  new URL("../../packages/spreadsheet-frontend/src/formula/tokenizeFormula.ts", import.meta.url),
);
const tauriConfigPath = fileURLToPath(new URL("./src-tauri/tauri.conf.json", import.meta.url));
const tauriCsp = (JSON.parse(readFileSync(tauriConfigPath, "utf8")) as any)?.app?.security?.csp as unknown;
const isE2E = process.env.FORMULA_E2E === "1";
const isPlaywright = process.env.FORMULA_E2E === "0" || process.env.FORMULA_E2E === "1";
const isBundleAnalyze = process.env.VITE_BUNDLE_ANALYZE === "1";
const isBundleAnalyzeSourcemap = process.env.VITE_BUNDLE_ANALYZE_SOURCEMAP === "1";
const visualizer: null | ((opts: any) => any) = (() => {
  if (!isBundleAnalyze) return null;
  try {
    return (require("rollup-plugin-visualizer") as any).visualizer as (opts: any) => any;
  } catch {
    return null;
  }
})();
const cacheDir =
  process.env.FORMULA_E2E === "1"
    ? "node_modules/.vite-e2e-csp"
    : process.env.FORMULA_E2E === "0"
      ? "node_modules/.vite-e2e"
      : undefined;

if (isE2E && typeof tauriCsp !== "string") {
  throw new Error("Missing `app.security.csp` in src-tauri/tauri.conf.json (required for CSP e2e tests)");
}
const crossOriginIsolationHeaders = {
  // Required for SharedArrayBuffer in Chromium (crossOriginIsolated === true).
  // PyodideRuntime relies on SharedArrayBuffer + Atomics for synchronous RPC.
  "Cross-Origin-Opener-Policy": "same-origin",
  "Cross-Origin-Embedder-Policy": "require-corp",
};

function resolveJsToTs() {
  return {
    name: "formula:resolve-js-to-ts",
    enforce: "pre" as const,
    /**
     * Workspace packages use ESM-style `.js` import specifiers in TS source
     * files (e.g. `import './types.js'` next to `types.ts`). TypeScript can
     * resolve these, but Vite treats workspace packages as dependencies and
     * does not apply that mapping by default.
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
  root: ".",
  cacheDir,
  // Expose desktop-specific tuning flags to the client bundle. Prefer URL query
  // params for ad-hoc overrides, but allow env-based configuration for packaged
  // builds where query params are harder to inject.
  envPrefix: ["VITE_", "DESKTOP_LOAD_"],
  plugins: [
    resolveJsToTs(),
    ...(typeof visualizer === "function"
      ? [
          visualizer({
            filename: "bundle-stats.html",
            template: "treemap",
            gzipSize: true,
            brotliSize: true,
            projectRoot: repoRoot,
            emitFile: true,
          }),
          visualizer({
            filename: "bundle-stats.json",
            template: "raw-data",
            gzipSize: true,
            brotliSize: true,
            projectRoot: repoRoot,
            emitFile: true,
          }),
          ...(isBundleAnalyzeSourcemap
            ? [
                visualizer({
                  filename: "bundle-stats-sourcemap.html",
                  template: "treemap",
                  // Use Rollup sourcemaps to attribute minified output back to source modules.
                  sourcemap: true,
                  // Compressed sizes are less meaningful in sourcemap mode; keep this report focused on attribution.
                  gzipSize: false,
                  brotliSize: false,
                  projectRoot: repoRoot,
                  emitFile: true,
                }),
              ]
            : []),
        ]
      : []),
  ],
  resolve: {
    alias: [
      { find: "@formula/extension-api", replacement: extensionApiEntry },
      { find: "@formula/extension-marketplace", replacement: extensionMarketplaceEntry },
      // `@formula/marketplace-shared` lives under the repo `shared/` directory. Some CI/dev environments
      // can have stale node_modules (cached installs) that miss the pnpm workspace link, so keep an
      // explicit alias to ensure Vite can resolve browser-only marketplace helpers during e2e runs.
      { find: /^@formula\/marketplace-shared/, replacement: marketplaceSharedEntry },
      { find: "@formula/collab-comments", replacement: collabCommentsEntry },
      { find: "@formula/collab-undo", replacement: collabUndoEntry },
      { find: "@formula/collab-yjs-utils", replacement: collabYjsUtilsEntry },
      { find: "@formula/collab-session", replacement: collabSessionEntry },
      { find: "@formula/collab-encryption", replacement: collabEncryptionEntry },
      { find: "@formula/collab-encrypted-ranges", replacement: collabEncryptedRangesEntry },
      { find: "@formula/collab-workbook", replacement: collabWorkbookEntry },
      { find: "@formula/collab-versioning", replacement: collabVersioningEntry },
      // Workspace packages are linked via pnpm's node_modules symlinks. Some CI/dev environments
      // can run with stale node_modules (e.g. cached installs), which causes Vite to fail to
      // resolve new workspace dependencies. Alias the persistence entrypoints directly to keep
      // the desktop dev server/e2e harness resilient.
      { find: "@formula/collab-persistence/indexeddb", replacement: collabPersistenceIndexedDbEntry },
      { find: /^@formula\/collab-persistence$/, replacement: collabPersistenceEntry },
      // Like the collab aliases above, keep formula-bar highlighting resilient when the workspace
      // link is stale and the new package subpath export isn't resolvable.
      { find: "@formula/spreadsheet-frontend/formula/tokenizeFormula", replacement: spreadsheetFrontendTokenizerEntry },
    ],
  },
  build: {
    // Desktop (Tauri/WebView) targets modern runtimes. Use a modern output target so
    // optional dependencies (e.g. apache-arrow) can rely on top-level await without
    // breaking production builds.
    target: "es2022",
    // Optional: enable Rollup sourcemaps when doing bundle analysis so visualizer can attribute
    // sizes more accurately back to source modules.
    sourcemap: isBundleAnalyzeSourcemap,
    commonjsOptions: {
      // `shared/` is CommonJS, but the desktop runtime imports ESM shims that depend on
      // `shared/extension-package/core/v2-core.js`. Ensure Rollup runs the CommonJS
      // transform on that file during production builds.
      include: [/node_modules/, /shared[\\/]+extension-package[\\/]+core[\\/]+/],
    },
  },
  server: {
    port: 4174,
    strictPort: true,
    fs: {
      // Allow serving workspace packages during dev (`packages/*`).
      allow: [repoRoot],
    },
    // The Grind/CI environment can run many agents in parallel, exhausting the default inotify
    // watcher limits and causing Vite startup to fail with EMFILE/ENOSPC. E2E does not rely on
    // file watching, so use polling to keep the dev server resilient under load.
    ...(isPlaywright ? { watch: { usePolling: true, interval: 1000 } } : {}),
    // Always enable COOP/COEP so the webview can use SharedArrayBuffer.
    headers: {
      ...crossOriginIsolationHeaders,
      ...(isE2E && typeof tauriCsp === "string" ? { "Content-Security-Policy": tauriCsp } : {}),
    },
    ...(isE2E
      ? {
          // Avoid Vite HMR WebSocket noise in CSP checks.
          hmr: false,
        }
      : {}),
  },
  preview: {
    headers: crossOriginIsolationHeaders,
  },
  test: {
    environment: "node",
    // Avoid spawning extra Node processes during test runs (important for resource-constrained CI
    // and sandboxed environments). Vitest's thread pool still provides per-file isolation via
    // worker_threads without relying on `child_process.spawn`.
    pool: "threads",
    poolOptions: {
      // Vitest defaults to scaling the thread pool based on CPU count. Some CI/sandbox runners
      // report a very high core count (e.g. `nproc` in the hundreds) even when OS thread limits
      // are much lower. Cap the pool so single-file runs don't pre-spawn an enormous number of
      // worker threads (which can lead to flaky shutdowns like tinypool "Failed to terminate worker").
      threads: {
        minThreads: 1,
        maxThreads: 8,
      },
    },
    // Only require a single worker to start. (Pool options cap the maximum.)
    minWorkers: 1,
    // Desktop unit tests can incur a fair amount of Vite/React compilation overhead on
    // shared runners; keep timeouts generous so we don't flake on cold caches.
    testTimeout: 30_000,
    hookTimeout: 30_000,
    // Node 22+ ships an experimental `localStorage` accessor that throws unless started with
    // `--localstorage-file`. Provide a stable in-memory fallback for tests.
    setupFiles: ["./vitest.setup.ts"],
    environmentMatchGlobs: [
      ["src/panels/ai-audit/AIAuditPanel.vitest.ts", "jsdom"],
    ],
    include: [
      "src/**/*.vitest.ts",
      "src/ai/tools/**/*.test.ts",
      "src/ai/dlp/__tests__/**/*.test.ts",
      "src/editor/cellEditorOverlay.f4.test.ts",
      "src/ai/inline-edit/__tests__/**/*.test.ts",
      // Drawing overlay interaction unit tests live under `src/drawings/__tests__` but use the
      // `.test.ts` suffix for historical reasons. Include the specific suite we rely on for
      // transform-aware selection handle hit-testing so `pnpm -C apps/desktop vitest run ...` works.
      "src/drawings/__tests__/selectionHandles.test.ts",
    ],
    exclude: ["tests/**", "node_modules/**"],
  },
});
