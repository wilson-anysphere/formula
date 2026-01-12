import { existsSync, readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { defineConfig } from "vite";

const repoRoot = fileURLToPath(new URL("../..", import.meta.url));

const extensionApiEntry = fileURLToPath(new URL("../../packages/extension-api/index.mjs", import.meta.url));
const collabUndoEntry = fileURLToPath(new URL("../../packages/collab/undo/index.js", import.meta.url));
const tauriConfigPath = fileURLToPath(new URL("./src-tauri/tauri.conf.json", import.meta.url));
const tauriCsp = (JSON.parse(readFileSync(tauriConfigPath, "utf8")) as any)?.app?.security?.csp as unknown;
const isE2E = process.env.FORMULA_E2E === "1";
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
  plugins: [resolveJsToTs()],
  resolve: {
    alias: {
      "@formula/extension-api": extensionApiEntry,
      "@formula/collab-undo": collabUndoEntry
    }
  },
  build: {
    // Desktop (Tauri/WebView) targets modern runtimes. Use a modern output target so
    // optional dependencies (e.g. apache-arrow) can rely on top-level await without
    // breaking production builds.
    target: "es2022",
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
      ["src/command-palette/commandPaletteController.vitest.ts", "jsdom"],
    ],
    include: [
      "src/**/*.vitest.ts",
      "src/ai/tools/**/*.test.ts",
      "src/editor/cellEditorOverlay.f4.test.ts",
      "src/ai/inline-edit/__tests__/**/*.test.ts",
    ],
    exclude: ["tests/**", "node_modules/**"],
  },
});
