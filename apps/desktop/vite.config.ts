import { existsSync, readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { defineConfig } from "vite";

const repoRoot = fileURLToPath(new URL("../..", import.meta.url));

const extensionApiEntry = fileURLToPath(new URL("../../packages/extension-api/index.mjs", import.meta.url));
const tauriConfigPath = fileURLToPath(new URL("./src-tauri/tauri.conf.json", import.meta.url));
const tauriCsp = (JSON.parse(readFileSync(tauriConfigPath, "utf8")) as any)?.app?.security?.csp as unknown;
const isE2E = process.env.FORMULA_E2E === "1";

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
  plugins: [resolveJsToTs()],
  resolve: {
    alias: {
      "@formula/extension-api": extensionApiEntry
    }
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
    environmentMatchGlobs: [["src/panels/ai-audit/AIAuditPanel.vitest.ts", "jsdom"]],
    include: ["src/**/*.vitest.ts"],
    exclude: ["tests/**", "node_modules/**"],
  },
});
