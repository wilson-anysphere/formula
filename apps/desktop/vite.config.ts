import { readFileSync } from "node:fs";
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

export default defineConfig({
  root: ".",
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
    ...(isE2E
      ? {
          // E2E runs in a real browser, so we emulate the Tauri CSP via HTTP headers
          // to catch regressions (e.g. WASM compilation or Worker bootstrapping being blocked).
          headers: { "Content-Security-Policy": tauriCsp as string },
          // Avoid Vite HMR WebSocket noise in CSP checks.
          hmr: false
        }
      : {})
  },
  test: {
    environment: "node",
    environmentMatchGlobs: [["src/panels/ai-audit/AIAuditPanel.vitest.ts", "jsdom"]],
    include: ["src/**/*.vitest.ts"],
    exclude: ["tests/**", "node_modules/**"],
  },
});
