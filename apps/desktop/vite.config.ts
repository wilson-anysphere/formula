import { fileURLToPath } from "node:url";
import { defineConfig } from "vite";
import { fileURLToPath } from "node:url";

const repoRoot = fileURLToPath(new URL("../..", import.meta.url));

const extensionApiEntry = fileURLToPath(new URL("../../packages/extension-api/index.mjs", import.meta.url));

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
  },
  test: {
    environment: "node",
    environmentMatchGlobs: [["src/panels/ai-audit/AIAuditPanel.vitest.ts", "jsdom"]],
    include: ["src/**/*.vitest.ts"],
    exclude: ["tests/**", "node_modules/**"],
  },
});
