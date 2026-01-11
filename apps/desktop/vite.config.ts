import { fileURLToPath } from "node:url";
import { defineConfig } from "vite";

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
    strictPort: true
  },
  test: {
    environment: "node",
    include: ["src/**/*.vitest.ts"],
    exclude: ["tests/**", "node_modules/**"]
  }
});
