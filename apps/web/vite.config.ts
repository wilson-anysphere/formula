import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";
import { fileURLToPath } from "node:url";

const repoRoot = fileURLToPath(new URL("../..", import.meta.url));

export default defineConfig({
  plugins: [react()],
  optimizeDeps: {
    // Ensure Vite transforms `new Worker(new URL(..., import.meta.url))` inside the
    // workspace engine package. When pre-bundled, the Worker URL can become
    // relative to the optimized dep chunk and fail to load.
    exclude: ["@formula/engine"],
  },
  server: {
    fs: {
      // Allow serving workspace packages during dev (`packages/*`).
      allow: [repoRoot],
    },
  },
});
