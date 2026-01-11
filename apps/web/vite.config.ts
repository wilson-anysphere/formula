import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";
import { fileURLToPath } from "node:url";

const repoRoot = fileURLToPath(new URL("../..", import.meta.url));

export default defineConfig({
  plugins: [react()],
  build: {
    commonjsOptions: {
      // `shared/` is CommonJS, but the web runtime imports the browser verifier (ESM)
      // which depends on `shared/extension-package/core/v2-core.js`. Ensure Rollup
      // runs the CommonJS transform on that file during production builds.
      include: [
        /node_modules/,
        /shared[\\/]+extension-package[\\/]+core[\\/]+/
      ]
    }
  },
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
