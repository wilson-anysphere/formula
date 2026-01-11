import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    // The repo includes several integration-style suites (API, sandboxed runtimes,
    // wasm-backed rendering) that can exceed Vitest's default 10s hook timeout on
    // shared/contended runners.
    testTimeout: 30_000,
    hookTimeout: 30_000,
    include: [
      "packages/**/*.test.ts",
      "packages/**/*.test.tsx",
      "apps/**/*.test.ts",
      "apps/**/*.test.tsx",
      "services/api/src/__tests__/**/*.test.ts"
    ],
    environment: "node",
    setupFiles: ["./vitest.setup.ts"],
    globalSetup: "./scripts/vitest.global-setup.mjs"
  }
});
