import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    // Only run the API's Vitest suites. This repo also contains some integration
    // tests written with Node's built-in `node:test` runner under `*/__tests__`.
    // Vitest will execute those files but report "No test suite found", causing
    // the run to fail.
    include: ["src/__tests__/**/*.test.ts"],
    // This service includes several e2e-style integration tests that can take
    // longer than Vitest's default 5s timeout, especially on shared CI runners.
    testTimeout: 20_000,
    hookTimeout: 20_000,
    environment: "node"
  }
});
