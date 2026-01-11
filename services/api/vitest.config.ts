import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    // Only run the API's Vitest suites. This repo also contains some integration
    // tests written with Node's built-in `node:test` runner under `*/__tests__`.
    // Vitest will execute those files but report "No test suite found", causing
    // the run to fail.
    include: ["src/__tests__/**/*.test.ts"],
    environment: "node"
  }
});
