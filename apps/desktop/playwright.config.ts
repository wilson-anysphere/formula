import { defineConfig } from "@playwright/test";

// Ensure Vite can detect e2e runs cross-platform without relying on shell-specific
// env var assignment syntax in the `webServer.command`.
if (!process.env.FORMULA_E2E) {
  process.env.FORMULA_E2E = "1";
}

export default defineConfig({
  testDir: "./tests/e2e",
  timeout: 30_000,
  retries: 0,
  use: {
    baseURL: "http://localhost:4174",
    headless: true
  },
  webServer: {
    command: "pnpm dev",
    port: 4174,
    reuseExistingServer: !process.env.CI
  }
});
