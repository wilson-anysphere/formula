import { defineConfig } from "@playwright/test";

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
