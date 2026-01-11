import { defineConfig } from "@playwright/test";

const port = (() => {
  const raw = process.env.PLAYWRIGHT_WEB_PORT;
  const parsed = raw ? Number.parseInt(raw, 10) : 4173;
  return Number.isFinite(parsed) && parsed > 0 ? parsed : 4173;
})();

export default defineConfig({
  testDir: "./tests/e2e",
  timeout: 30_000,
  retries: 0,
  use: {
    baseURL: `http://localhost:${port}`,
    headless: true
  },
  webServer: {
    command: `pnpm build && pnpm preview --port ${port} --strictPort`,
    port,
    timeout: 1_800_000,
    reuseExistingServer: !process.env.CI
  }
});
