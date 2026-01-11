import { defineConfig, firefox } from "@playwright/test";
import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..", "..");
const cargoHome = path.join(repoRoot, "target", "cargo-home-playwright");

const port = (() => {
  const raw = process.env.PLAYWRIGHT_WEB_PORT;
  const parsed = raw ? Number.parseInt(raw, 10) : 4173;
  return Number.isFinite(parsed) && parsed > 0 ? parsed : 4173;
})();

export default defineConfig({
  testDir: "./tests/e2e",
  timeout: 30_000,
  retries: 0,
  projects: [
    { name: "chromium", use: { browserName: "chromium" } },
    ...(existsSync(firefox.executablePath()) ? [{ name: "firefox", use: { browserName: "firefox" } }] : [])
  ],
  use: {
    baseURL: `http://localhost:${port}`,
    headless: true
  },
  webServer: {
    command: `pnpm build && pnpm preview --port ${port} --strictPort`,
    port,
    timeout: 1_800_000,
    reuseExistingServer: !process.env.CI,
    env: {
      ...process.env,
      // Use a repo-local cargo home to avoid cross-agent contention on ~/.cargo
      // (and to avoid picking up any global cargo config such as rustc-wrapper).
      CARGO_HOME: cargoHome
    }
  }
});
