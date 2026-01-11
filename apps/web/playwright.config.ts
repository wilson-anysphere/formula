import { defineConfig, firefox } from "@playwright/test";
import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..", "..");
const cargoHome = process.env.CARGO_HOME ?? path.join(repoRoot, "target", "cargo-home-playwright");

function stablePortFromString(input: string, { base = 4173, range = 1000 } = {}): number {
  // Deterministic port selection avoids collisions when multiple agents run Playwright tests
  // on the same host. `repoRoot` is unique per checkout in our agent environment.
  let hash = 0;
  for (let i = 0; i < input.length; i++) {
    hash = (hash * 31 + input.charCodeAt(i)) >>> 0;
  }
  return base + (hash % range);
}

const defaultPort = 4173;
const port = (() => {
  const raw = process.env.PLAYWRIGHT_WEB_PORT ?? process.env.PLAYWRIGHT_PORT;
  const parsed = raw ? Number.parseInt(raw, 10) : NaN;
  if (Number.isFinite(parsed) && parsed > 0) return parsed;
  if (process.env.CI) return defaultPort;
  return stablePortFromString(repoRoot, { base: defaultPort, range: 1000 });
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
