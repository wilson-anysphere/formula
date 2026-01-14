import { defineConfig, firefox } from "@playwright/test";
import { existsSync } from "node:fs";
import { homedir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

// Playwright sets `FORCE_COLOR` for its reporters, which causes some tooling
// (eg `supports-color`) to warn when `NO_COLOR` is also present. Ensure `NO_COLOR`
// is unset for the test runner + any spawned webServer processes so e2e output
// stays warning-free.
delete process.env.NO_COLOR;

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..", "..");
const defaultGlobalCargoHome = path.resolve(homedir(), ".cargo");
const envCargoHome = process.env.CARGO_HOME;
const normalizedEnvCargoHome = envCargoHome ? path.resolve(envCargoHome) : null;
const cargoHome =
  !envCargoHome ||
  (!process.env.CI &&
    !process.env.FORMULA_ALLOW_GLOBAL_CARGO_HOME &&
    normalizedEnvCargoHome === defaultGlobalCargoHome)
    ? path.join(repoRoot, "target", "cargo-home-playwright")
    : envCargoHome!;

function stablePortFromString(input: string, { base = 4173, range = 10_000 } = {}): number {
  // Deterministic port selection avoids collisions when multiple agents run Playwright tests
  // on the same host. `repoRoot` is unique per checkout in our agent environment.
  let hash = 0;
  for (let i = 0; i < input.length; i++) {
    hash = (hash * 31 + input.charCodeAt(i)) >>> 0;
  }
  return base + (hash % range);
}

const defaultPort = 4173;
const maxPort = 65535;
const port = (() => {
  const raw = process.env.PW_WEB_PORT ?? process.env.PLAYWRIGHT_WEB_PORT ?? process.env.PLAYWRIGHT_PORT;
  const parsed = raw ? Number.parseInt(raw, 10) : NaN;
  if (Number.isFinite(parsed) && parsed > 0) {
    if (parsed > maxPort) {
      throw new Error(
        `PW_WEB_PORT/PLAYWRIGHT_WEB_PORT/PLAYWRIGHT_PORT must be between 1 and ${maxPort} (got ${parsed}).`
      );
    }
    return parsed;
  }
  if (process.env.CI) return defaultPort;
  return stablePortFromString(repoRoot, { base: defaultPort, range: 10_000 });
})();
const baseURL = process.env.PW_BASE_URL ?? `http://localhost:${port}`;

export default defineConfig({
  testDir: "./tests/e2e",
  timeout: 30_000,
  retries: 0,
  projects: [
    { name: "chromium", use: { browserName: "chromium" } },
    ...(existsSync(firefox.executablePath()) ? [{ name: "firefox", use: { browserName: "firefox" } }] : [])
  ],
  use: {
    baseURL,
    headless: true
  },
  webServer: {
    command: `pnpm build && pnpm preview --port ${port} --strictPort`,
    port,
    timeout: 1_800_000,
    reuseExistingServer: !process.env.CI,
    env: {
      ...process.env,
      NO_COLOR: undefined,
      // Use a repo-local cargo home to avoid cross-agent contention on ~/.cargo
      // (and to avoid picking up any global cargo config such as rustc-wrapper).
      CARGO_HOME: cargoHome
    }
  }
});
