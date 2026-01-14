import { defineConfig } from "@playwright/test";
import path from "node:path";
import { fileURLToPath } from "node:url";

// Playwright sets `FORCE_COLOR` for its reporters, which causes some tooling
// (eg `supports-color`) to warn when `NO_COLOR` is also present. Ensure `NO_COLOR`
// is unset for the test runner + any spawned webServer processes so e2e output
// stays warning-free.
delete process.env.NO_COLOR;

const repoRoot = path.resolve(fileURLToPath(new URL(".", import.meta.url)), "../..");

function stablePortFromString(input: string, { base = 4174, range = 1000 } = {}): number {
  // Deterministic port selection avoids collisions when multiple agents run Playwright tests
  // on the same host. `repoRoot` is unique per checkout in our agent environment.
  let hash = 0;
  for (let i = 0; i < input.length; i++) {
    hash = (hash * 31 + input.charCodeAt(i)) >>> 0;
  }
  return base + (hash % range);
}

function withEnv(overrides: Record<string, string>): NodeJS.ProcessEnv {
  // Playwright sets `FORCE_COLOR` for its reporters, which causes some tooling
  // (eg `supports-color`) to warn when `NO_COLOR` is also present. Explicitly
  // unset `NO_COLOR` for the webServer environment to keep e2e output clean.
  const env: NodeJS.ProcessEnv = { ...process.env, ...overrides };
  env.NO_COLOR = undefined;
  return env;
}

const defaultBasePort = 4174;
const maxBasePort = 65533;
const basePort = (() => {
  const raw = process.env.FORMULA_E2E_PORT;
  const parsed = raw ? Number.parseInt(raw, 10) : NaN;
  if (Number.isFinite(parsed) && parsed > 0) {
    if (parsed > maxBasePort) {
      throw new Error(
        `FORMULA_E2E_PORT must be between 1 and ${maxBasePort} (got ${parsed}). ` +
          `Playwright reserves basePort+1/basePort+2 for the CSP + desktop servers.`
      );
    }
    return parsed;
  }
  if (process.env.CI) return defaultBasePort;
  // Use a wide range so developers are less likely to collide with a concurrently running
  // `pnpm dev` (fixed port) or another local service.
  return stablePortFromString(repoRoot, { base: defaultBasePort, range: 20_000 });
})();
const cspPort = basePort + 1;
const desktopPort = basePort + 2;

const parsedWorkers = Number.parseInt(process.env.FORMULA_E2E_WORKERS ?? "2", 10);
// The desktop e2e suite spins up additional workers for WASM/python/script runtimes. Running
// too many Playwright workers in parallel can starve the dev server and lead to timeouts.
// Default to a conservative value and allow overrides via FORMULA_E2E_WORKERS.
const workers = Number.isFinite(parsedWorkers) ? Math.max(1, parsedWorkers) : 2;

export default defineConfig({
  testDir: "./tests/e2e",
  // Store `expect(...).toHaveScreenshot(...)` baselines in a stable, flat location that can be
  // checked into git alongside the e2e specs.
  snapshotPathTemplate: "{testDir}/__screenshots__/{testFilePath}/{arg}{ext}",
  // First-run Vite dependency optimization (and WASM/python worker boot) can exceed the default
  // Playwright timeout under heavy CI load. Use a slightly more forgiving default; individual
  // tests still override this when they need longer.
  timeout: 60_000,
  retries: 0,
  workers,
  globalSetup: "./tests/e2e/global-setup.ts",
  use: {
    headless: true
  },
  projects: [
    {
      name: "desktop",
      // Run the regular desktop e2e suite against a vanilla Vite dev server.
      // (No CSP header emulation here so existing network-related tests can
      // exercise the permission gating layer.)
      testIgnore: ["csp.spec.ts"],
      use: { baseURL: `http://localhost:${desktopPort}` }
    },
    {
      name: "csp",
      // Run CSP/WASM/Worker checks against a dedicated server instance that
      // injects the Tauri CSP header via `FORMULA_E2E=1`.
      testMatch: ["csp.spec.ts"],
      use: { baseURL: `http://localhost:${cspPort}` }
    }
  ],
  webServer: [
    {
      // The python runtime e2e tests expect the Pyodide distribution to be
      // self-hosted under `apps/desktop/public/pyodide/...`. The assets are
      // intentionally gitignored, so ensure they exist before starting Vite.
      command: `node scripts/ensure-pyodide-assets.mjs && pnpm exec vite --port ${desktopPort} --strictPort`,
      port: desktopPort,
      reuseExistingServer: false,
      env: withEnv({ FORMULA_E2E: "0" })
    },
    {
      command: `pnpm exec vite --port ${cspPort} --strictPort`,
      port: cspPort,
      reuseExistingServer: false,
      env: withEnv({ FORMULA_E2E: "1" })
    }
  ]
});
