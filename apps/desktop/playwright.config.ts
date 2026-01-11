import { defineConfig } from "@playwright/test";

// Playwright sets `FORCE_COLOR` for its reporters, which causes some tooling
// (eg `supports-color`) to warn when `NO_COLOR` is also present. Ensure `NO_COLOR`
// is unset for the test runner + any spawned webServer processes so e2e output
// stays warning-free.
delete process.env.NO_COLOR;

function withEnv(overrides: Record<string, string>): NodeJS.ProcessEnv {
  // Playwright sets `FORCE_COLOR` for its reporters, which causes some tooling
  // (eg `supports-color`) to warn when `NO_COLOR` is also present. Explicitly
  // unset `NO_COLOR` for the webServer environment to keep e2e output clean.
  const env: NodeJS.ProcessEnv = { ...process.env, ...overrides };
  env.NO_COLOR = undefined;
  return env;
}

const parsedBasePort = Number.parseInt(process.env.FORMULA_E2E_PORT ?? "4174", 10);
const basePort = Number.isFinite(parsedBasePort) ? parsedBasePort : 4174;
const cspPort = basePort + 1;
const desktopPort = basePort + 2;

const parsedWorkers = Number.parseInt(process.env.FORMULA_E2E_WORKERS ?? "2", 10);
// The desktop e2e suite spins up additional workers for WASM/python/script runtimes. Running
// too many Playwright workers in parallel can starve the dev server and lead to timeouts.
// Default to a conservative value and allow overrides via FORMULA_E2E_WORKERS.
const workers = Number.isFinite(parsedWorkers) ? Math.max(1, parsedWorkers) : 2;

export default defineConfig({
  testDir: "./tests/e2e",
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
