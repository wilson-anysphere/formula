import { defineConfig } from "@playwright/test";

function withEnv(overrides: Record<string, string>): Record<string, string> {
  const env: Record<string, string> = {};
  for (const [key, value] of Object.entries(process.env)) {
    if (typeof value === "string") {
      env[key] = value;
    }
  }
  return { ...env, ...overrides };
}

const parsedBasePort = Number.parseInt(process.env.FORMULA_E2E_PORT ?? "4174", 10);
const basePort = Number.isFinite(parsedBasePort) ? parsedBasePort : 4174;
const cspPort = basePort + 1;
const desktopPort = basePort + 2;

export default defineConfig({
  testDir: "./tests/e2e",
  timeout: 30_000,
  retries: 0,
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
