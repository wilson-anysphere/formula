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

export default defineConfig({
  testDir: "./tests/e2e",
  timeout: 30_000,
  retries: 0,
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
      use: { baseURL: "http://localhost:4176" }
    },
    {
      name: "csp",
      // Run CSP/WASM/Worker checks against a dedicated server instance that
      // injects the Tauri CSP header via `FORMULA_E2E=1`.
      testMatch: ["csp.spec.ts"],
      use: { baseURL: "http://localhost:4175" }
    }
  ],
  webServer: [
    {
      command: "pnpm exec vite --port 4176 --strictPort",
      port: 4176,
      reuseExistingServer: false
    },
    {
      command: "pnpm exec vite --port 4175 --strictPort",
      port: 4175,
      reuseExistingServer: false,
      env: withEnv({ FORMULA_E2E: "1" })
    }
  ]
});
