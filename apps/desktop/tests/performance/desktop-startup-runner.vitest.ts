import { spawnSync } from "node:child_process";

import { describe, expect, it } from "vitest";

import { repoRoot } from "./desktopStartupUtil.ts";

describe("desktop-startup-runner shell first_render_ms skip policy", () => {
  it("skips first_render_ms percentiles when too few runs report it", () => {
    // This test uses a shebang-executed fixture binary, which is not portable to Windows.
    if (process.platform === "win32") return;

    const binRel = "apps/desktop/tests/performance/fixtures/fakeDesktopStartupBinPartialFirstRender.cjs";

    const env: NodeJS.ProcessEnv = { ...process.env };
    env.FORMULA_DESKTOP_BIN = binRel;
    env.FORMULA_ENFORCE_DESKTOP_STARTUP_BENCH = "0";
    env.DISPLAY = ":99";
    delete env.CI;

    const proc = spawnSync(
      process.execPath,
      [
        "scripts/run-node-ts.mjs",
        "apps/desktop/tests/performance/desktop-startup-runner.ts",
        "--shell",
        "--runs",
        "5",
        "--timeout-ms",
        "5000",
        "--allow-ci",
      ],
      {
        cwd: repoRoot,
        env,
        encoding: "utf8",
        maxBuffer: 5 * 1024 * 1024,
      },
    );

    expect(proc.error).toBeUndefined();
    expect(proc.status).toBe(0);

    const output = `${proc.stdout ?? ""}\n${proc.stderr ?? ""}`;

    expect(output).toContain(
      "[desktop-startup] first_render_ms only available for 1/5 runs (<80%); skipping metric",
    );
    expect(output).toContain("firstRender(p50=n/a,p95=n/a)");
  });
});

