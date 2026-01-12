import { afterEach, describe, expect, it } from "vitest";

import { resolveDesktopGridMode } from "../desktopGridMode";

type EnvSnapshot = {
  node: Record<string, string | undefined>;
  meta: {
    env: any;
    values: Record<string, unknown>;
  } | null;
};

let snapshot: EnvSnapshot | null = null;

function snapshotEnv(): EnvSnapshot {
  const node = {
    DESKTOP_GRID_MODE: process.env.DESKTOP_GRID_MODE,
    GRID_MODE: process.env.GRID_MODE,
    USE_SHARED_GRID: process.env.USE_SHARED_GRID,
  };

  const metaEnv = (import.meta as any)?.env as Record<string, unknown> | undefined;
  if (!metaEnv) return { node, meta: null };

  const keys = ["VITE_DESKTOP_GRID_MODE", "VITE_GRID_MODE", "VITE_USE_SHARED_GRID"] as const;
  const values: Record<string, unknown> = {};
  for (const key of keys) values[key] = metaEnv[key];
  return { node, meta: { env: metaEnv, values } };
}

function restoreEnv(state: EnvSnapshot): void {
  for (const [key, value] of Object.entries(state.node)) {
    if (value === undefined) delete (process.env as any)[key];
    else (process.env as any)[key] = value;
  }

  if (state.meta) {
    for (const [key, value] of Object.entries(state.meta.values)) {
      if (value === undefined) delete state.meta.env[key];
      else state.meta.env[key] = value;
    }
  }
}

function clearGridModeEnv(): void {
  delete process.env.DESKTOP_GRID_MODE;
  delete process.env.GRID_MODE;
  delete process.env.USE_SHARED_GRID;

  const metaEnv = (import.meta as any)?.env as Record<string, unknown> | undefined;
  if (metaEnv) {
    delete metaEnv.VITE_DESKTOP_GRID_MODE;
    delete metaEnv.VITE_GRID_MODE;
    delete metaEnv.VITE_USE_SHARED_GRID;
  }
}

describe("resolveDesktopGridMode", () => {
  afterEach(() => {
    if (snapshot) restoreEnv(snapshot);
    snapshot = null;
  });

  it("defaults to shared when there are no query/env overrides", () => {
    snapshot = snapshotEnv();
    clearGridModeEnv();

    expect(resolveDesktopGridMode("")).toBe("shared");
  });

  it("honors env overrides", () => {
    snapshot = snapshotEnv();
    clearGridModeEnv();

    process.env.DESKTOP_GRID_MODE = "legacy";
    expect(resolveDesktopGridMode("")).toBe("legacy");

    process.env.DESKTOP_GRID_MODE = "shared";
    expect(resolveDesktopGridMode("")).toBe("shared");
  });

  it("honors query string overrides over env", () => {
    snapshot = snapshotEnv();
    clearGridModeEnv();

    process.env.DESKTOP_GRID_MODE = "shared";
    expect(resolveDesktopGridMode("?grid=legacy")).toBe("legacy");

    process.env.DESKTOP_GRID_MODE = "legacy";
    expect(resolveDesktopGridMode("?grid=shared")).toBe("shared");
  });
});

