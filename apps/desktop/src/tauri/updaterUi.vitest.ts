import { afterEach, describe, expect, it, vi } from "vitest";

import { FORMULA_RELEASES_URL, openUpdateReleasePage } from "./updaterUi";

describe("tauri/updaterUi", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("opens the release URL from update metadata when provided", async () => {
    const invoke = vi.fn(async () => undefined);
    vi.stubGlobal("__TAURI__", { core: { invoke } });

    await openUpdateReleasePage({ releaseUrl: "https://example.com/releases/v1.2.3" });

    expect(invoke).toHaveBeenCalledWith("open_external_url", { url: "https://example.com/releases/v1.2.3" });
  });

  it("falls back to the repo releases page when no release URL metadata is available", async () => {
    const invoke = vi.fn(async () => undefined);
    vi.stubGlobal("__TAURI__", { core: { invoke } });

    await openUpdateReleasePage({ version: "1.2.3", body: "Notes" });

    expect(invoke).toHaveBeenCalledWith("open_external_url", { url: FORMULA_RELEASES_URL });
  });
});
