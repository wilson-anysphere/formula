/**
 * @vitest-environment jsdom
 */

import { afterEach, describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => {
  return {
    shellOpen: vi.fn<[], Promise<void>>().mockResolvedValue(undefined),
  };
});

vi.mock("../shellOpen", () => ({
  shellOpen: mocks.shellOpen,
}));

describe("updaterUi dialog", () => {
  afterEach(() => {
    mocks.shellOpen.mockClear();
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
    document.body.replaceChildren();
  });

  it("opens the GitHub Releases page from the update dialog", async () => {
    document.body.innerHTML = `<div id="toast-root"></div>`;

    const { handleUpdaterEvent, FORMULA_RELEASES_URL } = await import("../updaterUi");

    await handleUpdaterEvent("update-available", { source: "manual", version: "1.2.3", body: "Notes" });

    const dialog = document.querySelector('[data-testid="updater-dialog"]') as HTMLElement | null;
    expect(dialog).not.toBeNull();

    const viewBtn = dialog?.querySelector<HTMLButtonElement>('[data-testid="updater-view-versions"]');
    expect(viewBtn).not.toBeNull();

    viewBtn?.click();

    expect(mocks.shellOpen).toHaveBeenCalledTimes(1);
    expect(mocks.shellOpen).toHaveBeenCalledWith(FORMULA_RELEASES_URL);
  });
});
