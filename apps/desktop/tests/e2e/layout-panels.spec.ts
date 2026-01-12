import { expect, test } from "@playwright/test";

import { gotoDesktop, waitForDesktopReady } from "./helpers";

test.describe("dockable panels layout persistence", () => {
  test("open AI panel, dock left, reload restores layout", async ({ page }) => {
    await gotoDesktop(page);
    await page.evaluate(() => localStorage.clear());
    await page.reload();
    await waitForDesktopReady(page);

    // Open AI panel (defaults to right dock via panel registry).
    await page.getByTestId("open-ai-panel").click();
    await expect(page.getByTestId("dock-right").getByTestId("panel-aiChat")).toBeVisible();

    // Dock to left.
    await page.getByTestId("dock-ai-panel-left").click();
    await expect(page.getByTestId("dock-left").getByTestId("panel-aiChat")).toBeVisible();
    await expect(page.getByTestId("dock-right").getByTestId("panel-aiChat")).toHaveCount(0);

    // Reload: layout should restore from localStorage.
    await page.reload();
    await waitForDesktopReady(page);

    await expect(page.getByTestId("dock-left").getByTestId("panel-aiChat")).toBeVisible();
    await expect(page.getByTestId("dock-right").getByTestId("panel-aiChat")).toHaveCount(0);
  });

  test("dock tab strip switches between multiple open panels", async ({ page }) => {
    await gotoDesktop(page);
    await page.evaluate(() => localStorage.clear());
    await page.reload();
    await waitForDesktopReady(page);

    // Open AI panel (defaults to right dock).
    await page.getByTestId("open-ai-panel").click();
    const rightDock = page.getByTestId("dock-right");
    await expect(rightDock.getByTestId("panel-aiChat")).toBeVisible();

    // Open another panel in the same dock.
    await page.getByTestId("open-macros-panel").click();
    await expect(rightDock.getByTestId("panel-macros")).toBeVisible();

    // Tab strip should be visible and allow switching between panels.
    await expect(rightDock.getByRole("tablist")).toBeVisible();
    await expect(rightDock.getByTestId("dock-tab-aiChat")).toBeVisible();
    await expect(rightDock.getByTestId("dock-tab-macros")).toBeVisible();

    await rightDock.getByTestId("dock-tab-aiChat").click();
    await expect(rightDock.getByTestId("panel-aiChat")).toBeVisible();
    await expect(rightDock.getByTestId("panel-macros")).toHaveCount(0);

    await rightDock.getByTestId("dock-tab-macros").click();
    await expect(rightDock.getByTestId("panel-macros")).toBeVisible();
    await expect(rightDock.getByTestId("panel-aiChat")).toHaveCount(0);

    // Keyboard navigation (roving tabindex).
    await rightDock.getByTestId("dock-tab-macros").press("Home");
    await expect(rightDock.getByTestId("panel-aiChat")).toBeVisible();
    await expect(rightDock.getByTestId("panel-macros")).toHaveCount(0);

    await rightDock.getByTestId("dock-tab-aiChat").press("End");
    await expect(rightDock.getByTestId("panel-macros")).toBeVisible();
    await expect(rightDock.getByTestId("panel-aiChat")).toHaveCount(0);

    await rightDock.getByTestId("dock-tab-macros").press("ArrowLeft");
    await expect(rightDock.getByTestId("panel-aiChat")).toBeVisible();
    await expect(rightDock.getByTestId("panel-macros")).toHaveCount(0);
  });
});
