import { expect, test } from "@playwright/test";

import { gotoDesktop, waitForDesktopReady } from "./helpers";

test.describe("dockable panels layout persistence", () => {
  test("open AI panel, dock left, reload restores layout", async ({ page }) => {
    await gotoDesktop(page);
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);

    const ribbon = page.getByTestId("ribbon-root");

    // Open AI panel (defaults to right dock via panel registry).
    await ribbon.getByTestId("open-panel-ai-chat").click();
    await expect(page.getByTestId("dock-right").getByTestId("panel-aiChat")).toBeVisible();

    // Dock to left.
    await page.getByTestId("dock-ai-panel-left").click();
    await expect(page.getByTestId("dock-left").getByTestId("panel-aiChat")).toBeVisible();
    await expect(page.getByTestId("dock-right").getByTestId("panel-aiChat")).toHaveCount(0);

    // Reload: layout should restore from localStorage.
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);

    await expect(page.getByTestId("dock-left").getByTestId("panel-aiChat")).toBeVisible();
    await expect(page.getByTestId("dock-right").getByTestId("panel-aiChat")).toHaveCount(0);
  });

  test("different docId values isolate persisted layout", async ({ page }) => {
    await gotoDesktop(page, "/?docId=doc-a");
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);

    const ribbon = page.getByTestId("ribbon-root");

    // Persist a non-default layout under doc-a.
    await ribbon.getByTestId("open-panel-ai-chat").click();
    await page.getByTestId("dock-ai-panel-left").click();
    await expect(page.getByTestId("dock-left").getByTestId("panel-aiChat")).toBeVisible();

    // Load a different collab document; it should not pick up doc-a's layout.
    await gotoDesktop(page, "/?docId=doc-b");
    await expect(page.getByTestId("panel-aiChat")).toHaveCount(0);

    // Reload doc-b to ensure it remains isolated.
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await expect(page.getByTestId("panel-aiChat")).toHaveCount(0);
  });

  test("dock tab strip switches between multiple open panels", async ({ page }) => {
    await gotoDesktop(page);
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);

    const ribbon = page.getByTestId("ribbon-root");

    // Open AI panel (defaults to right dock).
    await ribbon.getByTestId("open-panel-ai-chat").click();
    const rightDock = page.getByTestId("dock-right");
    await expect(rightDock.getByTestId("panel-aiChat")).toBeVisible();

    // Open another panel in the same dock.
    await ribbon.getByTestId("open-macros-panel").click();
    await expect(rightDock.getByTestId("panel-macros")).toBeVisible();

    // Tab strip should be visible and allow switching between panels.
    await expect(rightDock.getByRole("tablist")).toBeVisible();
    await expect(rightDock.getByRole("tablist")).toHaveAttribute("aria-label", "Docked panels");
    await expect(rightDock.getByTestId("dock-tab-aiChat")).toBeVisible();
    await expect(rightDock.getByTestId("dock-tab-macros")).toBeVisible();
    await expect(rightDock.getByTestId("dock-tab-macros")).toHaveAttribute("aria-selected", "true");
    await expect(rightDock.getByTestId("dock-tab-aiChat")).toHaveAttribute("aria-selected", "false");

    await rightDock.getByTestId("dock-tab-aiChat").click();
    await expect(rightDock.getByTestId("panel-aiChat")).toBeVisible();
    await expect(rightDock.getByTestId("panel-macros")).toHaveCount(0);
    await expect(rightDock.getByTestId("dock-tab-aiChat")).toHaveAttribute("aria-selected", "true");
    await expect(rightDock.getByTestId("dock-tab-macros")).toHaveAttribute("aria-selected", "false");

    await rightDock.getByTestId("dock-tab-macros").click();
    await expect(rightDock.getByTestId("panel-macros")).toBeVisible();
    await expect(rightDock.getByTestId("panel-aiChat")).toHaveCount(0);
    await expect(rightDock.getByTestId("dock-tab-macros")).toHaveAttribute("aria-selected", "true");
    await expect(rightDock.getByTestId("dock-tab-aiChat")).toHaveAttribute("aria-selected", "false");

    // Keyboard navigation (roving tabindex).
    await rightDock.getByTestId("dock-tab-macros").press("Home");
    await expect(rightDock.getByTestId("panel-aiChat")).toBeVisible();
    await expect(rightDock.getByTestId("panel-macros")).toHaveCount(0);
    await expect(rightDock.getByTestId("dock-tab-aiChat")).toHaveAttribute("aria-selected", "true");

    await rightDock.getByTestId("dock-tab-aiChat").press("End");
    await expect(rightDock.getByTestId("panel-macros")).toBeVisible();
    await expect(rightDock.getByTestId("panel-aiChat")).toHaveCount(0);
    await expect(rightDock.getByTestId("dock-tab-macros")).toHaveAttribute("aria-selected", "true");

    await rightDock.getByTestId("dock-tab-macros").press("ArrowLeft");
    await expect(rightDock.getByTestId("panel-aiChat")).toBeVisible();
    await expect(rightDock.getByTestId("panel-macros")).toHaveCount(0);
    await expect(rightDock.getByTestId("dock-tab-aiChat")).toHaveAttribute("aria-selected", "true");
  });

  test("ribbon button opens Version History panel", async ({ page }) => {
    await gotoDesktop(page);
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);

    await page.getByRole("tab", { name: "View", exact: true }).click();
    await page.getByTestId("ribbon-root").getByTestId("open-version-history-panel").click();
    await expect(
      page.locator(
        "[data-testid='dock-left'] [data-testid='panel-versionHistory'], [data-testid='dock-right'] [data-testid='panel-versionHistory'], [data-testid='dock-bottom'] [data-testid='panel-versionHistory']",
      ),
    ).toBeVisible();

    // Toggle again closes it.
    await page.getByTestId("ribbon-root").getByTestId("open-version-history-panel").click();
    await expect(page.getByTestId("panel-versionHistory")).toHaveCount(0);
  });

  test("ribbon button opens Branch Manager panel", async ({ page }) => {
    await gotoDesktop(page);
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);

    await page.getByRole("tab", { name: "View", exact: true }).click();
    await page.getByTestId("ribbon-root").getByTestId("open-branch-manager-panel").click();
    await expect(
      page.locator(
        "[data-testid='dock-left'] [data-testid='panel-branchManager'], [data-testid='dock-right'] [data-testid='panel-branchManager'], [data-testid='dock-bottom'] [data-testid='panel-branchManager']",
      ),
    ).toBeVisible();

    // Toggle again closes it.
    await page.getByTestId("ribbon-root").getByTestId("open-branch-manager-panel").click();
    await expect(page.getByTestId("panel-branchManager")).toHaveCount(0);
  });

  test("status bar buttons toggle Version History + Branch Manager panels", async ({ page }) => {
    await gotoDesktop(page);
    await page.evaluate(() => localStorage.clear());
    await page.reload();
    await waitForDesktopReady(page);

    const statusbar = page.locator(".statusbar__main");

    await statusbar.getByTestId("open-version-history-panel").click();
    await expect(page.getByTestId("dock-right").getByTestId("panel-versionHistory")).toBeVisible();

    await statusbar.getByTestId("open-version-history-panel").click();
    await expect(page.getByTestId("panel-versionHistory")).toHaveCount(0);

    await statusbar.getByTestId("open-branch-manager-panel").click();
    await expect(page.getByTestId("dock-right").getByTestId("panel-branchManager")).toBeVisible();

    await statusbar.getByTestId("open-branch-manager-panel").click();
    await expect(page.getByTestId("panel-branchManager")).toHaveCount(0);
  });

  test("Cmd+Shift+A toggles AI chat panel open/closed", async ({ page }) => {
    await gotoDesktop(page);
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);

    // Ensure focus is on the grid (not an input) so the global shortcut should fire.
    // Avoid clicking the shared-grid corner header (select-all), which can be slow/flaky under Playwright.
    await page.locator("#grid").focus();

    await page.keyboard.press("Meta+Shift+A");
    await expect(page.getByTestId("dock-right").getByTestId("panel-aiChat")).toBeVisible();

    await page.keyboard.press("Meta+Shift+A");
    await expect(page.getByTestId("panel-aiChat")).toHaveCount(0);
  });

  test("Cmd+Shift+A does not toggle AI chat while typing in the formula bar", async ({ page }) => {
    await gotoDesktop(page);
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);

    // Enter formula-bar edit mode (this reveals + focuses the textarea).
    await page.getByTestId("formula-highlight").click();
    await expect(page.getByTestId("formula-input")).toBeVisible();
    await expect(page.getByTestId("formula-input")).toBeFocused();

    await page.keyboard.press("Meta+Shift+A");
    await expect(page.getByTestId("panel-aiChat")).toHaveCount(0);
  });
});
