import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("ribbon zoom", () => {
  test("View → Zoom dropdown items update the grid zoom and status bar selector", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");

    const ribbon = page.getByTestId("ribbon-root");
    await expect(ribbon).toBeVisible();

    const viewTab = ribbon.getByRole("tab", { name: "View", exact: true });
    await viewTab.click();
    await expect(viewTab).toHaveAttribute("aria-selected", "true");

    const zoomControl = page.getByTestId("zoom-control");
    await expect(zoomControl).toHaveValue("100");

    const zoomDropdown = ribbon.locator('button[data-command-id="view.zoom.zoom"]');
    await expect(zoomDropdown).toBeVisible();

    // Menu item zoom presets.
    await zoomDropdown.click();
    const zoom200 = ribbon.locator('button[data-command-id="view.zoom.zoom200"]');
    await expect(zoom200).toBeVisible();
    await zoom200.click();

    await expect.poll(() => page.evaluate(() => (window.__formulaApp as any).getZoom())).toBe(2);
    await expect(zoomControl).toHaveValue("200");

    // Menu item "Custom…" should route to the existing QuickPick zoom flow.
    await zoomDropdown.click();
    const customZoom = ribbon.locator('button[data-command-id="view.zoom.openPicker"]');
    await expect(customZoom).toBeVisible();
    await customZoom.click();

    const quickPick = page.getByTestId("quick-pick");
    await expect(quickPick).toBeVisible();
    await quickPick.getByRole("button", { name: "125%" }).click();

    await page.waitForFunction(() => Math.abs((window.__formulaApp as any)?.getZoom?.() - 1.25) < 0.01);
    await expect(zoomControl).toHaveValue("125");
  });
});
