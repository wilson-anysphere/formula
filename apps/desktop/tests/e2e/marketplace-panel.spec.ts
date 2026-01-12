import { expect, test } from "@playwright/test";

import { gotoDesktop, waitForDesktopReady } from "./helpers";

test.describe("marketplace panel", () => {
  test("opens from the ribbon (View → Panels)", async ({ page }) => {
    await gotoDesktop(page);
    await waitForDesktopReady(page);

    const viewTab = page.getByRole("tab", { name: "View", exact: true });
    await expect(viewTab).toBeVisible();
    await viewTab.click();

    const openMarketplacePanel = page.getByTestId("open-marketplace-panel");
    await expect(openMarketplacePanel).toBeVisible();
    await openMarketplacePanel.click();

    await expect(page.getByTestId("panel-marketplace")).toBeVisible();
  });
});

