import { expect, test } from "@playwright/test";

import { gotoDesktop, waitForDesktopReady } from "./helpers";

test.describe("ribbon shell smoke", () => {
  test("renders titlebar + ribbon and hosts key controls", async ({ page }) => {
    await gotoDesktop(page);
    await waitForDesktopReady(page);

    await expect(page.getByTestId("titlebar")).toBeVisible();

    const ribbon = page.getByTestId("ribbon");
    await expect(ribbon).toBeVisible();

    // Ensure the AI panel toggle has been migrated into the ribbon container.
    const openAiPanel = ribbon.getByTestId("open-ai-panel");
    await expect(openAiPanel).toBeVisible();
    await openAiPanel.click();
    await expect(page.getByTestId("dock-right").getByTestId("panel-aiChat")).toBeVisible();
  });
});

