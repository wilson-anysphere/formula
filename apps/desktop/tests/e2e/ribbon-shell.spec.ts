import { expect, test } from "@playwright/test";

import { gotoDesktop, waitForDesktopReady } from "./helpers";

test.describe("ribbon shell smoke", () => {
  test("renders titlebar + ribbon and hosts key controls", async ({ page }) => {
    await gotoDesktop(page);
    await waitForDesktopReady(page);

    await expect(page.locator("#titlebar .formula-titlebar")).toBeVisible();

    const ribbon = page.getByTestId("ribbon-root");
    await expect(ribbon).toBeVisible();

    // Ensure the AI panel toggle is available in the desktop shell.
    const toggleAiChatPanel = page.getByTestId("open-panel-ai-chat");
    await expect(toggleAiChatPanel).toBeVisible();
    await toggleAiChatPanel.click();
    await expect(page.getByTestId("dock-right").getByTestId("panel-aiChat")).toBeVisible();
  });
});
