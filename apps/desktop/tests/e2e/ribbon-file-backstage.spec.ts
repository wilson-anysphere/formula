import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("ribbon File backstage", () => {
  test("opens from File tab and supports Escape + focus trapping", async ({ page }) => {
    await gotoDesktop(page);

    const ribbon = page.getByTestId("ribbon-root");
    await expect(ribbon).toBeVisible();

    const fileTab = ribbon.getByRole("tab", { name: "File" });
    await fileTab.click();

    const fileNew = page.getByTestId("file-new");
    const fileQuit = page.getByTestId("file-quit");

    await expect(fileNew).toBeVisible();
    await expect(fileNew).toBeFocused();

    // Focus trap: Shift+Tab wraps back to the last action.
    await page.keyboard.press("Shift+Tab");
    await expect(fileQuit).toBeFocused();

    // Focus trap: Tab from the last action wraps back to the first.
    await page.keyboard.press("Tab");
    await expect(fileNew).toBeFocused();

    for (let i = 0; i < 5; i += 1) {
      await page.keyboard.press("Tab");
    }
    await expect(fileQuit).toBeFocused();

    await page.keyboard.press("Tab");
    await expect(fileNew).toBeFocused();

    // Escape closes the backstage and returns focus to the previous tab.
    await page.keyboard.press("Escape");
    await expect(fileNew).toHaveCount(0);

    const homeTab = ribbon.getByRole("tab", { name: "Home" });
    await expect(homeTab).toHaveAttribute("aria-selected", "true");
  });
});

