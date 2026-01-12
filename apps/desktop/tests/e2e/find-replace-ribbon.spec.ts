import { expect, test } from "@playwright/test";

import { gotoDesktop, waitForDesktopReady } from "./helpers";

test.describe("ribbon Find & Select", () => {
  test("opens Find/Replace/Go To dialogs", async ({ page }) => {
    await gotoDesktop(page);
    await waitForDesktopReady(page);

    // Desktop currently defaults to the View tab (where debug controls live). Switch to Home
    // so we can access the Find & Select dropdown.
    await page.getByRole("tab", { name: "Home" }).click();

    const findSelect = page.getByRole("button", { name: "Find and Select" });
    await expect(findSelect).toBeVisible();

    // --- Find ---
    await findSelect.click();
    await page.getByTestId("ribbon-root").getByTestId("ribbon-find").click();
    const findDialog = page.locator("dialog.find-replace-dialog[open]");
    await expect(findDialog).toBeVisible();
    await expect(findDialog.locator("input").first()).toBeFocused();
    await page.keyboard.press("Escape");
    await expect(page.locator("dialog.find-replace-dialog[open]")).toHaveCount(0);

    // --- Replace ---
    await findSelect.click();
    await page.getByTestId("ribbon-root").getByTestId("ribbon-replace").click();
    const replaceDialog = page.locator("dialog.find-replace-dialog[open]");
    await expect(replaceDialog).toBeVisible();
    await expect(replaceDialog.locator('input[placeholder="Replace with…"]')).toBeVisible();
    await expect(replaceDialog.locator("input").first()).toBeFocused();
    await page.keyboard.press("Escape");
    await expect(page.locator("dialog.find-replace-dialog[open]")).toHaveCount(0);

    // --- Go To ---
    await findSelect.click();
    await page.getByTestId("ribbon-root").getByTestId("ribbon-goto").click();
    const goToDialog = page.locator("dialog.goto-dialog[open]");
    await expect(goToDialog).toBeVisible();
    await expect(goToDialog.locator("input").first()).toBeFocused();
    await page.keyboard.press("Escape");
    await expect(page.locator("dialog.goto-dialog[open]")).toHaveCount(0);
  });
});
