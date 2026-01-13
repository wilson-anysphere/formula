import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("ribbon File backstage", () => {
  test("opens from File tab and supports Escape + focus trapping", async ({ page }) => {
    await gotoDesktop(page);

    const ribbon = page.getByTestId("ribbon-root");
    await expect(ribbon).toBeVisible();

    const fileTab = ribbon.getByRole("tab", { name: "File" });
    await fileTab.click();

    const fileNew = ribbon.getByTestId("file-new");
    const fileQuit = ribbon.getByTestId("file-quit");

    await expect(fileNew).toBeVisible();
    await expect(fileNew).toBeFocused();
    const expectedShortcut = process.platform === "darwin" ? "Meta+N" : "Control+N";
    await expect(fileNew).toHaveAttribute("aria-keyshortcuts", expectedShortcut);

    // Focus trap: Shift+Tab wraps back to the last action.
    await page.keyboard.press("Shift+Tab");
    await expect(fileQuit).toBeFocused();

    // Focus trap: Tab from the last action wraps back to the first.
    await page.keyboard.press("Tab");
    await expect(fileNew).toBeFocused();

    // Arrow navigation should move focus between menuitems.
    await page.keyboard.press("ArrowDown");
    await expect(ribbon.getByTestId("file-open")).toBeFocused();
    await page.keyboard.press("ArrowUp");
    await expect(fileNew).toBeFocused();

    // Web demo behavior: actions should surface a toast indicating the feature is
    // only available in the desktop app.
    await ribbon.getByTestId("file-open").click();
    const toast = page.getByTestId("toast").last();
    await expect(toast).toContainText(/desktop/i);

    // Re-open and ensure Escape still closes the backstage and returns to the Home tab.
    await fileTab.click();
    await expect(fileNew).toBeVisible();
    await expect(fileNew).toBeFocused();

    // Tab through all enabled backstage actions to reach the last item.
    const backstageItems = page
      .locator(".ribbon-backstage")
      .locator('[role="menuitem"]:not([disabled]), [role="menuitemcheckbox"]:not([disabled])');
    const itemCount = await backstageItems.count();
    for (let i = 0; i < Math.max(0, itemCount - 1); i += 1) await page.keyboard.press("Tab");
    await expect(fileQuit).toBeFocused();

    await page.keyboard.press("Tab");
    await expect(fileNew).toBeFocused();

    // Escape closes the backstage and returns focus to the previous tab.
    await page.keyboard.press("Escape");
    await expect(fileNew).toHaveCount(0);

    const homeTab = ribbon.getByRole("tab", { name: "Home" });
    await expect(homeTab).toHaveAttribute("aria-selected", "true");
  });

  test("can toggle Version History + Branch Manager panels from the backstage actions", async ({ page }) => {
    await gotoDesktop(page);

    const ribbon = page.getByTestId("ribbon-root");
    await expect(ribbon).toBeVisible();

    const fileTab = ribbon.getByRole("tab", { name: "File" });

    // Version History: open
    await fileTab.click();
    await expect(ribbon.getByTestId("file-version-history")).toBeVisible();
    await ribbon.getByTestId("file-version-history").click();
    await expect(page.getByTestId("panel-versionHistory")).toBeVisible();

    // Version History: close (toggle)
    await fileTab.click();
    await expect(ribbon.getByTestId("file-version-history")).toBeVisible();
    await ribbon.getByTestId("file-version-history").click();
    await expect(page.getByTestId("panel-versionHistory")).toHaveCount(0);

    // Branch Manager: open
    await fileTab.click();
    await expect(ribbon.getByTestId("file-branch-manager")).toBeVisible();
    await ribbon.getByTestId("file-branch-manager").click();
    await expect(page.getByTestId("panel-branchManager")).toBeVisible();

    // Branch Manager: close (toggle)
    await fileTab.click();
    await expect(ribbon.getByTestId("file-branch-manager")).toBeVisible();
    await ribbon.getByTestId("file-branch-manager").click();
    await expect(page.getByTestId("panel-branchManager")).toHaveCount(0);
  });
});
