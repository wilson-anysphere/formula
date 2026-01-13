import { expect, test } from "@playwright/test";

import { gotoDesktop, waitForDesktopReady } from "./helpers";

test.describe("theme selector", () => {
  // Validate the UX spec: new users should start in Light theme even when the OS prefers dark.
  test.use({ colorScheme: "dark" });

  test("switches to Dark theme via ribbon and persists across reload", async ({ page }) => {
    await gotoDesktop(page);

    // Default should be light (not System).
    await expect(page.locator("html")).toHaveAttribute("data-theme", "light");

    await page.getByRole("tab", { name: "View", exact: true }).click();

    const themeDropdown = page.getByTestId("ribbon-root").getByTestId("theme-selector");
    await expect(themeDropdown).toBeVisible();
    await themeDropdown.click();

    await page.locator('[role="menuitem"][data-command-id="view.appearance.theme.dark"]').click();

    await expect(page.locator("html")).toHaveAttribute("data-theme", "dark");

    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);

    await expect(page.locator("html")).toHaveAttribute("data-theme", "dark");
  });
});
