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

  test("System theme option follows OS preference", async ({ page }) => {
    await gotoDesktop(page);

    // Start with the UX default (Light), even though the OS is dark.
    await expect(page.locator("html")).toHaveAttribute("data-theme", "light");

    await page.getByRole("tab", { name: "View", exact: true }).click();

    const themeDropdown = page.getByTestId("ribbon-root").getByTestId("theme-selector");
    await expect(themeDropdown).toBeVisible();
    await themeDropdown.click();

    await page.locator('[role="menuitem"][data-command-id="view.appearance.theme.system"]').click();

    // With OS dark, System should resolve to Dark.
    await expect(page.locator("html")).toHaveAttribute("data-theme", "dark");
  });

  test("System theme updates when OS preference changes", async ({ page }) => {
    await gotoDesktop(page);

    await page.getByRole("tab", { name: "View", exact: true }).click();
    const themeDropdown = page.getByTestId("ribbon-root").getByTestId("theme-selector");
    await themeDropdown.click();
    await page.locator('[role="menuitem"][data-command-id="view.appearance.theme.system"]').click();

    // Starts dark because the OS is emulated as dark.
    await expect(page.locator("html")).toHaveAttribute("data-theme", "dark");

    // Switching the emulated OS preference should update the resolved theme.
    await page.emulateMedia({ colorScheme: "light" });
    await expect(page.locator("html")).toHaveAttribute("data-theme", "light");
  });
});
