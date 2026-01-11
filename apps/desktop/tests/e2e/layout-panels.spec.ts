import { expect, test } from "@playwright/test";

test.describe("dockable panels layout persistence", () => {
  test("open AI panel, dock left, reload restores layout", async ({ page }) => {
    await page.goto("/");
    await page.evaluate(() => localStorage.clear());
    await page.reload();

    // Open AI panel (defaults to right dock via panel registry).
    await page.getByTestId("open-ai-panel").click();
    await expect(page.getByTestId("dock-right").getByTestId("panel-aiChat")).toBeVisible();

    // Dock to left.
    await page.getByTestId("dock-ai-panel-left").click();
    await expect(page.getByTestId("dock-left").getByTestId("panel-aiChat")).toBeVisible();
    await expect(page.getByTestId("dock-right").getByTestId("panel-aiChat")).toHaveCount(0);

    // Reload: layout should restore from localStorage.
    await page.reload();

    await expect(page.getByTestId("dock-left").getByTestId("panel-aiChat")).toBeVisible();
    await expect(page.getByTestId("dock-right").getByTestId("panel-aiChat")).toHaveCount(0);
  });
});
