import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("AI inline edit (ribbon)", () => {
  test("opens from the ribbon", async ({ page }) => {
    await gotoDesktop(page);

    const ribbon = page.getByTestId("ribbon-root");
    await expect(ribbon).toBeVisible();

    await ribbon.getByTestId("open-inline-ai-edit").click();
    await expect(page.getByTestId("inline-edit-overlay")).toBeVisible();
  });
});

