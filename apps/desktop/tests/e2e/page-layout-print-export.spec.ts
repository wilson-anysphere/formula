import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("Page Layout print/export (web fallback)", () => {
  test("print/export controls are disabled when Tauri APIs are unavailable", async ({ page }) => {
    await gotoDesktop(page);

    const ribbon = page.getByTestId("ribbon-root");
    await expect(ribbon).toBeVisible();

    await ribbon.getByRole("tab", { name: "Page Layout" }).click();

    await expect(ribbon.getByTestId("ribbon-page-setup")).toBeDisabled();
    await expect(ribbon.getByTestId("ribbon-set-print-area")).toBeDisabled();
    await expect(ribbon.getByTestId("ribbon-clear-print-area")).toBeDisabled();
    await expect(ribbon.getByTestId("ribbon-export-pdf")).toBeDisabled();
  });
});

