import { expect, test } from "@playwright/test";

test("formula bar F4 cycles absolute/relative A1 references", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });

  const input = page.getByTestId("formula-input");
  await input.click();
  await input.fill("=A1");

  // Place caret inside A1.
  await input.press("ArrowLeft");

  await input.press("F4");
  await expect(input).toHaveValue("=$A$1");
  await expect(input).toHaveJSProperty("selectionStart", 1);
  await expect(input).toHaveJSProperty("selectionEnd", 5);

  await input.press("F4");
  await expect(input).toHaveValue("=A$1");
  await expect(input).toHaveJSProperty("selectionStart", 1);
  await expect(input).toHaveJSProperty("selectionEnd", 4);

  await input.press("F4");
  await expect(input).toHaveValue("=$A1");
  await expect(input).toHaveJSProperty("selectionStart", 1);
  await expect(input).toHaveJSProperty("selectionEnd", 4);

  await input.press("F4");
  await expect(input).toHaveValue("=A1");
  await expect(input).toHaveJSProperty("selectionStart", 1);
  await expect(input).toHaveJSProperty("selectionEnd", 3);
});
