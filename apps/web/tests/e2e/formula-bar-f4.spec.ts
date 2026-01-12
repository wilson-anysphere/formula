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

test("formula bar F4 expands full-token selections to cover the toggled token", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });

  const input = page.getByTestId("formula-input");
  await input.click();
  await input.fill("=A1+B1");

  await input.evaluate((el) => {
    const inputEl = el as HTMLInputElement;
    inputEl.focus();
    inputEl.setSelectionRange(1, 3);
  });

  await input.press("F4");
  await expect(input).toHaveValue("=$A$1+B1");
  await expect(input).toHaveJSProperty("selectionStart", 1);
  await expect(input).toHaveJSProperty("selectionEnd", 5);
});

test("formula bar F4 is a no-op when the selection is not contained within a reference token", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });

  const input = page.getByTestId("formula-input");
  await input.click();
  await input.fill("=A1+B1");

  // Select across A1 and the "+" operator (not fully contained within a reference).
  await input.evaluate((el) => {
    const inputEl = el as HTMLInputElement;
    inputEl.focus();
    inputEl.setSelectionRange(1, 4);
  });

  await input.press("F4");
  await expect(input).toHaveValue("=A1+B1");
});

test("formula bar F4 preserves sheet qualifiers and toggles range endpoints", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });

  const input = page.getByTestId("formula-input");
  await input.click();
  await input.fill("='My Sheet'!A1:B2");

  // Place caret inside the reference token.
  await input.press("ArrowLeft");

  await input.press("F4");
  await expect(input).toHaveValue("='My Sheet'!$A$1:$B$2");
  const value = await input.inputValue();
  await expect(input).toHaveJSProperty("selectionStart", 1);
  await expect(input).toHaveJSProperty("selectionEnd", value.length);
});
