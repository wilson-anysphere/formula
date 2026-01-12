import { expect, test } from "@playwright/test";

test("in-cell editor F4 toggles absolute/relative A1 references", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });

  const selectionCanvas = page.getByTestId("canvas-grid-selection");

  // Click B1 (row 1, col 2) so the formula can reference A1 without being self-referential.
  await selectionCanvas.click({ position: { x: 250, y: 31 } });
  await expect(page.getByTestId("active-address")).toHaveText("B1");

  // Start editing by typing "=" (type-to-edit).
  await page.keyboard.type("=");
  const editor = page.getByTestId("cell-editor");
  await expect(editor).toBeVisible();
  await editor.fill("=A1");

  // Place caret inside A1.
  await editor.press("ArrowLeft");

  await editor.press("F4");
  await expect(editor).toHaveValue("=$A$1");
  await expect(editor).toHaveJSProperty("selectionStart", 1);
  await expect(editor).toHaveJSProperty("selectionEnd", 5);

  await editor.press("F4");
  await expect(editor).toHaveValue("=A$1");
  await expect(editor).toHaveJSProperty("selectionStart", 1);
  await expect(editor).toHaveJSProperty("selectionEnd", 4);

  await editor.press("F4");
  await expect(editor).toHaveValue("=$A1");
  await expect(editor).toHaveJSProperty("selectionStart", 1);
  await expect(editor).toHaveJSProperty("selectionEnd", 4);

  await editor.press("F4");
  await expect(editor).toHaveValue("=A1");
  await expect(editor).toHaveJSProperty("selectionStart", 1);
  await expect(editor).toHaveJSProperty("selectionEnd", 3);
});

test("in-cell editor F4 expands full-token selections to cover the toggled token", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });

  const selectionCanvas = page.getByTestId("canvas-grid-selection");

  // Click B1 (row 1, col 2) so the formula can reference A1 without being self-referential.
  await selectionCanvas.click({ position: { x: 250, y: 31 } });
  await expect(page.getByTestId("active-address")).toHaveText("B1");

  // Start editing by typing "=" (type-to-edit).
  await page.keyboard.type("=");
  const editor = page.getByTestId("cell-editor");
  await expect(editor).toBeVisible();
  await editor.fill("=A1");

  // Select the full reference token.
  await editor.evaluate((el) => {
    const input = el as HTMLInputElement;
    input.focus();
    input.setSelectionRange(1, 3);
  });

  await editor.press("F4");
  await expect(editor).toHaveValue("=$A$1");
  await expect(editor).toHaveJSProperty("selectionStart", 1);
  await expect(editor).toHaveJSProperty("selectionEnd", 5);
});

test("in-cell editor F4 is a no-op when the selection is not contained within a reference token", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });

  const selectionCanvas = page.getByTestId("canvas-grid-selection");

  // Click B1 (row 1, col 2) so the formula can reference A1 without being self-referential.
  await selectionCanvas.click({ position: { x: 250, y: 31 } });
  await expect(page.getByTestId("active-address")).toHaveText("B1");

  await page.keyboard.type("=");
  const editor = page.getByTestId("cell-editor");
  await expect(editor).toBeVisible();
  await editor.fill("=A1+B1");

  // Select across A1 and "+" (not fully contained within a reference).
  await editor.evaluate((el) => {
    const input = el as HTMLInputElement;
    input.focus();
    input.setSelectionRange(1, 4);
  });

  await editor.press("F4");
  await expect(editor).toHaveValue("=A1+B1");
});

test("in-cell editor F4 toggles sheet-qualified range references", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });

  const selectionCanvas = page.getByTestId("canvas-grid-selection");

  // Click B1 (row 1, col 2) so the formula can reference A1 without being self-referential.
  await selectionCanvas.click({ position: { x: 250, y: 31 } });
  await expect(page.getByTestId("active-address")).toHaveText("B1");

  await page.keyboard.type("=");
  const editor = page.getByTestId("cell-editor");
  await expect(editor).toBeVisible();
  await editor.fill("='My Sheet'!A1:B2");

  // Place caret inside the reference token.
  await editor.press("ArrowLeft");

  await editor.press("F4");
  await expect(editor).toHaveValue("='My Sheet'!$A$1:$B$2");
  const value = await editor.inputValue();
  await expect(editor).toHaveJSProperty("selectionStart", 1);
  await expect(editor).toHaveJSProperty("selectionEnd", value.length);
});
