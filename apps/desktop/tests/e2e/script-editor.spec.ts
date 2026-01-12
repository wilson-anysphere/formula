import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("script editor panel", () => {
  test("runs a script that writes to the workbook", async ({ page }) => {
    test.setTimeout(60_000);
    const script = `// Write a value to C1
await ctx.activeSheet.getRange("C1").setValue(99);
`;

    // Vite may trigger a one-time full reload after dependency optimization (e.g. when the
    // scripting worker is first instantiated). If that happens, retry the whole interaction
    // once after the navigation completes.
    for (let attempt = 0; attempt < 2; attempt += 1) {
      await gotoDesktop(page);

      await page.getByTestId("ribbon-root").getByTestId("open-script-editor-panel").click();
      const panel = page.getByTestId("dock-bottom").getByTestId("panel-scriptEditor");
      await expect(panel).toBeVisible();

      const editor = panel.getByTestId("script-editor-code");
      await expect(editor).toBeVisible();
      await editor.fill(script);

      const runButton = panel.getByTestId("script-editor-run");
      await runButton.click();
      await expect(runButton).toBeDisabled();

      try {
        await expect(runButton).toBeEnabled({ timeout: 30_000 });
        await expect
          .poll(
            async () => page.evaluate(() => (window.__formulaApp as any)?.getCellValueA1?.("C1") ?? ""),
            { timeout: 20_000 },
          )
          .toBe("99");
        break;
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        if (attempt === 0 && message.includes("Execution context was destroyed")) {
          await page.waitForLoadState("domcontentloaded");
          continue;
        }
        throw err;
      }
    }
  });
});
