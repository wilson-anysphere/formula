import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("python panel", () => {
  test("runs a script that prints output and updates the workbook", async ({ page }) => {
    test.setTimeout(120_000);

    const script = `import formula

sheet = formula.active_sheet
sheet["A1"] = 777
sheet["A2"] = "=A1*2"
print("Hello from Python")
`;

    // Vite may trigger a one-time full reload after dependency optimization (e.g. when Pyodide
    // dependencies are first loaded). If that happens, retry the whole interaction once after
    // the navigation completes.
    for (let attempt = 0; attempt < 2; attempt += 1) {
      await gotoDesktop(page);
      try {
        const isolation = await page.evaluate(() => ({
          crossOriginIsolated: globalThis.crossOriginIsolated,
          sharedArrayBuffer: typeof (globalThis as any).SharedArrayBuffer !== "undefined",
        }));
        expect(isolation.crossOriginIsolated).toBe(true);
        expect(isolation.sharedArrayBuffer).toBe(true);

        await page.getByTestId("ribbon-root").getByTestId("open-python-panel").click();
        const panel = page.getByTestId("dock-bottom").getByTestId("panel-python");
        await expect(panel).toBeVisible();

        const editor = panel.getByTestId("python-panel-code");
        await expect(editor).toBeVisible();
        await editor.fill(script);

        await panel.getByTestId("python-panel-run").click();

        await expect(panel.getByTestId("python-panel-output")).toContainText("Hello from Python", { timeout: 120_000 });

        await expect
          .poll(async () => page.evaluate(() => (window.__formulaApp as any).getCellValueA1("A1")))
          .toBe("777");
        await expect
          .poll(async () => page.evaluate(() => (window.__formulaApp as any).getCellValueA1("A2")))
          .toBe("1554");
        break;
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        if (
          attempt === 0 &&
          (message.includes("Execution context was destroyed") || message.includes("element(s) not found"))
        ) {
          await page.waitForLoadState("domcontentloaded");
          continue;
        }
        throw err;
      }
    }
  });
});
