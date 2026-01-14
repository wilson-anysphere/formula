import { expect, test, type Page } from "@playwright/test";

import { gotoDesktop } from "./helpers";

const TINY_PNG_BASE64 =
  // 1×1 transparent PNG
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PYpgVwAAAABJRU5ErkJggg==";

async function whenIdle(page: Page, timeoutMs: number = 15_000): Promise<void> {
  // Vite may trigger a one-time full reload after dependency optimization.
  // Retry once if the execution context is destroyed mid-wait.
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      await page.waitForFunction(() => typeof (window.__formulaApp as any)?.whenIdle === "function", undefined, {
        timeout: timeoutMs,
      });
      await page.evaluate(async (timeoutMs) => {
        const app = window.__formulaApp as any;
        if (!app || typeof app.whenIdle !== "function") return;
        await Promise.race([app.whenIdle(), new Promise<void>((resolve) => setTimeout(resolve, timeoutMs))]);
      }, timeoutMs);
      return;
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      if (attempt === 0 && message.includes("Execution context was destroyed")) {
        await page.waitForLoadState("domcontentloaded");
        continue;
      }
      throw err;
    }
  }
}

async function getImageDrawingCount(page: Page): Promise<number> {
  return await page.evaluate(() => {
    const app = window.__formulaApp as any;
    if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");

    if (typeof app.getDrawingsDebugState !== "function") {
      throw new Error("Missing window.__formulaApp.getDrawingsDebugState() (required for pictures ribbon e2e)");
    }
    const state = app.getDrawingsDebugState();
    const drawings = Array.isArray(state?.drawings) ? state.drawings : [];
    return drawings.filter((drawing: any) => drawing?.kind === "image").length;
  });
}

test.describe("Insert → Pictures", () => {
  test("Insert → Pictures → This Device opens file picker and inserts image drawings", async ({ page }) => {
    await gotoDesktop(page);
    await whenIdle(page);

    const ribbon = page.getByTestId("ribbon-root");
    await ribbon.getByRole("tab", { name: "Insert" }).click();

    const picturesDropdown = ribbon.getByTestId("ribbon-insert-pictures");
    await expect(picturesDropdown).toBeVisible();
    await picturesDropdown.click();

    const thisDevice = ribbon.getByTestId("ribbon-insert-pictures-this-device");
    await expect(thisDevice).toBeVisible();
    await expect(thisDevice).toBeEnabled();

    const pngBytes = Buffer.from(TINY_PNG_BASE64, "base64");

    const beforeCount = await getImageDrawingCount(page);

    let fileChooser: import("@playwright/test").FileChooser;
    try {
      [fileChooser] = await Promise.all([page.waitForEvent("filechooser", { timeout: 10_000 }), thisDevice.click()]);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      if (message.includes("filechooser") && message.includes("Timeout")) {
        throw new Error(
          `Expected a file chooser to open after clicking Insert → Pictures → This Device… but none was observed.\n\nOriginal error: ${message}`,
        );
      }
      throw err;
    }

    const selectedFiles = fileChooser.isMultiple()
      ? [
          { name: "tiny-1.png", mimeType: "image/png", buffer: pngBytes },
          { name: "tiny-2.png", mimeType: "image/png", buffer: pngBytes },
        ]
      : [{ name: "tiny.png", mimeType: "image/png", buffer: pngBytes }];
    await fileChooser.setFiles(selectedFiles);

    await expect
      .poll(
        async () => {
          await whenIdle(page, 5_000);
          return await getImageDrawingCount(page);
        },
        {
          timeout: 20_000,
          message: `Expected inserting ${selectedFiles.length} image file(s) to create ${selectedFiles.length} image drawing(s).`,
        },
      )
      .toBe(beforeCount + selectedFiles.length);
  });
});
