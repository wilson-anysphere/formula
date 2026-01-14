import { expect, test, type Page } from "@playwright/test";

import { gotoDesktop } from "./helpers";

const TINY_PNG_BASE64 =
  // 1×1 transparent PNG
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PYpgVwAAAABJRU5ErkJggg==";

async function evaluateWithRetry<T>(page: Page, fn: () => T | Promise<T>): Promise<T> {
  // Vite may trigger a one-time full reload after dependency optimization.
  // Retry once if the execution context is destroyed mid-evaluate.
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      return await page.evaluate(fn);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      if (
        attempt === 0 &&
        (message.includes("Execution context was destroyed") ||
          message.includes("frame was detached") ||
          message.includes("net::ERR_ABORTED"))
      ) {
        await page.waitForLoadState("domcontentloaded");
        continue;
      }
      throw err;
    }
  }
  // Unreachable, but keeps TypeScript happy.
  throw new Error("evaluateWithRetry exhausted retries");
}

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
      if (
        attempt === 0 &&
        (message.includes("Execution context was destroyed") ||
          message.includes("frame was detached") ||
          message.includes("net::ERR_ABORTED"))
      ) {
        await page.waitForLoadState("domcontentloaded");
        continue;
      }
      throw err;
    }
  }
}

async function getImageDrawingCount(page: Page): Promise<number> {
  return await evaluateWithRetry(page, () => {
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

async function getDrawingsDebugSummary(page: Page): Promise<{ sheetId: string; selectedId: number | null; imageCount: number }> {
  return await evaluateWithRetry(page, () => {
    const app = window.__formulaApp as any;
    if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
    if (typeof app.getDrawingsDebugState !== "function") {
      throw new Error("Missing window.__formulaApp.getDrawingsDebugState()");
    }
    const state = app.getDrawingsDebugState();
    const drawings = Array.isArray(state?.drawings) ? state.drawings : [];
    const imageCount = drawings.filter((drawing: any) => drawing?.kind === "image").length;
    return { sheetId: String(state?.sheetId ?? ""), selectedId: (state?.selectedId ?? null) as number | null, imageCount };
  });
}

test.describe("Insert → Pictures", () => {
  const GRID_MODES = ["legacy", "shared"] as const;

  for (const mode of GRID_MODES) {
    test(`Insert → Pictures → This Device opens file picker and inserts image drawings (${mode})`, async ({ page }) => {
      const url = mode === "legacy" ? "/?grid=legacy" : "/?grid=shared";
      // Pictures insertion triggers async image decode/IndexedDB writes. Avoid relying on an
      // unbounded `whenIdle()` at navigation time: give the app a moment to settle, but don't
      // hang forever if background work keeps it busy.
      await gotoDesktop(page, url, { idleTimeoutMs: 10_000 });
      await whenIdle(page);

      // Ensure insertion starts from a deterministic location so the inserted pictures land in the viewport.
      await evaluateWithRetry(page, () => {
        const app = window.__formulaApp as any;
        app.activateCell({ row: 0, col: 0 });
        app.focus();
      });

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

      const multiFiles = [
        { name: "tiny-1.png", mimeType: "image/png", buffer: pngBytes },
        { name: "tiny-2.png", mimeType: "image/png", buffer: pngBytes },
      ];
      const singleFile = [{ name: "tiny.png", mimeType: "image/png", buffer: pngBytes }];

      // Preferred: filechooser event.
      // Fallback: if the implementation uses a persistent/hidden <input type=file>, set files directly.
      let fileChooserError: unknown = null;
      // The file chooser should appear immediately after the click, but be a bit
      // forgiving in CI/headless where scheduling can be noisier.
      const fileChooserPromise = page.waitForEvent("filechooser", { timeout: 5_000 }).catch((err) => {
        fileChooserError = err;
        return null;
      });

      await thisDevice.click();
      const fileChooser = await fileChooserPromise;

      let selectedFiles = singleFile;
      if (fileChooser) {
        selectedFiles = fileChooser.isMultiple() ? multiFiles : singleFile;
        await fileChooser.setFiles(selectedFiles);
      } else {
        const input = page.locator('input[type="file"]:not([disabled])[accept*="image"]').last();
        try {
          await expect(input).toBeAttached({ timeout: 5_000 });
        } catch {
          const rawError = fileChooserError instanceof Error ? fileChooserError.message : String(fileChooserError);
          throw new Error(
            `Expected a file picker to open after clicking Insert → Pictures → This Device… but none was observed (no filechooser event and no image <input type=file> found).\n\nfilechooser error: ${rawError}`,
          );
        }
        const multiple = await input.evaluate((el) => (el as HTMLInputElement).multiple).catch(() => false);
        selectedFiles = multiple ? multiFiles : singleFile;
        await input.setInputFiles(selectedFiles);
      }

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

      // Assert at least one inserted image has a resolved rect (i.e. would be visible as a drawing overlay).
      await expect
        .poll(
          async () => {
            await whenIdle(page, 5_000);
            return await evaluateWithRetry(page, () => {
              const state = (window.__formulaApp as any).getDrawingsDebugState();
              const drawings = Array.isArray(state?.drawings) ? state.drawings : [];
              const visibleImages = drawings.filter((d: any) => {
                if (d?.kind !== "image") return false;
                const rect = d?.rectPx;
                return rect && rect.width > 0 && rect.height > 0;
              });
              return visibleImages.length;
            });
          },
          { timeout: 20_000, message: "Expected at least one inserted picture to have a non-empty drawing rect." },
        )
        .toBeGreaterThanOrEqual(1);

      // Sheet switching should be sheet-scoped: pictures belong to their owning sheet and do not
      // render/appear in `getDrawingsDebugState` when switching to a different sheet.
      await evaluateWithRetry(page, () => {
        const app = window.__formulaApp as any;
        // Lazily create Sheet2 so the sheet tab appears.
        app.getDocument().setCellValue("Sheet2", "A1", "Hello from Sheet2");
      });
      await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

      await page.getByTestId("sheet-tab-Sheet2").click();
      await whenIdle(page);

      await expect
        .poll(
          async () => {
            await whenIdle(page, 5_000);
            return await getDrawingsDebugSummary(page);
          },
          { timeout: 20_000, message: "Expected switching sheets to hide Sheet1 pictures and clear selection." },
        )
        .toEqual({ sheetId: "Sheet2", selectedId: null, imageCount: 0 });

      await page.getByTestId("sheet-tab-Sheet1").click();
      await whenIdle(page);

      await expect
        .poll(
          async () => {
            await whenIdle(page, 5_000);
            return await getDrawingsDebugSummary(page);
          },
          { timeout: 20_000, message: "Expected switching back to restore Sheet1 pictures without reselecting them." },
        )
        .toEqual({ sheetId: "Sheet1", selectedId: null, imageCount: beforeCount + selectedFiles.length });
    });
  }
});
