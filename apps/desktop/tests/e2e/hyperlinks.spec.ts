import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

test.describe("external hyperlink opening", () => {
  test("Ctrl/Cmd+click URL cell uses the Tauri shell plugin (no webview navigation)", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    await page.evaluate(() => {
      (window as any).__shellOpenCalls = [];
      (window as any).__windowOpenCalls = [];

      (window as any).__TAURI__ = {
        plugin: {
          shell: {
            open: (url: string) => {
              (window as any).__shellOpenCalls.push(url);
              return Promise.resolve();
            },
          },
        },
      };

      const original = window.open;
      (window as any).__originalWindowOpen = original;
      window.open = (...args: any[]) => {
        (window as any).__windowOpenCalls.push(args);
        return null;
      };
    });

    await page.evaluate(async () => {
      const app = (window as any).__formulaApp;
      const sheetId = app.getCurrentSheetId();
      app.getDocument().setCellValue(sheetId, "A1", "https://example.com");
      await app.whenIdle();
    });

    const a1Rect = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("A1"));
    expect(a1Rect).not.toBeNull();

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    await page.click("#grid", {
      position: { x: a1Rect!.x + a1Rect!.width / 2, y: a1Rect!.y + a1Rect!.height / 2 },
      modifiers: [modifier],
    });

    await page.waitForFunction(() => (window as any).__shellOpenCalls?.length === 1);

    const [shellCalls, windowCalls] = await Promise.all([
      page.evaluate(() => (window as any).__shellOpenCalls),
      page.evaluate(() => (window as any).__windowOpenCalls),
    ]);

    expect(shellCalls).toEqual(["https://example.com"]);
    expect(windowCalls).toEqual([]);
  });
});

