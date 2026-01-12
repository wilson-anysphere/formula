import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.waitForFunction(() => Boolean((window as any).__formulaApp?.whenIdle), null, { timeout: 10_000 });
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

async function stubDateNow(page: import("@playwright/test").Page): Promise<void> {
  await page.evaluate(() => {
    const fixed = new Date(2020, 0, 2, 3, 4, 5).valueOf();
    const OriginalDate = Date;

    class MockDate extends OriginalDate {
      constructor(...args: any[]) {
        if (args.length === 0) {
          super(fixed);
          return;
        }
        if (args.length === 1) {
          super(args[0] as any);
          return;
        }
        // Avoid `super(...args)` to keep TypeScript happy with the overloaded `Date` constructor.
        // (TS requires spread arguments to be tuple-typed when targeting non-rest call signatures.)
        super(
          args[0] as any,
          args[1] as any,
          args[2] as any,
          args[3] as any,
          args[4] as any,
          args[5] as any,
          args[6] as any
        );
      }

      static now() {
        return fixed;
      }
    }

    // eslint-disable-next-line no-global-assign
    (globalThis as any).Date = MockDate;
  });
}

test.describe("insert date/time shortcuts (Ctrl/Cmd+;)", () => {
  const gridModes = ["legacy", "shared"] as const;

  for (const mode of gridModes) {
    test(`Ctrl+; inserts the current date (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);
      await waitForIdle(page);

      await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        app.selectRange({ range: { startRow: 0, endRow: 1, startCol: 0, endCol: 1 } });
        app.focus();
      });
      await stubDateNow(page);

      await page.evaluate(() => {
        const isMac = /mac/i.test(navigator.platform);
        const root = document.getElementById("grid");
        root?.dispatchEvent(
          new KeyboardEvent("keydown", {
            bubbles: true,
            cancelable: true,
            code: "Semicolon",
            key: ";",
            ctrlKey: !isMac,
            metaKey: isMac,
          }),
        );
      });
      await waitForIdle(page);

      const [a1, b1, a2, b2] = await Promise.all([
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1")),
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("B1")),
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A2")),
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("B2")),
      ]);
      expect(a1).toBe("2020-01-02");
      expect(b1).toBe("2020-01-02");
      expect(a2).toBe("2020-01-02");
      expect(b2).toBe("2020-01-02");
    });

    test(`Ctrl+Shift+; inserts the current time (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);
      await waitForIdle(page);

      await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        app.selectRange({ range: { startRow: 0, endRow: 1, startCol: 0, endCol: 1 } });
        app.focus();
      });
      await stubDateNow(page);

      await page.evaluate(() => {
        const isMac = /mac/i.test(navigator.platform);
        const root = document.getElementById("grid");
        root?.dispatchEvent(
          new KeyboardEvent("keydown", {
            bubbles: true,
            cancelable: true,
            code: "Semicolon",
            key: ":",
            shiftKey: true,
            ctrlKey: !isMac,
            metaKey: isMac,
          }),
        );
      });
      await waitForIdle(page);

      const [a1, b1, a2, b2] = await Promise.all([
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1")),
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("B1")),
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A2")),
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("B2")),
      ]);
      expect(a1).toBe("03:04:05");
      expect(b1).toBe("03:04:05");
      expect(a2).toBe("03:04:05");
      expect(b2).toBe("03:04:05");
    });
  }
});
