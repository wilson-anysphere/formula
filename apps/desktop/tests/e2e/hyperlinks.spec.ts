import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.evaluate(() => (window.__formulaApp as any).whenIdle());
}

type InvokeCall = [string, any];

function filterInvokeCalls(calls: InvokeCall[], cmd: string): InvokeCall[] {
  return (Array.isArray(calls) ? calls : []).filter((call) => Array.isArray(call) && call[0] === cmd);
}

async function waitForInvoke(page: import("@playwright/test").Page, cmd: string): Promise<void> {
  await page.waitForFunction((expectedCmd) => {
    const calls = (window as any).__invokeCalls;
    return Array.isArray(calls) && calls.some((call: any[]) => Array.isArray(call) && call[0] === expectedCmd);
  }, cmd);
}

test.describe("external hyperlink opening", () => {
  const GRID_MODES = ["legacy", "shared"] as const;

  for (const mode of GRID_MODES) {
    test(`Ctrl/Cmd+click URL cell opens via the desktop Rust shell command (no webview navigation) (${mode})`, async ({
      page,
    }) => {
      await gotoDesktop(page, `/?grid=${mode}`);
      await waitForIdle(page);

      await page.evaluate(() => {
        (window as any).__invokeCalls = [];
        (window as any).__windowOpenCalls = [];

        (window as any).__TAURI__ = {
          core: {
            invoke: (cmd: string, args: any) => {
              (window as any).__invokeCalls.push([cmd, args]);
              return Promise.resolve();
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
        const app = window.__formulaApp as any;
        const sheetId = app.getCurrentSheetId();
        app.getDocument().setCellValue(sheetId, "A1", "https://example.com");
        await app.whenIdle();
      });

      const a1Rect = await page.evaluate(() => (window.__formulaApp as any).getCellRectA1("A1"));
      expect(a1Rect).not.toBeNull();

      const modifier = process.platform === "darwin" ? "Meta" : "Control";
      await page.click("#grid", {
        position: { x: a1Rect!.x + a1Rect!.width / 2, y: a1Rect!.y + a1Rect!.height / 2 },
        modifiers: [modifier],
      });

      await waitForInvoke(page, "open_external_url");

      const [invokeCalls, windowCalls] = await Promise.all([
        page.evaluate(() => (window as any).__invokeCalls),
        page.evaluate(() => (window as any).__windowOpenCalls),
      ]);

      expect(filterInvokeCalls(invokeCalls, "open_external_url")).toEqual([
        ["open_external_url", { url: "https://example.com" }],
      ]);
      expect(windowCalls).toEqual([]);
    });
  }

  test("clicking an <a href> opens via the desktop Rust shell command (no webview navigation)", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    await page.evaluate(() => {
      (window as any).__invokeCalls = [];
      (window as any).__windowOpenCalls = [];

      (window as any).__TAURI__ = {
        core: {
          invoke: (cmd: string, args: any) => {
            (window as any).__invokeCalls.push([cmd, args]);
            return Promise.resolve();
          },
        },
      };

      const original = window.open;
      (window as any).__originalWindowOpen = original;
      window.open = (...args: any[]) => {
        (window as any).__windowOpenCalls.push(args);
        return null;
      };

      const a = document.createElement("a");
      a.id = "e2e-external-anchor";
      a.href = "https://example.com";
      a.textContent = "external link";
      document.body.appendChild(a);
    });

    const urlBefore = page.url();

    await page.click("#e2e-external-anchor");
    await waitForInvoke(page, "open_external_url");

    const [invokeCalls, windowCalls, urlAfter] = await Promise.all([
      page.evaluate(() => (window as any).__invokeCalls),
      page.evaluate(() => (window as any).__windowOpenCalls),
      page.url(),
    ]);

    const openCalls = filterInvokeCalls(invokeCalls, "open_external_url");
    expect(openCalls).toHaveLength(1);
    expect(openCalls[0][1]?.url).toMatch(/^https:\/\/example\.com\/?$/);
    expect(windowCalls).toEqual([]);
    expect(urlAfter).toBe(urlBefore);
  });

  test("middle-clicking an <a href> opens via the desktop Rust shell command (auxclick)", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    await page.evaluate(() => {
      (window as any).__invokeCalls = [];
      (window as any).__TAURI__ = {
        core: {
          invoke: (cmd: string, args: any) => {
            (window as any).__invokeCalls.push([cmd, args]);
            return Promise.resolve();
          },
        },
      };

      const a = document.createElement("a");
      a.id = "e2e-external-anchor-middle";
      a.href = "https://example.com";
      a.textContent = "external link";
      document.body.appendChild(a);
    });

    await page.click("#e2e-external-anchor-middle", { button: "middle" });
    await waitForInvoke(page, "open_external_url");

    const invokeCalls = await page.evaluate(() => (window as any).__invokeCalls);
    const openCalls = filterInvokeCalls(invokeCalls, "open_external_url");
    expect(openCalls).toHaveLength(1);
    expect(openCalls[0][1]?.url).toMatch(/^https:\/\/example\.com\/?$/);
  });

  test("clicking a mailto: <a href> opens via the desktop Rust shell command", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    await page.evaluate(() => {
      (window as any).__invokeCalls = [];
      (window as any).__TAURI__ = {
        core: {
          invoke: (cmd: string, args: any) => {
            (window as any).__invokeCalls.push([cmd, args]);
            return Promise.resolve();
          },
        },
      };

      const a = document.createElement("a");
      a.id = "e2e-external-anchor-mailto";
      a.href = "mailto:test@example.com";
      a.textContent = "mailto link";
      document.body.appendChild(a);
    });

    const urlBefore = page.url();
    await page.click("#e2e-external-anchor-mailto");
    await waitForInvoke(page, "open_external_url");

    const [invokeCalls, urlAfter] = await Promise.all([page.evaluate(() => (window as any).__invokeCalls), page.url()]);

    expect(filterInvokeCalls(invokeCalls, "open_external_url")).toEqual([
      ["open_external_url", { url: "mailto:test@example.com" }],
    ]);
    expect(urlAfter).toBe(urlBefore);
  });

  test("blocked protocols (javascript:) are never opened", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    await page.evaluate(() => {
      (window as any).__invokeCalls = [];
      (window as any).__TAURI__ = {
        core: {
          invoke: (cmd: string, args: any) => {
            (window as any).__invokeCalls.push([cmd, args]);
            return Promise.resolve();
          },
        },
      };

      const a = document.createElement("a");
      a.id = "e2e-external-anchor-js";
      a.href = "javascript:alert(1)";
      a.textContent = "bad link";
      document.body.appendChild(a);
    });

    const urlBefore = page.url();
    await page.click("#e2e-external-anchor-js");

    const [shellCalls, urlAfter] = await Promise.all([
      page.evaluate(() => (window as any).__invokeCalls),
      page.url(),
    ]);

    expect(filterInvokeCalls(shellCalls, "open_external_url")).toEqual([]);
    expect(urlAfter).toBe(urlBefore);
  });

  test("blocked protocols (data:) are never opened", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    await page.evaluate(() => {
      (window as any).__invokeCalls = [];
      (window as any).__TAURI__ = {
        core: {
          invoke: (cmd: string, args: any) => {
            (window as any).__invokeCalls.push([cmd, args]);
            return Promise.resolve();
          },
        },
      };

      const a = document.createElement("a");
      a.id = "e2e-external-anchor-data";
      a.href = "data:text/plain,hello";
      a.textContent = "bad link";
      document.body.appendChild(a);
    });

    const urlBefore = page.url();
    await page.click("#e2e-external-anchor-data");

    const [shellCalls, urlAfter] = await Promise.all([
      page.evaluate(() => (window as any).__invokeCalls),
      page.url(),
    ]);

    expect(filterInvokeCalls(shellCalls, "open_external_url")).toEqual([]);
    expect(urlAfter).toBe(urlBefore);
  });

  test("blocked protocols (file:) are never opened", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    await page.evaluate(() => {
      (window as any).__invokeCalls = [];
      (window as any).__TAURI__ = {
        core: {
          invoke: (cmd: string, args: any) => {
            (window as any).__invokeCalls.push([cmd, args]);
            return Promise.resolve();
          },
        },
      };

      const a = document.createElement("a");
      a.id = "e2e-external-anchor-file";
      a.href = "file:///etc/passwd";
      a.textContent = "bad link";
      document.body.appendChild(a);
    });

    const urlBefore = page.url();
    await page.click("#e2e-external-anchor-file");
    await page.waitForTimeout(50);

    const [invokeCalls, urlAfter] = await Promise.all([page.evaluate(() => (window as any).__invokeCalls), page.url()]);
    expect(filterInvokeCalls(invokeCalls, "open_external_url")).toEqual([]);
    expect(urlAfter).toBe(urlBefore);
  });

  test("untrusted protocols (ftp:) are blocked in Tauri builds (no prompt)", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    await page.evaluate(() => {
      (window as any).__invokeCalls = [];
      (window as any).__confirmCalls = [];

      window.confirm = (message?: string) => {
        (window as any).__confirmCalls.push(message ?? "");
        return false;
      };

      (window as any).__TAURI__ = {
        core: {
          invoke: (cmd: string, args: any) => {
            (window as any).__invokeCalls.push([cmd, args]);
            return Promise.resolve();
          },
        },
      };

      const a = document.createElement("a");
      a.id = "e2e-external-anchor-ftp-cancel";
      a.href = "ftp://example.com/resource";
      a.textContent = "ftp link";
      document.body.appendChild(a);
    });

    const urlBefore = page.url();
    await page.click("#e2e-external-anchor-ftp-cancel");

    // In desktop/Tauri builds, only http/https/mailto are allowed; untrusted protocols are blocked
    // at the Rust boundary so we should not prompt.
    await page.waitForTimeout(50);

    const [shellCalls, confirmCalls, urlAfter] = await Promise.all([
      page.evaluate(() => (window as any).__invokeCalls),
      page.evaluate(() => (window as any).__confirmCalls),
      page.url(),
    ]);

    expect(confirmCalls).toEqual([]);
    expect(filterInvokeCalls(shellCalls, "open_external_url")).toEqual([]);
    expect(urlAfter).toBe(urlBefore);
  });

  test("untrusted protocols (ftp:) do not call invoke in Tauri builds (no prompt)", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    await page.evaluate(() => {
      (window as any).__invokeCalls = [];
      (window as any).__confirmCalls = [];

      window.confirm = (message?: string) => {
        (window as any).__confirmCalls.push(message ?? "");
        return true;
      };

      (window as any).__TAURI__ = {
        core: {
          invoke: (cmd: string, args: any) => {
            (window as any).__invokeCalls.push([cmd, args]);
            return Promise.resolve();
          },
        },
      };

      const a = document.createElement("a");
      a.id = "e2e-external-anchor-ftp-ok";
      a.href = "ftp://example.com/resource";
      a.textContent = "ftp link";
      document.body.appendChild(a);
    });

    const urlBefore = page.url();
    await page.click("#e2e-external-anchor-ftp-ok");
    await page.waitForTimeout(50);

    const [invokeCalls, confirmCalls, urlAfter] = await Promise.all([
      page.evaluate(() => (window as any).__invokeCalls),
      page.evaluate(() => (window as any).__confirmCalls),
      page.url(),
    ]);

    expect(confirmCalls).toEqual([]);
    expect(filterInvokeCalls(invokeCalls, "open_external_url")).toEqual([]);
    expect(urlAfter).toBe(urlBefore);
  });
});
