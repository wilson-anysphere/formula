import { expect, test, type Page } from "@playwright/test";

import { gotoDesktop, openExtensionsPanel, waitForDesktopReady } from "./helpers";

async function grantSampleHelloPermissions(page: Page, extra: Record<string, boolean> = {}): Promise<void> {
  await page.evaluate((extra) => {
    // E2E stability: auto-accept any permission prompts so modal dialogs don't
    // intercept pointer events during UI interactions.
    (window as any).__formulaPermissionPrompt = () => true;

    const extensionId = "formula.sample-hello";
    const key = "formula.extensionHost.permissions";
    const existing = (() => {
      try {
        const raw = localStorage.getItem(key);
        return raw ? JSON.parse(raw) : {};
      } catch {
        return {};
      }
    })();

    existing[extensionId] = {
      ...(existing[extensionId] ?? {}),
      "ui.commands": true,
      "ui.panels": true,
      "cells.read": true,
      "cells.write": true,
      "workbook.manage": true,
      network: true,
      clipboard: true,
      storage: true,
      // Spread after the defaults so callers can grant additional permissions (e.g. clipboard)
      // without needing to repeat the baseline set used across most tests.
      ...(extra ?? {}),
    };

    localStorage.setItem(key, JSON.stringify(existing));
  }, extra);
}

test.describe("Extensions UI integration", () => {
  // The desktop shell has a large ribbon; the default Playwright viewport height can
  // leave too little space for the grid, making hit-testing unreliable. Use a
  // taller viewport so context menu/selection interactions have room.
  test.use({ viewport: { width: 1280, height: 900 } });

  test("shows extension commands in the command palette after lazy-loading extensions", async ({ page }) => {
    await gotoDesktop(page);

    // Open the command palette first (without opening the Extensions panel) and
    // ensure that extension-contributed commands appear once the extension host loads.
    const primary = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${primary}+Shift+P`);
    await expect(page.getByTestId("command-palette")).toBeVisible();

    await page.getByTestId("command-palette-input").fill("Sum Selection");

    // Command palette groups commands by category, rendering the category as a group header
    // and the command title as the selectable row.
    const list = page.getByTestId("command-palette-list");
    await expect(list).toContainText("Sample Hello", { timeout: 10_000 });
    await expect(list).toContainText("Sum Selection", { timeout: 10_000 });
  });

  test("runs sampleHello.openPanel and renders the panel webview", async ({ page }) => {
    // Simulate a runtime that injects Tauri globals into iframe documents.
    // The injected webview hardening script should remove them before any extension code can access them.
    await page.addInitScript(() => {
      // Only inject into iframes so we don't accidentally toggle the desktop app into "Tauri mode"
      // during web-based e2e runs.
      try {
        if (window.top === window) return;
      } catch {
        // Assume we're in a frame if we can't compare `top` safely.
      }

      try {
        // Extension webviews are loaded from a blob URL; avoid affecting other nested frames.
        if (window.location?.protocol !== "blob:") return;
      } catch {
        // Ignore.
      }

      const injectTauriGlobals = () => {
        try {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const w = window as any;
          const define = (key: string) => {
            try {
              // Simulate a native runtime defining globals as non-configurable properties.
              Object.defineProperty(w, key, {
                value: {},
                writable: true,
                configurable: false,
                enumerable: true,
              });
            } catch {
              try {
                w[key] = {};
              } catch {
                // Ignore.
              }
            }
          };

          define("__TAURI__");
          define("__TAURI_IPC__");
          define("__TAURI_INTERNALS__");
          define("__TAURI_METADATA__");
          define("__TAURI_INVOKE__");
        } catch {
          // Ignore.
        }
      };

      // Inject globals later in the document lifecycle to ensure our hardening script scrubs both
      // early and late injections.
      try {
        document.addEventListener("DOMContentLoaded", injectTauriGlobals, { once: true });
        window.addEventListener("load", injectTauriGlobals, { once: true });
        setTimeout(() => {
          injectTauriGlobals();
          try {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            (window as any).__formulaTestTauriInjectedLate = true;
          } catch {
            // Ignore.
          }
        }, 500);
      } catch {
        // Ignore.
      }
    });

    await gotoDesktop(page);
    await grantSampleHelloPermissions(page);

    await openExtensionsPanel(page);
    const openPanelBtn = page.getByTestId("run-command-sampleHello.openPanel");
    await expect(openPanelBtn).toBeVisible({ timeout: 30_000 });
    // Avoid hit-target flakiness from fixed overlays by dispatching a click directly.
    await openPanelBtn.dispatchEvent("click");

    await expect(page.getByTestId("panel-sampleHello.panel")).toBeAttached();
    const iframeLocator = page.locator('iframe[data-testid="extension-webview-sampleHello.panel"]');
    await expect(iframeLocator, "webview iframe should be sandboxed").toHaveAttribute("sandbox", /allow-scripts/);
    const sandboxAttr = await iframeLocator.getAttribute("sandbox");
    expect(sandboxAttr ?? "").not.toContain("allow-same-origin");
    await expect(iframeLocator, "webview should not send referrers").toHaveAttribute("referrerpolicy", "no-referrer");
    await expect(iframeLocator, "webview should load from a blob: URL").toHaveAttribute("src", /^blob:/);
    await expect(iframeLocator, "webview should explicitly disable clipboard features").toHaveAttribute(
      "allow",
      /clipboard-read 'none'; clipboard-write 'none'/,
    );
    await expect(iframeLocator, "webview should explicitly disable camera").toHaveAttribute("allow", /camera 'none'/);
    await expect(iframeLocator, "webview should explicitly disable microphone").toHaveAttribute("allow", /microphone 'none'/);
    await expect(iframeLocator, "webview should explicitly disable geolocation").toHaveAttribute("allow", /geolocation 'none'/);

    const frame = page.frameLocator('iframe[data-testid="extension-webview-sampleHello.panel"]');
    await expect(frame.locator("h1")).toHaveText("Sample Hello Panel");
    await expect(
      frame.locator('meta[http-equiv="Content-Security-Policy"]'),
      "webview should inject a restrictive CSP meta tag",
    ).toHaveCount(1);
    await expect(
      frame.locator('script[src^="data:text/javascript"]').first(),
      "webview should inject a hardening script via a data: URL (not inline)",
    ).toBeAttached();
    const cspContent = await frame.locator('meta[http-equiv="Content-Security-Policy"]').getAttribute("content");
    expect(cspContent).toContain("default-src 'none'");
    expect(cspContent).toContain("connect-src 'none'");
    expect(cspContent).toContain("frame-src 'none'");
    expect(cspContent).toContain("script-src blob: data:");

    const iframeHandle = await page
      .locator('iframe[data-testid="extension-webview-sampleHello.panel"]')
      .elementHandle();
    expect(iframeHandle, "expected webview iframe to exist").not.toBeNull();

    const webviewFrame = await iframeHandle!.contentFrame();
    expect(webviewFrame, "expected webview iframe to have a content frame").not.toBeNull();

    const hasHardeningDataScript = await webviewFrame!.evaluate(() => {
      try {
        const scripts = Array.from(document.querySelectorAll('script[src^="data:text/javascript"]'));
        return scripts.some((script) => {
          const src = script.getAttribute("src") ?? "";
          const comma = src.indexOf(",");
          if (comma === -1) return false;
          try {
            return decodeURIComponent(src.slice(comma + 1)).includes("__formulaWebviewSandbox");
          } catch {
            return false;
          }
        });
      } catch {
        return false;
      }
    });
    expect(hasHardeningDataScript, "webview should include a hardening data: script").toBe(true);

    // Ensure the document lifecycle has advanced far enough that our simulated Tauri injection
    // (DOMContentLoaded/load) and the hardening script's re-scrub passes have both executed.
    await webviewFrame!.waitForFunction(() => document.readyState === "complete");
    await webviewFrame!.waitForFunction(() => (window as any).__formulaTestTauriInjectedLate === true, undefined, {
      timeout: 5_000,
    });
    await webviewFrame!.waitForFunction(
      () => (window as any).__formulaWebviewSandbox?.tauriGlobalsPresent === true,
      undefined,
      { timeout: 5_000 },
    );
    await webviewFrame!.waitForFunction(
      () => (window as any).__formulaWebviewSandbox?.tauriGlobalsScrubbed === true,
      undefined,
      { timeout: 5_000 },
    );
    // `Location.origin` reflects the URL origin (blob:http://...), which can still serialize to the
    // parent origin even when the iframe is sandboxed with an opaque origin. Use `window.origin`
    // (global environment origin) instead.
    const iframeOrigin = await webviewFrame!.evaluate(() => window.origin);
    expect(iframeOrigin, "webview should have an opaque origin (no allow-same-origin sandbox flag)").toBe("null");
    const parentAccessBlocked = await webviewFrame!.evaluate(() => {
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const app = (window.top as any).__formulaApp;
        // Access should either throw or return `undefined` for cross-origin blocked properties.
        return typeof app === "undefined";
      } catch {
        return true;
      }
    });
    expect(parentAccessBlocked, "webview should not be able to access parent window properties").toBe(true);
    const clipboardPolicy = await webviewFrame!.evaluate(() => {
      // `featurePolicy` is deprecated but still widely supported; `permissionsPolicy` is the newer name.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const policy: any = (document as any).permissionsPolicy ?? (document as any).featurePolicy;
      if (!policy || typeof policy.allowsFeature !== "function") return null;
      try {
        return {
          read: policy.allowsFeature("clipboard-read"),
          write: policy.allowsFeature("clipboard-write"),
        };
      } catch {
        return null;
      }
    });
    if (clipboardPolicy) {
      expect(clipboardPolicy.read, "webview should not allow clipboard-read").toBe(false);
      expect(clipboardPolicy.write, "webview should not allow clipboard-write").toBe(false);
    }

    const sandboxInfo = await webviewFrame!.evaluate(() => (window as any).__formulaWebviewSandbox);
    expect(sandboxInfo, "webview should inject a sandbox hardening script").toBeTruthy();
    expect(typeof sandboxInfo.tauriGlobalsPresent).toBe("boolean");
    expect(sandboxInfo.tauriGlobalsPresent, "webview should detect injected Tauri globals").toBe(true);
    expect(sandboxInfo.tauriGlobalsScrubbed, "webview should scrub injected Tauri globals").toBe(true);

    const sandboxDescriptor = await webviewFrame!.evaluate(() => {
      const desc = Object.getOwnPropertyDescriptor(window, "__formulaWebviewSandbox") as any;
      if (!desc) return null;
      return { writable: desc.writable, configurable: desc.configurable };
    });
    expect(sandboxDescriptor, "webview sandbox marker should be a defined window property").toBeTruthy();
    expect(sandboxDescriptor?.writable, "webview sandbox marker should not be writable").toBe(false);
    expect(sandboxDescriptor?.configurable, "webview sandbox marker should not be configurable").toBe(false);

    const tauriTypes = await webviewFrame!.evaluate(() => ({
      tauri: typeof (window as any).__TAURI__,
      tauriIpc: typeof (window as any).__TAURI_IPC__,
      tauriInternals: typeof (window as any).__TAURI_INTERNALS__,
      tauriMetadata: typeof (window as any).__TAURI_METADATA__,
      tauriInvoke: typeof (window as any).__TAURI_INVOKE__,
    }));
    expect(tauriTypes.tauri, "webview should not expose __TAURI__").toBe("undefined");
    expect(tauriTypes.tauriIpc, "webview should not expose __TAURI_IPC__").toBe("undefined");
    expect(tauriTypes.tauriInternals, "webview should not expose __TAURI_INTERNALS__").toBe("undefined");
    expect(tauriTypes.tauriMetadata, "webview should not expose __TAURI_METADATA__").toBe("undefined");
    expect(tauriTypes.tauriInvoke, "webview should not expose __TAURI_INVOKE__").toBe("undefined");

    const tauriDescriptors = await webviewFrame!.evaluate(() => {
      const read = (key: string) => {
        const desc = Object.getOwnPropertyDescriptor(window, key) as any;
        if (!desc) return null;
        return {
          writable: desc.writable,
          configurable: desc.configurable,
          valueType: typeof desc.value,
        };
      };

      return {
        tauri: read("__TAURI__"),
        tauriIpc: read("__TAURI_IPC__"),
        tauriInternals: read("__TAURI_INTERNALS__"),
        tauriMetadata: read("__TAURI_METADATA__"),
        tauriInvoke: read("__TAURI_INVOKE__"),
      };
    });

    // If the runtime injects globals as non-configurable properties, the hardening script should
    // lock them down to `undefined` so they can't be repopulated later.
    for (const [key, desc] of Object.entries(tauriDescriptors)) {
      if (!desc) continue;
      expect(desc.valueType, `${key} should be scrubbed to undefined`).toBe("undefined");
      expect(desc.writable, `${key} should not be writable after scrubbing`).toBe(false);
      expect(desc.configurable, `${key} should not be configurable after scrubbing`).toBe(false);
    }
  });

  test("runs sampleHello.sumSelection via the Extensions panel and shows a toast", async ({ page }) => {
    await gotoDesktop(page);
    await grantSampleHelloPermissions(page);

    // Production safety check: built-in extensions must not be fetched from the repo filesystem at
    // runtime (no Vite `/@fs/...` dependency). If the loader still uses
    // `BrowserExtensionHost.loadExtensionFromUrl(...)`, this would throw and the extension would
    // fail to load.
    await page.evaluate(() => {
      const originalFetch = window.fetch.bind(window);
      window.fetch = async (input: RequestInfo | URL, init?: RequestInit) => {
        const url =
          typeof input === "string"
            ? input
            : input instanceof URL
              ? input.toString()
              : typeof (input as any)?.url === "string"
                ? String((input as any).url)
                : String(input);
        if (url.includes("extensions/sample-hello/") && url.includes("package.json")) {
          throw new Error(`Blocked fetch for built-in extension asset: ${url}`);
        }
        return originalFetch(input, init);
      };
    });

    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();

      doc.setCellValue(sheetId, { row: 0, col: 0 }, 1);
      doc.setCellValue(sheetId, { row: 0, col: 1 }, 2);
      doc.setCellValue(sheetId, { row: 1, col: 0 }, 3);
      doc.setCellValue(sheetId, { row: 1, col: 1 }, 4);

      app.selectRange({
        sheetId,
        range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
      });
    });

    await openExtensionsPanel(page);
    const sumSelectionBtn = page.getByTestId("run-command-sampleHello.sumSelection");
    await expect(sumSelectionBtn).toBeVisible({ timeout: 30_000 });
    await sumSelectionBtn.dispatchEvent("click");

    await expect(page.getByTestId("toast-root")).toContainText("Sum: 10");
  });

  test("writes selection sum to the real clipboard via formula.clipboard.writeText", async ({ page }) => {
    await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
    await gotoDesktop(page);
    await grantSampleHelloPermissions(page);

    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();

      doc.setCellValue(sheetId, { row: 0, col: 0 }, 101);
      doc.setCellValue(sheetId, { row: 0, col: 1 }, 202);
      doc.setCellValue(sheetId, { row: 1, col: 0 }, 303);
      doc.setCellValue(sheetId, { row: 1, col: 1 }, 404);

      app.selectRange({
        sheetId,
        range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
      });
    });

    await openExtensionsPanel(page);
    const copyBtn = page.getByTestId("run-command-sampleHello.copySumToClipboard");
    await expect(copyBtn).toBeVisible({ timeout: 30_000 });
    // Avoid hit-target flakiness from fixed overlays by dispatching a click directly.
    await copyBtn.dispatchEvent("click");

    await expect
      .poll(() => page.evaluate(async () => (await navigator.clipboard.readText()).trim()), { timeout: 10_000 })
      .toBe("1010");
  });

  test("blocks sampleHello.copySumToClipboard when the selection is Restricted by DLP", async ({ page }) => {
    await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
    await gotoDesktop(page);
    await grantSampleHelloPermissions(page);

    const sentinel = `sentinel-${Date.now()}`;
    await page.evaluate(async (text) => {
      try {
        await navigator.clipboard.writeText(text);
        return;
      } catch {
        // Fall back to legacy DOM copy.
      }

      const textarea = document.createElement("textarea");
      textarea.value = text;
      textarea.style.position = "fixed";
      textarea.style.left = "-9999px";
      textarea.style.top = "0";
      document.body.appendChild(textarea);
      textarea.focus();
      textarea.select();
      const ok = document.execCommand("copy");
      textarea.remove();
      if (!ok) throw new Error("Failed to seed clipboard with sentinel text");
    }, sentinel);

    const workbookId = "local-workbook";

    await page.evaluate((docId) => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();

      doc.setCellValue(sheetId, { row: 0, col: 0 }, 1);
      doc.setCellValue(sheetId, { row: 0, col: 1 }, 2);
      doc.setCellValue(sheetId, { row: 1, col: 0 }, 3);
      doc.setCellValue(sheetId, { row: 1, col: 1 }, 4);

      app.selectRange({
        sheetId,
        range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
      });

      const record = {
        selector: {
          scope: "range",
          documentId: docId,
          sheetId,
          range: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } },
        },
        classification: { level: "Restricted" },
        updatedAt: new Date().toISOString(),
      };
      localStorage.setItem(`dlp:classifications:${docId}`, JSON.stringify([record]));
    }, workbookId);

    await openExtensionsPanel(page);
    const copyBtn = page.getByTestId("run-command-sampleHello.copySumToClipboard");
    await expect(copyBtn).toBeVisible({ timeout: 30_000 });
    await copyBtn.dispatchEvent("click");

    await expect(page.getByTestId("toast-root")).toContainText(
      "Clipboard copy is blocked by your organization's data loss prevention policy",
    );
    await expect(page.getByTestId("toast-root")).toContainText("Restricted");
    await expect(page.getByTestId("toast-root")).toContainText("Confidential");

    await expect
      .poll(() => page.evaluate(async () => (await navigator.clipboard.readText()).trim()), { timeout: 10_000 })
      .toBe(sentinel);
  });

  test("supports Sheet.getRange/setRange and sheet-qualified refs via the desktop spreadsheetApi adapter", async ({ page }) => {
    await gotoDesktop(page);
    await grantSampleHelloPermissions(page);

    // Pre-create Sheet2 so sheet-qualified refs resolve (desktop spreadsheetApi should error on unknown sheets).
    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      app.getDocument().setCellValue("Sheet2", "A1", "Hello from Sheet2");
    });

    await openExtensionsPanel(page);
    const rangeApiBtn = page.getByTestId("run-command-sampleHello.rangeApi");
    await expect(rangeApiBtn).toBeVisible({ timeout: 30_000 });
    await rangeApiBtn.dispatchEvent("click");

    await expect(page.getByTestId("toast-root")).toContainText("Range API: From extension / 123");

    const snapshot = await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheet1A1 = (doc.getCell("Sheet1", { row: 0, col: 0 }) as any)?.value ?? null;
      const sheet2B2 = (doc.getCell("Sheet2", { row: 1, col: 1 }) as any)?.value ?? null;
      return { sheet1A1, sheet2B2 };
    });

    expect(snapshot.sheet1A1).toBe("From extension");
    expect(snapshot.sheet2B2).toBe(123);
  });

  test("persists an extension panel in the layout and re-activates it after reload", async ({ page }) => {
    await gotoDesktop(page);
    await grantSampleHelloPermissions(page);

    await openExtensionsPanel(page);
    const openPanelBtn = page.getByTestId("run-command-sampleHello.openPanel");
    await expect(openPanelBtn).toBeVisible({ timeout: 30_000 });
    await openPanelBtn.dispatchEvent("click");

    await expect(page.getByTestId("panel-sampleHello.panel")).toBeAttached();
    const frame = page.frameLocator('iframe[data-testid="extension-webview-sampleHello.panel"]');
    await expect(frame.locator("h1")).toHaveText("Sample Hello Panel");

    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await grantSampleHelloPermissions(page);

    // Reloading the page resets the extension host; opening the command palette triggers
    // the lazy extension host boot so persisted extension panels can re-activate.
    const primary = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${primary}+Shift+P`);
    await expect(page.getByTestId("command-palette")).toBeVisible();
    await page.getByTestId("command-palette-input").fill("Sum Selection");
    // Command palette groups commands by category header (Sample Hello) and command title row.
    const list = page.getByTestId("command-palette-list");
    await expect(list).toContainText("Sample Hello", { timeout: 10_000 });
    await expect(list).toContainText("Sum Selection", { timeout: 10_000 });
    await page.keyboard.press("Escape");

    await expect(page.getByTestId("panel-sampleHello.panel")).toBeAttached();
    const frameAfter = page.frameLocator('iframe[data-testid="extension-webview-sampleHello.panel"]');
    await expect(frameAfter.locator("h1")).toHaveText("Sample Hello Panel");
  });

  test("executes a contributed keybinding when its when-clause matches", async ({ page }) => {
    await gotoDesktop(page);
    await grantSampleHelloPermissions(page);

    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();

      doc.setCellValue(sheetId, { row: 0, col: 0 }, 1);
      doc.setCellValue(sheetId, { row: 0, col: 1 }, 2);
      doc.setCellValue(sheetId, { row: 1, col: 0 }, 3);
      doc.setCellValue(sheetId, { row: 1, col: 1 }, 4);

      app.selectRange({
        sheetId,
        range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
      });
    });

    await openExtensionsPanel(page);
    await expect(page.getByTestId("run-command-sampleHello.sumSelection")).toBeVisible({ timeout: 30_000 });

    await page.keyboard.press("Control+Shift+Y");
    await expect(page.getByTestId("toast-root")).toContainText("Sum: 10");
  });

  test("executes a contributed shifted punctuation keybinding via KeyboardEvent.code fallback", async ({ page }) => {
    await gotoDesktop(page);
    await grantSampleHelloPermissions(page);

    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();

      doc.setCellValue(sheetId, { row: 0, col: 0 }, 1);
      doc.setCellValue(sheetId, { row: 0, col: 1 }, 2);
      doc.setCellValue(sheetId, { row: 1, col: 0 }, 3);
      doc.setCellValue(sheetId, { row: 1, col: 1 }, 4);

      app.selectRange({
        sheetId,
        range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
      });
    });

    await openExtensionsPanel(page);
    await expect(page.getByTestId("run-command-sampleHello.sumSelection")).toBeVisible({ timeout: 30_000 });

    // Simulate a shifted punctuation key where `event.key` changes ("'" -> "\""),
    // but the physical key stays stable via `event.code === "Quote"`.
    await page.evaluate(() => {
      const isMac = /mac/i.test(navigator.platform);
      const root = document.getElementById("grid");
      root?.dispatchEvent(
        new KeyboardEvent("keydown", {
          bubbles: true,
          cancelable: true,
          key: '"',
          code: "Quote",
          ctrlKey: !isMac,
          metaKey: isMac,
          shiftKey: true,
        }),
      );
    });

    await expect(page.getByTestId("toast-root")).toContainText("Sum: 10");
  });

  test("does not execute a keybinding when its when-clause fails", async ({ page }) => {
    await gotoDesktop(page);
    await grantSampleHelloPermissions(page);

    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.setCellValue(sheetId, { row: 0, col: 0 }, 5);
    });

    await openExtensionsPanel(page);
    await expect(page.getByTestId("run-command-sampleHello.sumSelection")).toBeVisible({ timeout: 30_000 });

    const before = await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      const cell = doc.getCell(sheetId, { row: 2, col: 0 }) as any;
      return cell?.value ?? null;
    });

    // Default selection is a single cell, so `hasSelection` should be false and the keybinding should be ignored.
    await page.keyboard.press("Control+Shift+Y");

    // Give the UI a brief moment in case a command mistakenly fires.
    await page.waitForTimeout(250);

    const after = await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      const cell = doc.getCell(sheetId, { row: 2, col: 0 }) as any;
      return cell?.value ?? null;
    });

    expect(after).toEqual(before);
  });

  test("loads extensions when opening the grid context menu", async ({ page }) => {
    await gotoDesktop(page);
    await grantSampleHelloPermissions(page);

    // Open the context menu without first opening the Extensions panel.
    await page.locator("#grid").click({ button: "right", position: { x: 100, y: 40 } });
    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();

    await page.waitForFunction(() => Boolean((window as any).__formulaExtensionHostManager?.ready), undefined, {
      timeout: 30_000,
    });
    const extensionLoadError = await page.evaluate(() => {
      const manager = (window as any).__formulaExtensionHostManager;
      const error = manager?.error;
      if (!error) return null;
      return {
        message: typeof error?.message === "string" ? error.message : String(error),
        stack: typeof error?.stack === "string" ? error.stack : null,
      };
    });
    if (extensionLoadError) {
      throw new Error(
        `Extension host failed to load during context menu open: ${extensionLoadError.message}\n${extensionLoadError.stack ?? ""}`,
      );
    }

    // Extension contributions should appear once the lazy-loaded extension host finishes
    // initializing.
    const item = menu.getByRole("button", { name: "Sample Hello: Open Sample Panel" });
    await expect(item).toBeVisible({ timeout: 30_000 });
    await expect(item, "context menu item should be clickable").toBeEnabled();
    await item.click();

    // Clicking the contributed item should invoke the extension command.
    await expect(page.getByTestId("panel-sampleHello.panel")).toBeAttached({ timeout: 30_000 });
  });

  test("executes a contributed context menu item when its when-clause matches", async ({ page }) => {
    await gotoDesktop(page);
    await grantSampleHelloPermissions(page);

    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();

      doc.setCellValue(sheetId, { row: 0, col: 0 }, 1);
      doc.setCellValue(sheetId, { row: 0, col: 1 }, 2);
      doc.setCellValue(sheetId, { row: 1, col: 0 }, 3);
      doc.setCellValue(sheetId, { row: 1, col: 1 }, 4);

      app.selectRange({
        sheetId,
        range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
      });
    });

    await openExtensionsPanel(page);
    await expect(page.getByTestId("run-command-sampleHello.sumSelection")).toBeVisible({ timeout: 30_000 });

    // Right-click inside the selection so the selection remains intact and `hasSelection` stays true.
    await page.evaluate(() => {
      const grid = document.getElementById("grid");
      if (!grid) throw new Error("Missing #grid container");
      const rect = grid.getBoundingClientRect();
      grid.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          button: 2,
          clientX: rect.left + 100,
          clientY: rect.top + 40,
        }),
      );
    });
    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();

    const item = menu.getByRole("button", { name: "Sample Hello: Sum Selection" });
    await expect(item).toBeEnabled({ timeout: 30_000 });
    await item.click();

    await expect(page.getByTestId("toast-root")).toContainText("Sum: 10");
  });

  test("right-clicking outside a multi-cell selection moves the active cell before showing the menu", async ({ page }) => {
    await gotoDesktop(page);
    await grantSampleHelloPermissions(page);

    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const sheetId = app.getCurrentSheetId();
      app.selectRange({
        sheetId,
        range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
      });
    });

    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect(page.getByTestId("selection-range")).toHaveText("A1:B2");

    // Ensure the extensions host is running so the contributed context menu renders.
    await openExtensionsPanel(page);
    await expect(page.getByTestId("run-command-sampleHello.openPanel")).toBeVisible({ timeout: 30_000 });

    // Ensure the grid has a usable hit-test surface. In headless e2e environments the
    // surrounding shell (ribbon/status bar) can leave the grid with near-zero layout
    // height, which makes `pickCellAtClientPoint` return null for all coordinates.
    await page.evaluate(() => {
      const grid = document.getElementById("grid");
      if (!grid) return;
      grid.style.height = "600px";
      grid.style.minHeight = "600px";
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      try {
        app?.onResize?.();
      } catch {
        // ignore
      }
    });

    // Wait for the grid renderer to fully initialize its viewport mapping so hit-testing
    // works reliably (otherwise `pickCellAtClientPoint` can report A1 for all points).
    await page.waitForFunction(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      const rect = app?.getCellRectA1?.("D4");
      return Boolean(rect && rect.width > 0 && rect.height > 0);
    });

    // Right-click a cell outside the current selection. Excel/Sheets move the active
    // cell to the clicked cell before showing the menu so commands apply to it.
    const d4Point = await page.evaluate(() => {
      const grid = document.getElementById("grid");
      if (!grid) throw new Error("Missing #grid container");
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (
        !app?.getCellRectA1 ||
        !app?.pickCellAtClientPoint ||
        typeof app.getCellRectA1 !== "function"
      ) {
        throw new Error("Missing required SpreadsheetApp test helpers");
      }

      const target = { row: 3, col: 3 };
      const rect = app.getCellRectA1("D4");
      if (!rect) throw new Error("Missing D4 rect");

      const gridRect = grid.getBoundingClientRect();
      // `getCellRectA1` is a test helper, but its coordinate space differs depending
      // on the underlying grid renderer. Use `pickCellAtClientPoint` to validate
      // which candidate coordinate maps back to D4.
      const candidates = [
        // Treat rect as already viewport-relative.
        { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 },
        // Treat rect as grid-root relative (need to add the grid's viewport offset).
        { x: gridRect.left + rect.x + rect.width / 2, y: gridRect.top + rect.y + rect.height / 2 },
      ];

      for (const point of candidates) {
        const picked = app.pickCellAtClientPoint(point.x, point.y);
        if (picked && picked.row === target.row && picked.col === target.col) return point;
      }

      const debug = {
        rect,
        gridRect: { left: gridRect.left, top: gridRect.top, width: gridRect.width, height: gridRect.height },
        picked: candidates.map((point) => ({ point, picked: app.pickCellAtClientPoint(point.x, point.y) })),
      };
      throw new Error(`Failed to locate D4 client coordinates for context menu test: ${JSON.stringify(debug)}`);
    });

    await page.evaluate((point) => {
      const grid = document.getElementById("grid");
      if (!grid) throw new Error("Missing #grid container");
      grid.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          button: 2,
          clientX: point.x,
          clientY: point.y,
        }),
      );
    }, d4Point);

    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();
    await expect(page.getByTestId("active-cell")).toHaveText("D4");
  });

  test("shared grid: right-click inside selection preserves it; outside selection moves active cell", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await grantSampleHelloPermissions(page);

    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const sheetId = app.getCurrentSheetId();
      app.selectRange({
        sheetId,
        range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
      });
    });

    // Ensure the extensions host is running so the contributed context menu renders.
    await openExtensionsPanel(page);
    await expect(page.getByTestId("panel-extensions")).toBeVisible();
    await expect(page.getByTestId("run-command-sampleHello.openPanel")).toBeVisible({ timeout: 30_000 });

    // Ensure the grid has a usable hit-test surface (headless environments can end up with a
    // near-zero grid height depending on viewport/layout).
    await page.evaluate(() => {
      const grid = document.getElementById("grid");
      if (!grid) return;
      grid.style.height = "600px";
      grid.style.minHeight = "600px";
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      try {
        app?.onResize?.();
      } catch {
        // ignore
      }
    });
    await page.waitForFunction(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      const rect = app?.getCellRectA1?.("D4");
      return Boolean(rect && rect.width > 0 && rect.height > 0);
    });

    // Right-click inside the selection on a non-active cell; selection should remain multi-cell.
    const b2Point = await page.evaluate(() => {
      const grid = document.getElementById("grid");
      if (!grid) throw new Error("Missing #grid container");
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app?.getCellRectA1 || !app?.pickCellAtClientPoint) {
        throw new Error("Missing required SpreadsheetApp test helpers");
      }

      const target = { row: 1, col: 1 };
      const rect = app.getCellRectA1("B2");
      if (!rect) throw new Error("Missing B2 rect");
      const gridRect = grid.getBoundingClientRect();

      const candidates = [
        { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 },
        { x: gridRect.left + rect.x + rect.width / 2, y: gridRect.top + rect.y + rect.height / 2 },
      ];

      for (const point of candidates) {
        const picked = app.pickCellAtClientPoint(point.x, point.y);
        if (picked && picked.row === target.row && picked.col === target.col) return point;
      }

      throw new Error("Failed to resolve B2 client coordinates for context menu test");
    });

    await page.evaluate((point) => {
      const grid = document.getElementById("grid");
      if (!grid) throw new Error("Missing #grid container");
      grid.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          button: 2,
          clientX: point.x,
          clientY: point.y,
        }),
      );
    }, b2Point);

    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();
    const sumItem = menu.getByRole("button", { name: "Sample Hello: Sum Selection" });
    await expect(sumItem, "inside selection should keep hasSelection=true").toBeEnabled();

    // Close the menu so we can open it again on a different cell.
    await page.keyboard.press("Escape");
    await expect(menu).toBeHidden();

    // Right-click outside the selection should move active cell (and collapse selection).
    const d4Point = await page.evaluate(() => {
      const grid = document.getElementById("grid");
      if (!grid) throw new Error("Missing #grid container");
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app?.getCellRectA1 || !app?.pickCellAtClientPoint) {
        throw new Error("Missing required SpreadsheetApp test helpers");
      }

      const target = { row: 3, col: 3 };
      const rect = app.getCellRectA1("D4");
      if (!rect) throw new Error("Missing D4 rect");
      const gridRect = grid.getBoundingClientRect();

      const candidates = [
        { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 },
        { x: gridRect.left + rect.x + rect.width / 2, y: gridRect.top + rect.y + rect.height / 2 },
      ];

      for (const point of candidates) {
        const picked = app.pickCellAtClientPoint(point.x, point.y);
        if (picked && picked.row === target.row && picked.col === target.col) return point;
      }

      throw new Error("Failed to resolve D4 client coordinates for context menu test");
    });

    await page.evaluate((point) => {
      const grid = document.getElementById("grid");
      if (!grid) throw new Error("Missing #grid container");
      grid.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          button: 2,
          clientX: point.x,
          clientY: point.y,
        }),
      );
    }, d4Point);

    await expect(menu).toBeVisible();
    await expect(page.getByTestId("active-cell")).toHaveText("D4");
    const sumItemAfter = menu.getByRole("button", { name: "Sample Hello: Sum Selection" });
    await expect(sumItemAfter, "outside selection should collapse selection (hasSelection=false)").toBeDisabled();
  });

  test("blocks extension clipboard writes after reading a Restricted cell via getCell()", async ({ page }) => {
    await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
    await gotoDesktop(page);

    const modifier = process.platform === "darwin" ? "Meta" : "Control";

    // Grant the test extension the permissions it needs before we load it.
    await page.evaluate(() => {
      const extensionId = "e2e.dlp-test";
      const key = "formula.extensionHost.permissions";
      const existing = (() => {
        try {
          const raw = localStorage.getItem(key);
          return raw ? JSON.parse(raw) : {};
        } catch {
          return {};
        }
      })();

      existing[extensionId] = {
        ...(existing[extensionId] ?? {}),
        "ui.commands": true,
        "cells.read": true,
        clipboard: true,
      };

      localStorage.setItem(key, JSON.stringify(existing));
    });

    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();

      doc.beginBatch({ label: "Seed DLP cells" });
      doc.setCellValue(sheetId, "A1", "SECRET");
      doc.setCellValue(sheetId, "B1", "SAFE");
      doc.endBatch();
      app.refresh();

      // Mark A1 Restricted via localStorage-backed DLP store.
      const documentId = "local-workbook";
      const record = {
        selector: { scope: "cell", documentId, sheetId, row: 0, col: 0 },
        classification: { level: "Restricted", labels: [] },
        updatedAt: new Date().toISOString(),
      };
      localStorage.setItem(`dlp:classifications:${documentId}`, JSON.stringify([record]));
    });
    await page.evaluate(() => (window as any).__formulaApp.whenIdle());

    // Copy B1 -> clipboard (baseline) using a real user gesture.
    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      const sheetId = app.getCurrentSheetId();
      app.activateCell({ sheetId, row: 0, col: 1 }); // B1
    });
    await expect(page.getByTestId("active-cell")).toHaveText("B1");
    await page.keyboard.press(`${modifier}+C`);
    await page.evaluate(() => (window as any).__formulaApp.whenIdle());

    const result = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const host: any = (window as any).__formulaExtensionHost;
      if (!host) throw new Error("Missing window.__formulaExtensionHost (desktop e2e harness)");

      const extensionId = "e2e.dlp-test";
      const commandId = "dlpTest.copyA1ToClipboard";
      const mainSource = `
        const api = globalThis[Symbol.for("formula.extensionApi.api")];
        export async function activate(context) {
          context.subscriptions.push(
            await api.commands.registerCommand(${JSON.stringify(commandId)}, async () => {
              try {
                const value = await api.cells.getCell(0, 0);
                await api.clipboard.writeText(String(value ?? ""));
                await api.ui.showMessage("Copied A1 to clipboard");
                return value;
              } catch (err) {
                const msg = String(err?.message ?? err);
                await api.ui.showMessage(msg, "error");
                throw err;
              }
            }),
          );
        }
      `;

      const mainUrl = URL.createObjectURL(new Blob([mainSource], { type: "text/javascript" }));
      const manifest = {
        name: "dlp-test",
        publisher: "e2e",
        version: "1.0.0",
        engines: { formula: "^1.0.0" },
        activationEvents: [`onCommand:${commandId}`],
        contributes: { commands: [{ command: commandId, title: "Copy A1 to Clipboard", category: "DLP Test" }] },
        permissions: ["cells.read", "clipboard", "ui.commands"],
      };

      try {
        await host.unloadExtension(extensionId);
      } catch {
        // ignore
      }

      await host.loadExtension({ extensionId, extensionPath: "memory://e2e/dlp-test", manifest, mainUrl });

      try {
        await host.executeCommand(commandId);
        return { ok: true };
      } catch (err) {
        return { ok: false, message: String((err as any)?.message ?? err) };
      }
    });

    expect(result.ok).toBe(false);
    await expect(page.getByTestId("toast-root")).toContainText("Clipboard copy is blocked");

    // Paste into C1: the clipboard should still contain the baseline "SAFE" value, not "SECRET".
    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      const sheetId = app.getCurrentSheetId();
      app.activateCell({ sheetId, row: 0, col: 2 }); // C1
    });
    await expect(page.getByTestId("active-cell")).toHaveText("C1");
    await page.keyboard.press(`${modifier}+V`);
    await page.evaluate(() => (window as any).__formulaApp.whenIdle());
    await expect
      .poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C1")))
      .toBe("SAFE");
  });
});
