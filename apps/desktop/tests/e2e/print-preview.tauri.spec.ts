import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("Print Preview (tauri)", () => {
  test("File → Print Preview opens the preview dialog and allows downloading the PDF", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, Array<(event: any) => void>> = {};
      const invokes: Array<{ cmd: string; args: any }> = [];
      const downloadClicks: Array<{ download: string | null; href: string | null }> = [];
      let printCalls = 0;

      (window as any).__tauriListeners = listeners;
      (window as any).__tauriInvokes = invokes;
      (window as any).__downloadClicks = downloadClicks;
      (window as any).__printCalls = () => printCalls;

      // Avoid modal prompts blocking the test.
      window.confirm = () => true;

      // Avoid creating real browser downloads; capture intent instead.
      HTMLAnchorElement.prototype.click = function click() {
        try {
          downloadClicks.push({
            download: typeof (this as any).download === "string" ? (this as any).download : null,
            href: typeof (this as any).href === "string" ? (this as any).href : null,
          });
        } catch {
          // ignore
        }
      };

      // Best-effort: avoid opening a real OS print dialog in CI/headless. Our print preview
      // implementation calls `iframe.contentWindow.print()`. Stub it only for the print preview
      // iframe (identified by being inside `dialog.print-preview-dialog`).
      try {
        const original = Object.getOwnPropertyDescriptor(HTMLIFrameElement.prototype, "contentWindow");
        if (original && typeof original.get === "function") {
          Object.defineProperty(HTMLIFrameElement.prototype, "contentWindow", {
            configurable: true,
            enumerable: original.enumerable,
            get: function () {
              try {
                const el = this as any;
                const isPrintPreview = typeof el?.closest === "function" && el.closest("dialog.print-preview-dialog");
                if (isPrintPreview) {
                  return {
                    print: () => {
                      printCalls += 1;
                    },
                    focus: () => {},
                  };
                }
              } catch {
                // ignore
              }
              return original.get!.call(this);
            },
          });
        }
      } catch {
        // ignore
      }

      const windowHandle = {
        hide: async () => {},
        close: async () => {},
        show: async () => {},
        setFocus: async () => {},
      };

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            invokes.push({ cmd, args });

            switch (cmd) {
              case "get_sheet_print_settings":
                return {
                  sheet_name: args?.sheet_id ?? "Sheet1",
                  print_area: null,
                  print_titles: null,
                  page_setup: {
                    orientation: "portrait",
                    paper_size: 1,
                    margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
                    scaling: { kind: "percent", percent: 100 },
                  },
                  manual_page_breaks: { row_breaks_after: [], col_breaks_after: [] },
                };

              case "export_sheet_range_pdf":
                // A tiny placeholder payload; need not be a valid PDF for frontend wiring tests.
                return btoa("pdf");

              // Minimal workbook backend stubs so the app can boot in "tauri" mode.
              case "list_defined_names":
              case "list_tables":
              case "get_workbook_theme_palette":
              case "get_macro_security_status":
              case "set_macro_ui_context":
              case "fire_workbook_open":
                return null;

              default:
                return null;
            }
          },
        },
        event: {
          listen: async (name: string, handler: any) => {
            if (!Array.isArray(listeners[name])) listeners[name] = [];
            listeners[name].push(handler);
            return () => {
              const arr = listeners[name];
              if (!Array.isArray(arr)) return;
              const idx = arr.indexOf(handler);
              if (idx >= 0) arr.splice(idx, 1);
            };
          },
          emit: async () => {},
        },
        dialog: {
          open: async () => null,
          save: async () => null,
        },
        window: {
          // Provide all common handle accessors used by our Tauri abstractions so this
          // test stays resilient to future refactors.
          getCurrentWebviewWindow: () => windowHandle,
          getCurrentWindow: () => windowHandle,
          getCurrent: () => windowHandle,
          appWindow: windowHandle,
        },
        notification: {
          notify: async () => {},
        },
      };
    });

    await gotoDesktop(page);

    const ribbon = page.getByTestId("ribbon-root");
    await expect(ribbon).toBeVisible();

    await ribbon.getByRole("tab", { name: "File" }).click();
    await page.getByTestId("file-print-preview").click();

    await page.waitForFunction(() => (window as any).__tauriInvokes?.some((e: any) => e?.cmd === "export_sheet_range_pdf"));
    await expect(page.locator("dialog.print-preview-dialog")).toBeVisible();

    await page.locator("dialog.print-preview-dialog").getByRole("button", { name: "Print" }).click();
    await page.waitForFunction(() => typeof (window as any).__printCalls === "function" && (window as any).__printCalls() > 0);

    await page.locator("dialog.print-preview-dialog").getByRole("button", { name: "Download PDF" }).click();
    await page.waitForFunction(() => Array.isArray((window as any).__downloadClicks) && (window as any).__downloadClicks.length > 0);
    const download = await page.evaluate(() => (window as any).__downloadClicks?.[(window as any).__downloadClicks.length - 1] ?? null);
    expect(download?.download).toMatch(/\.pdf$/i);
  });
});
