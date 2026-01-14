import { expect, test } from "@playwright/test";

async function gotoPythonRuntime(page: import("@playwright/test").Page): Promise<void> {
  // Vite may occasionally trigger a one-time full reload after dependency optimization. Retry once
  // if the execution context is destroyed during startup.
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      await page.goto("/python-runtime-test.html", { waitUntil: "domcontentloaded" });
      await page.waitForFunction(() => Boolean((globalThis as any).__formulaPythonRuntime));
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

test("python-runtime: runs Pyodide with SharedArrayBuffer-backed formula bridge", async ({ page }) => {
  test.setTimeout(120_000);

  await gotoPythonRuntime(page);

  const runInPage = async () => {
    const { PyodideRuntime, MockWorkbook } = (globalThis as any).__formulaPythonRuntime;
    const indexURL = (globalThis as any).__pyodideIndexURL;

    const isolation = {
      crossOriginIsolated: globalThis.crossOriginIsolated,
      sharedArrayBuffer: typeof (globalThis as any).SharedArrayBuffer !== "undefined",
    };

    const workbook = new MockWorkbook();
    const runtime = new PyodideRuntime({
      api: workbook,
      indexURL,
      // The default is 2s; allow a bit more slack in CI.
      rpcTimeoutMs: 5_000,
    });

    try {
      await runtime.initialize();
      const execution = await runtime.execute(`import formula\nformula.active_sheet["A1"] = 1\n`);

      const sheet_id = workbook.get_active_sheet_id();
      const values = workbook.get_range_values({
        range: {
          sheet_id,
          start_row: 0,
          start_col: 0,
          end_row: 0,
          end_col: 0,
        },
      });

      return { isolation, execution, values };
    } finally {
      runtime.destroy();
    }
  };

  let result: any;
  // Vite may trigger a one-time full reload after dependency optimization. If
  // that happens mid-evaluate, retry once after the navigation completes.
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      result = await page.evaluate(runInPage);
      break;
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      if (attempt === 0 && message.includes("Execution context was destroyed")) {
        await page.waitForLoadState("domcontentloaded");
        await page.waitForFunction(() => Boolean((globalThis as any).__formulaPythonRuntime));
        continue;
      }
      throw err;
    }
  }

  expect(result).toBeTruthy();
  expect(result.isolation.crossOriginIsolated).toBe(true);
  expect(result.isolation.sharedArrayBuffer).toBe(true);
  expect(result.execution.success).toBe(true);
  expect(result.values).toEqual([[1]]);
});
