import { expect, test } from "@playwright/test";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "../../../..");

test("can load real .xlsx bytes in the WASM worker engine and recalculate formulas", async ({ page }) => {
  const basicBase64 = (await readFile(path.join(repoRoot, "fixtures", "xlsx", "basic", "basic.xlsx"))).toString("base64");
  const formulasBase64 = (await readFile(path.join(repoRoot, "fixtures", "xlsx", "formulas", "formulas.xlsx"))).toString(
    "base64"
  );

  await page.addInitScript(() => {
    (globalThis as any).__FORMULA_E2E__ = true;
  });

  await page.goto("/");
  await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });

  const result = await page.evaluate(
    async ({ basicBase64, formulasBase64 }) => {
      const decode = (b64: string) => {
        const binary = atob(b64);
        const out = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i);
        return out;
      };

      try {
        const createEngineClient = (globalThis as any).__FORMULA_ENGINE_E2E__?.createEngineClient;
        if (typeof createEngineClient !== "function") {
          throw new Error(
            "E2E engine hook is missing. Ensure `globalThis.__FORMULA_E2E__ = true` is set before the app loads."
          );
        }

        const engine = createEngineClient();
        try {
          await engine.init();

          await engine.loadWorkbookFromXlsxBytes(decode(basicBase64));
          await engine.recalculate();
          const basicA1 = await engine.getCell("A1", "Sheet1");
          const basicB1 = await engine.getCell("B1", "Sheet1");

          await engine.loadWorkbookFromXlsxBytes(decode(formulasBase64));
          await engine.recalculate();
          const formulaA1 = await engine.getCell("A1", "Sheet1");
          const formulaB1 = await engine.getCell("B1", "Sheet1");
          const formulaC1 = await engine.getCell("C1", "Sheet1");

          // Verify sparse clear semantics: `null` inputs clear stored cells and are omitted from JSON.
          await engine.newWorkbook();
          await engine.setCell("A1", 1, "Sheet1");
          await engine.setCell("A2", "=A1*2", "Sheet1");
          await engine.recalculate();
          const beforeClearJson = await engine.toJson();

          await engine.setCell("A1", null, "Sheet1");
          await engine.recalculate();
          const afterClearA1 = await engine.getCell("A1", "Sheet1");
          const afterClearA2 = await engine.getCell("A2", "Sheet1");
          const afterClearJson = await engine.toJson();

          // Verify Goal Seek over the WASM worker API surface.
          await engine.newWorkbook();
          await engine.setCell("A1", 1, "Sheet1");
          await engine.setCell("B1", "=A1*A1", "Sheet1");
          await engine.recalculate();

          const goalSeek = (engine as any).goalSeek;
          if (typeof goalSeek !== "function") {
            throw new Error("engine.goalSeek is not available");
          }
          const goalSeekResult = await goalSeek({
            sheet: "Sheet1",
            targetCell: "B1",
            targetValue: 25,
            changingCell: "A1",
            tolerance: 1e-9,
          });
          await engine.recalculate();
          const goalSeekA1 = await engine.getCell("A1", "Sheet1");
          const goalSeekB1 = await engine.getCell("B1", "Sheet1");

          return {
            ok: true as const,
            basicA1,
            basicB1,
            formulaA1,
            formulaB1,
            formulaC1,
            beforeClearJson,
            afterClearA1,
            afterClearA2,
            afterClearJson,
            goalSeekResult,
            goalSeekA1,
            goalSeekB1,
          };
        } finally {
          engine.terminate();
        }
      } catch (error) {
        return {
          ok: false as const,
          error: error instanceof Error ? error.stack ?? error.message : String(error)
        };
      }
    },
    { basicBase64, formulasBase64 }
  );

  expect(result.ok, result.ok ? undefined : result.error).toBe(true);
  if (!result.ok) return;

  expect(result.basicA1.value).toBe(1);
  expect(result.basicB1.value).toBe("Hello");

  expect(result.formulaA1.value).toBe(1);
  expect(result.formulaB1.value).toBe(2);
  // XLSX stores formulas without a leading "=", but the engine should expose the canonical input.
  expect(result.formulaC1.input).toBe("=A1+B1");
  expect(result.formulaC1.value).toBe(3);

  const beforeClear = JSON.parse(result.beforeClearJson);
  expect(beforeClear.sheets.Sheet1.cells).toHaveProperty("A1", 1);
  expect(beforeClear.sheets.Sheet1.cells).toHaveProperty("A2", "=A1*2");

  expect(result.afterClearA1.input).toBeNull();
  expect(result.afterClearA1.value).toBeNull();
  expect(result.afterClearA2.value).toBe(0);

  const afterClear = JSON.parse(result.afterClearJson);
  expect(afterClear.sheets.Sheet1.cells).not.toHaveProperty("A1");
  expect(afterClear.sheets.Sheet1.cells).toHaveProperty("A2", "=A1*2");

  expect(result.goalSeekResult.success).toBe(true);
  expect(Math.abs(result.goalSeekResult.solution - 5)).toBeLessThan(1e-6);
  expect(Math.abs(result.goalSeekA1.value - 5)).toBeLessThan(1e-6);
  expect(Math.abs(result.goalSeekB1.value - 25)).toBeLessThan(1e-6);
});
