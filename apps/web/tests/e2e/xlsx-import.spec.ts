import { expect, test } from "@playwright/test";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

test("imports xlsx with formulas and evaluates dependent cells", async ({ page }) => {
  const fixturePath = path.resolve(__dirname, "../../../../fixtures/xlsx/formulas/formulas.xlsx");

  await page.goto("/");
  await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });

  const fileInput = page.getByTestId("xlsx-file-input");
  await fileInput.setInputFiles(fixturePath);

  await expect(page.getByTestId("engine-status")).toContainText("imported xlsx", { timeout: 30_000 });

  const grid = page.getByTestId("canvas-grid-selection");

  // B1 is a literal (2) in the fixture, distinct from the demo workbook's formula.
  await grid.click({ position: { x: 250, y: 31 } });
  await expect(page.getByTestId("active-address")).toHaveText("B1");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("2");

  // C1 contains a formula (=A1+B1) and should evaluate to 3 after recalc.
  await grid.click({ position: { x: 350, y: 31 } });
  await expect(page.getByTestId("active-address")).toHaveText("C1");
  await expect(page.getByTestId("formula-input")).toHaveValue("=A1+B1");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("3");
});

