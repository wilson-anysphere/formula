# Testing & Validation Strategy

## Overview

Testing a spreadsheet application requires unprecedented rigor. Excel has 40 years of edge cases, and users have built careers on its exact behavior. We must validate formula compatibility, file format fidelity, and performance at scale.

---

## Testing Pyramid

```
                    ┌─────────────┐
                    │   E2E       │  ~100 tests
                    │   Tests     │  Real user flows
                    ├─────────────┤
                    │ Integration │  ~1,000 tests
                    │   Tests     │  Component interaction
                    ├─────────────┤
                    │    Unit     │  ~10,000 tests
                    │   Tests     │  Individual functions
                    └─────────────┘
```

---

## Unit Testing

### Formula Function Tests

Every Excel function needs comprehensive test coverage:

```typescript
// tests/functions/sum.test.ts

describe("SUM function", () => {
  describe("basic behavior", () => {
    it("sums numbers", () => {
      expect(evaluate("=SUM(1, 2, 3)")).toBe(6);
    });
    
    it("sums range of numbers", () => {
      const sheet = createSheet([
        [1], [2], [3], [4], [5]
      ]);
      expect(evaluate("=SUM(A1:A5)", sheet)).toBe(15);
    });
    
    it("handles empty cells as zero", () => {
      const sheet = createSheet([
        [1], [null], [3]
      ]);
      expect(evaluate("=SUM(A1:A3)", sheet)).toBe(4);
    });
  });
  
  describe("type coercion", () => {
    it("ignores text in ranges", () => {
      const sheet = createSheet([
        [1], ["text"], [3]
      ]);
      expect(evaluate("=SUM(A1:A3)", sheet)).toBe(4);
    });
    
    it("converts text arguments to numbers", () => {
      expect(evaluate('=SUM("5", 3)')).toBe(8);
    });
    
    it("treats TRUE as 1", () => {
      expect(evaluate("=SUM(TRUE, 3)")).toBe(4);
    });
    
    it("treats FALSE as 0", () => {
      expect(evaluate("=SUM(FALSE, 3)")).toBe(3);
    });
    
    it("returns #VALUE! for non-numeric text argument", () => {
      expect(evaluate('=SUM("abc", 3)')).toBeError("#VALUE!");
    });
  });
  
  describe("edge cases", () => {
    it("handles large numbers", () => {
      expect(evaluate("=SUM(1E308, 1E308)")).toBe(Infinity);
    });
    
    it("handles negative numbers", () => {
      expect(evaluate("=SUM(-1, -2, 3)")).toBe(0);
    });
    
    it("handles zero arguments", () => {
      expect(evaluate("=SUM()")).toBe(0);
    });
    
    it("handles 255 arguments (max)", () => {
      const args = Array(255).fill(1).join(",");
      expect(evaluate(`=SUM(${args})`)).toBe(255);
    });
    
    it("propagates errors", () => {
      const sheet = createSheet([
        [1], ["=1/0"], [3]
      ]);
      expect(evaluate("=SUM(A1:A3)", sheet)).toBeError("#DIV/0!");
    });
  });
});
```

### Formula Parser Tests

```typescript
// tests/parser/parser.test.ts

describe("Formula Parser", () => {
  describe("cell references", () => {
    it("parses A1 notation", () => {
      const ast = parse("=A1");
      expect(ast).toEqual({
        type: "CellRef",
        row: 0,
        col: 0,
        rowAbsolute: false,
        colAbsolute: false
      });
    });
    
    it("parses absolute references", () => {
      const ast = parse("=$A$1");
      expect(ast.rowAbsolute).toBe(true);
      expect(ast.colAbsolute).toBe(true);
    });
    
    it("parses mixed references", () => {
      const ast = parse("=A$1");
      expect(ast.rowAbsolute).toBe(true);
      expect(ast.colAbsolute).toBe(false);
    });
    
    it("parses sheet references", () => {
      const ast = parse("=Sheet2!A1");
      expect(ast.sheet).toBe("Sheet2");
    });
    
    it("parses quoted sheet names", () => {
      const ast = parse("='My Sheet'!A1");
      expect(ast.sheet).toBe("My Sheet");
    });
  });
  
  describe("structured references", () => {
    it("parses table column reference", () => {
      const ast = parse("=Table1[Column1]");
      expect(ast).toMatchObject({
        type: "StructuredRef",
        table: "Table1",
        column: "Column1"
      });
    });
    
    it("parses current row reference", () => {
      const ast = parse("=[@Column1]");
      expect(ast).toMatchObject({
        type: "StructuredRef",
        column: "Column1",
        currentRow: true
      });
    });
  });
  
  describe("operator precedence", () => {
    it("respects multiplication over addition", () => {
      expect(evaluate("=2+3*4")).toBe(14);
    });
    
    it("respects exponentiation over multiplication", () => {
      expect(evaluate("=2*3^2")).toBe(18);
    });
    
    it("respects parentheses", () => {
      expect(evaluate("=(2+3)*4")).toBe(20);
    });
  });
});
```

### Sheet Rename / Sheet Reference Rewrite Tests

Sheet rename is a **structural edit** that must update formulas everywhere in the workbook. We need focused unit tests that validate the rewrite logic for the tricky Excel cases:
- Unquoted vs quoted sheet names (`Sheet1!A1` vs `'My Sheet'!A1`)
- Escaped apostrophes in quoted names (`'O''Brien'!A1`)
- 3D references (`Sheet1:Sheet3!A1`)
- External workbook prefixes (`'[Book1.xlsx]Sheet1'!A1`)
- Ensure **string literals** are not modified (`="Sheet1!A1"`)

```typescript
describe("Sheet reference rewrite", () => {
  it("rewrites a simple sheet reference", () => {
    expect(rewriteSheetRefs("=Sheet1!A1", "Sheet1", "Data")).toBe("=Data!A1");
  });

  it("rewrites a quoted sheet reference", () => {
    expect(rewriteSheetRefs("='Sheet 1'!A1", "Sheet 1", "My Sheet")).toBe("='My Sheet'!A1");
  });

  it("does not touch string literals", () => {
    expect(rewriteSheetRefs('="Sheet1!A1"', "Sheet1", "Data")).toBe('="Sheet1!A1"');
  });
});
```

### Dependency Graph Tests

```typescript
// tests/engine/dependency.test.ts

describe("Dependency Graph", () => {
  it("tracks simple dependencies", () => {
    const engine = new CalcEngine();
    engine.setCell("A1", 10);
    engine.setCell("A2", "=A1*2");
    
    expect(engine.getDependents("A1")).toContain("A2");
    expect(engine.getPrecedents("A2")).toContain("A1");
  });
  
  it("handles range dependencies", () => {
    const engine = new CalcEngine();
    engine.setCell("A1", 1);
    engine.setCell("A2", 2);
    engine.setCell("A3", "=SUM(A1:A2)");
    
    expect(engine.getDependents("A1")).toContain("A3");
    expect(engine.getDependents("A2")).toContain("A3");
  });
  
  it("detects circular references", () => {
    const engine = new CalcEngine();
    engine.setCell("A1", "=B1");
    
    expect(() => {
      engine.setCell("B1", "=A1");
    }).toThrow("Circular reference detected");
  });
  
  it("recalculates in correct order", () => {
    const engine = new CalcEngine();
    engine.setCell("A1", 10);
    engine.setCell("A2", "=A1*2");
    engine.setCell("A3", "=A2*2");
    
    engine.setCell("A1", 20);
    engine.recalculate();
    
    expect(engine.getValue("A2")).toBe(40);
    expect(engine.getValue("A3")).toBe(80);
  });
});
```

---

## Integration Testing

### File I/O Tests

```typescript
// tests/integration/xlsx.test.ts

describe("XLSX Round-Trip", () => {
  const testFiles = [
    "simple-data.xlsx",
    "formulas.xlsx",
    "formatting.xlsx",
    "pivot-tables.xlsx",
    "conditional-formatting.xlsx",
    "data-validation.xlsx",
    "large-file.xlsx",
    "macros.xlsm"
  ];

  // Chart coverage should come from a dedicated fixture set (one workbook per
  // chart family/type) so we can validate both preservation and rendering.
  const chartFixtures = globSync("charts/*.xlsx", { cwd: "fixtures" });
  
  [...testFiles, ...chartFixtures].forEach(filename => {
    it(`round-trips ${filename} without data loss`, async () => {
      const original = await readFile(`fixtures/${filename}`);
      const workbook = await loadWorkbook(original);
      const saved = await saveWorkbook(workbook);
      const reloaded = await loadWorkbook(saved);
      
      // Compare structure
      expect(reloaded.sheets.length).toBe(workbook.sheets.length);
      
      // Compare cell values
      for (const sheet of workbook.sheets) {
        const reloadedSheet = reloaded.sheets.find(s => s.name === sheet.name);
        expect(reloadedSheet).toBeDefined();
        
        for (const [cellId, cell] of sheet.cells) {
          const reloadedCell = reloadedSheet.cells.get(cellId);
          expect(reloadedCell?.value).toEqual(cell.value);
          expect(reloadedCell?.formula).toBe(cell.formula);
        }
      }
    });
  });
  
  it("preserves VBA project binary", async () => {
    const original = await readFile("fixtures/macros.xlsm");
    const workbook = await loadWorkbook(original);
    const saved = await saveWorkbook(workbook);
    
    const originalVBA = extractVBAProject(original);
    const savedVBA = extractVBAProject(saved);
    
    expect(savedVBA).toEqual(originalVBA);
  });
  
  it("preserves chart parts byte-for-byte (no-op save)", async () => {
    const original = await readFile("fixtures/charts/waterfall.xlsx");
    const workbook = await loadWorkbook(original);

    // Important: if we didn't edit charts, saving must keep the original
    // chart-related OPC parts intact (including extension lists / ChartEx).
    const saved = await saveWorkbook(workbook);

    expect(extractOpcParts(saved, { prefix: "xl/charts/" }))
      .toEqual(extractOpcParts(original, { prefix: "xl/charts/" }));

    expect(extractOpcParts(saved, { prefix: "xl/drawings/" }))
      .toEqual(extractOpcParts(original, { prefix: "xl/drawings/" }));
  });
});
```

### Collaboration Tests

```typescript
// tests/integration/collaboration.test.ts

describe("Real-Time Collaboration", () => {
  let server: TestServer;
  let client1: TestClient;
  let client2: TestClient;
  
  beforeEach(async () => {
    server = await createTestServer();
    client1 = await createTestClient(server);
    client2 = await createTestClient(server);
    
    // Both clients open same document
    await client1.openDocument("test-doc");
    await client2.openDocument("test-doc");
    await waitForSync();
  });
  
  afterEach(async () => {
    await client1.disconnect();
    await client2.disconnect();
    await server.close();
  });
  
  it("syncs cell changes between clients", async () => {
    await client1.setCell("A1", "Hello");
    await waitForSync();
    
    expect(await client2.getCell("A1")).toBe("Hello");
  });
  
  it("handles concurrent edits to different cells", async () => {
    // Simultaneous edits
    await Promise.all([
      client1.setCell("A1", "From Client 1"),
      client2.setCell("B1", "From Client 2")
    ]);
    
    await waitForSync();
    
    // Both changes should be present
    expect(await client1.getCell("A1")).toBe("From Client 1");
    expect(await client1.getCell("B1")).toBe("From Client 2");
    expect(await client2.getCell("A1")).toBe("From Client 1");
    expect(await client2.getCell("B1")).toBe("From Client 2");
  });
  
  it("handles concurrent edits to same cell (last write wins)", async () => {
    // Simulate network delay
    client2.setNetworkDelay(100);
    
    await client1.setCell("A1", "First");
    await client2.setCell("A1", "Second");
    
    await waitForSync();
    
    // Both clients should see same value
    const value1 = await client1.getCell("A1");
    const value2 = await client2.getCell("A1");
    expect(value1).toBe(value2);
  });
  
  it("recovers from offline mode", async () => {
    // Client 2 goes offline
    client2.goOffline();
    
    // Both make changes
    await client1.setCell("A1", "Online");
    await client2.setCell("A2", "Offline");
    
    // Client 2 comes back online
    await client2.goOnline();
    await waitForSync();
    
    // Both changes should merge
    expect(await client1.getCell("A1")).toBe("Online");
    expect(await client1.getCell("A2")).toBe("Offline");
    expect(await client2.getCell("A1")).toBe("Online");
    expect(await client2.getCell("A2")).toBe("Offline");
  });
});
```

---

## End-to-End Testing

### User Flow Tests

```typescript
// tests/e2e/user-flows.test.ts

import { test, expect } from "@playwright/test";

test.describe("User Flows", () => {
  test("creates new workbook and enters data", async ({ page }) => {
    await page.goto("/");
    
    // New workbook button
    await page.click('[data-testid="new-workbook"]');
    
    // Enter data
    await page.click('[data-cell="A1"]');
    await page.keyboard.type("Revenue");
    await page.keyboard.press("Tab");
    await page.keyboard.type("100");
    await page.keyboard.press("Enter");
    
    // Verify
    expect(await page.textContent('[data-cell="A1"]')).toBe("Revenue");
    expect(await page.textContent('[data-cell="B1"]')).toBe("100");
  });

  test("manages sheets (add/rename/reorder) and keeps formulas correct", async ({ page }) => {
    await page.goto("/");

    // Create a second sheet
    await page.click('[data-testid="sheet-add"]'); // "+" button in sheet tab strip
    await expect(page.locator('[data-testid="sheet-tab-Sheet2"]')).toBeVisible();

    // Put data on Sheet2
    await page.click('[data-testid="sheet-tab-Sheet2"]');
    await page.click('[data-cell="A1"]');
    await page.keyboard.type("10");

    // Reference Sheet2 from Sheet1
    await page.click('[data-testid="sheet-tab-Sheet1"]');
    await page.click('[data-cell="B1"]');
    await page.keyboard.type("=Sheet2!A1+5");
    await page.keyboard.press("Enter");
    await expect(page.locator('[data-cell="B1"]')).toHaveText("15");

    // Rename Sheet2 -> Data (double click)
    await page.dblclick('[data-testid="sheet-tab-Sheet2"]');
    await page.keyboard.press("Control+A");
    await page.keyboard.type("Data");
    await page.keyboard.press("Enter");
    await expect(page.locator('[data-testid="sheet-tab-Data"]')).toBeVisible();

    // Formula should rewrite and still compute
    await page.click('[data-cell="B1"]');
    await expect(page.locator('[data-testid="formula-bar"]')).toHaveValue("=Data!A1+5");
    await expect(page.locator('[data-cell="B1"]')).toHaveText("15");
  });
  
  test("opens Excel file and calculates", async ({ page }) => {
    await page.goto("/");
    
    // Open file
    const fileChooser = await page.waitForFileChooser();
    await page.click('[data-testid="open-file"]');
    await fileChooser.setFiles("fixtures/budget.xlsx");
    
    // Wait for load
    await page.waitForSelector('[data-sheet="Budget"]');
    
    // Verify calculation
    const total = await page.textContent('[data-cell="B10"]');
    expect(total).toBe("$125,000");
  });
  
  test("uses AI to generate formula", async ({ page }) => {
    await page.goto("/");
    
    // Enter data
    await page.fill('[data-cell="A1"]', "Sales");
    await page.fill('[data-cell="A2"]', "100");
    await page.fill('[data-cell="A3"]', "200");
    await page.fill('[data-cell="A4"]', "300");
    
    // Open AI panel
    const isMac = process.platform === "darwin";
    await page.keyboard.press(isMac ? "Meta+I" : "Control+Shift+A");
    
    // Ask AI
    await page.fill('[data-testid="ai-input"]', "Sum the sales");
    await page.keyboard.press("Enter");
    
    // Wait for response
    await page.waitForSelector('[data-testid="ai-suggestion"]');
    
    // Accept suggestion
    await page.click('[data-testid="accept-suggestion"]');
    
    // Verify formula was inserted
    expect(await page.textContent('[data-cell="A5"]')).toBe("600");
  });
});
```

### Visual Regression Tests

```typescript
// tests/e2e/visual.test.ts

import { test, expect } from "@playwright/test";

test.describe("Visual Regression", () => {
  test("grid renders correctly", async ({ page }) => {
    await page.goto("/");
    await page.setViewportSize({ width: 1280, height: 800 });
    
    // Load test data
    await loadTestWorkbook(page, "visual-test.xlsx");
    
    // Take screenshot
    expect(await page.screenshot()).toMatchSnapshot("grid-default.png");
  });
  
  test("selection renders correctly", async ({ page }) => {
    await page.goto("/");
    
    // Select range
    await page.click('[data-cell="B2"]');
    await page.keyboard.down("Shift");
    await page.click('[data-cell="D5"]');
    await page.keyboard.up("Shift");
    
    expect(await page.screenshot()).toMatchSnapshot("selection-range.png");
  });
  
  test("conditional formatting renders correctly", async ({ page }) => {
    await page.goto("/");
    await loadTestWorkbook(page, "conditional-formatting.xlsx");
    
    expect(await page.screenshot()).toMatchSnapshot("conditional-formatting.png");
  });

  test("charts render consistently", async ({ page }) => {
    await page.goto("/");
    await page.setViewportSize({ width: 1280, height: 800 });

    // Each fixture should contain a single chart anchored in a predictable
    // location. Snapshots should be compared against golden images exported
    // from Excel.
    await loadTestWorkbook(page, "charts/waterfall.xlsx");

    expect(await page.screenshot()).toMatchSnapshot("chart-waterfall.png");
  });
});
```

---

## Excel Compatibility Testing

### Excel Function Compatibility Matrix

```typescript
// tests/compatibility/excel-functions.test.ts

import { excelTestCases } from "./excel-test-data";

describe("Excel Function Compatibility", () => {
  // Generated from Excel itself
  excelTestCases.forEach(({ formula, inputs, expected, excelVersion }) => {
    it(`${formula} matches Excel ${excelVersion}`, () => {
      const sheet = createSheetFromInputs(inputs);
      const result = evaluate(formula, sheet);
      
      if (typeof expected === "number") {
        expect(result).toBeCloseTo(expected, 10);
      } else {
        expect(result).toEqual(expected);
      }
    });
  });
});

// Generate test cases by running formulas in Excel
async function generateExcelTestCases() {
  const excel = await connectToExcel();
  const testCases = [];
  
  for (const formula of formulasToTest) {
    for (const inputs of inputVariations) {
      const result = await excel.evaluate(formula, inputs);
      testCases.push({
        formula,
        inputs,
        expected: result,
        excelVersion: excel.version
      });
    }
  }
  
  return testCases;
}
```

**Implementation note:** This repository includes a concrete Excel-oracle harness under
`tools/excel-oracle/` (Excel COM automation + JSON export) with a curated corpus under
`tests/compatibility/excel-oracle/`. This is the intended mechanism for generating and
maintaining Excel-validated expected results over time.

### Cross-Application File Testing

```typescript
// tests/compatibility/cross-app.test.ts

describe("Cross-Application Compatibility", () => {
  const apps = ["excel-windows", "excel-mac", "google-sheets", "libreoffice"];
  
  apps.forEach(sourceApp => {
    apps.forEach(targetApp => {
      if (sourceApp === targetApp) return;
      
      it(`files created in ${sourceApp} open correctly in ${targetApp}`, async () => {
        const testFile = `fixtures/created-in-${sourceApp}.xlsx`;
        
        // Open in our app
        const workbook = await loadWorkbook(testFile);
        
        // Verify key properties
        expect(workbook.sheets.length).toBeGreaterThan(0);
        
        // Verify formulas calculate
        const calculatedCells = workbook.sheets[0].cells
          .filter(c => c.formula)
          .filter(c => c.value !== null);
        
        expect(calculatedCells.length).toBeGreaterThan(0);
      });
    });
  });
});
```

---

## Performance Testing

### Benchmark Suite

```typescript
// tests/performance/benchmarks.test.ts

import { benchmark } from "./benchmark-utils";

describe("Performance Benchmarks", () => {
  describe("Startup", () => {
    benchmark("cold start", async () => {
      const start = performance.now();
      await launchApp();
      const end = performance.now();
      
      expect(end - start).toBeLessThan(1000); // < 1 second
    });
    
    benchmark("open 1MB file", async () => {
      const app = await launchApp();
      
      const start = performance.now();
      await app.openFile("fixtures/1mb-file.xlsx");
      const end = performance.now();
      
      expect(end - start).toBeLessThan(3000); // < 3 seconds
    });
  });
  
  describe("Calculation", () => {
    benchmark("recalculate 100K cells", async () => {
      const engine = new CalcEngine();
      
      // Set up 100K cells with formulas
      for (let row = 0; row < 1000; row++) {
        for (let col = 0; col < 100; col++) {
          engine.setCell(row, col, `=ROW()*COL()`);
        }
      }
      
      const start = performance.now();
      await engine.recalculate();
      const end = performance.now();
      
      expect(end - start).toBeLessThan(100); // < 100ms
    });
    
    benchmark("VLOOKUP 10K rows", async () => {
      const engine = new CalcEngine();
      
      // Set up lookup table
      for (let row = 0; row < 10000; row++) {
        engine.setCell(row, 0, row);
        engine.setCell(row, 1, `Value${row}`);
      }
      
      // Set up 1000 lookups
      for (let row = 0; row < 1000; row++) {
        engine.setCell(row, 3, `=VLOOKUP(${row * 10}, A:B, 2, FALSE)`);
      }
      
      const start = performance.now();
      await engine.recalculate();
      const end = performance.now();
      
      expect(end - start).toBeLessThan(500); // < 500ms
    });
  });
  
  describe("Rendering", () => {
    benchmark("scroll 60fps with 1M rows", async () => {
      const renderer = new GridRenderer();
      renderer.setData(generateLargeDataset(1_000_000, 100));
      
      const frameTimes: number[] = [];
      
      // Simulate scroll
      for (let i = 0; i < 60; i++) {
        const start = performance.now();
        renderer.scrollTo(0, i * 100);
        renderer.render();
        frameTimes.push(performance.now() - start);
        await new Promise(r => requestAnimationFrame(r));
      }
      
      const avgFrameTime = frameTimes.reduce((a, b) => a + b) / frameTimes.length;
      expect(avgFrameTime).toBeLessThan(16.67); // 60fps
    });
  });
  
  describe("Memory", () => {
    benchmark("memory for 100MB file", async () => {
      const memBefore = process.memoryUsage().heapUsed;
      
      await loadWorkbook("fixtures/100mb-file.xlsx");
      
      // Force GC if available
      if (global.gc) global.gc();
      
      const memAfter = process.memoryUsage().heapUsed;
      const memUsed = (memAfter - memBefore) / (1024 * 1024);
      
      expect(memUsed).toBeLessThan(500); // < 500MB
    });
  });
});
```

### Load Testing

```typescript
// tests/performance/load.test.ts

describe("Load Testing", () => {
  it("handles 100 concurrent users", async () => {
    const server = await startServer();
    const clients: TestClient[] = [];
    
    // Create 100 clients
    for (let i = 0; i < 100; i++) {
      clients.push(await createClient(server));
    }
    
    // All open same document
    await Promise.all(clients.map(c => c.openDocument("shared-doc")));
    
    // All make concurrent edits
    const editPromises = clients.map((client, i) => 
      client.setCell(`A${i + 1}`, `Client ${i}`)
    );
    
    const start = performance.now();
    await Promise.all(editPromises);
    await waitForAllSynced(clients);
    const end = performance.now();
    
    // All edits should complete within 5 seconds
    expect(end - start).toBeLessThan(5000);
    
    // All clients should have same state
    for (let i = 0; i < 100; i++) {
      for (const client of clients) {
        expect(await client.getCell(`A${i + 1}`)).toBe(`Client ${i}`);
      }
    }
  });
});
```

---

## AI Testing

### AI Response Quality

```typescript
// tests/ai/response-quality.test.ts

describe("AI Response Quality", () => {
  describe("Formula Generation", () => {
    const testCases = [
      {
        prompt: "Sum column A",
        expectedContains: "SUM(A:A)",
        expectedType: "formula"
      },
      {
        prompt: "Average of B1 to B10",
        expectedContains: "AVERAGE(B1:B10)",
        expectedType: "formula"
      },
      {
        prompt: "Look up value in column A and return column B",
        expectedContains: ["VLOOKUP", "XLOOKUP", "INDEX", "MATCH"],
        expectedType: "formula"
      }
    ];
    
    testCases.forEach(({ prompt, expectedContains, expectedType }) => {
      it(`generates correct formula for: ${prompt}`, async () => {
        const response = await aiEngine.generate(prompt, testContext);
        
        expect(response.type).toBe(expectedType);
        
        if (Array.isArray(expectedContains)) {
          expect(expectedContains.some(e => response.formula.includes(e))).toBe(true);
        } else {
          expect(response.formula).toContain(expectedContains);
        }
        
        // Verify formula is valid
        expect(() => parseFormula(response.formula)).not.toThrow();
      });
    });
  });
  
  describe("Data Analysis", () => {
    it("correctly identifies trends", async () => {
      const data = [
        ["Month", "Sales"],
        ["Jan", 100],
        ["Feb", 120],
        ["Mar", 150],
        ["Apr", 180]
      ];
      
      const response = await aiEngine.analyze("What's the trend?", data);
      
      expect(response.toLowerCase()).toContain("upward");
      expect(response.toLowerCase()).toContain("increasing");
    });
  });
});
```

---

## Test Infrastructure

### CI/CD Pipeline

```yaml
# .github/workflows/test.yml
name: Tests

on: [push, pull_request]

env:
  # Keep in sync with `.nvmrc` and the pinned CI/release workflows to avoid drift.
  NODE_VERSION: 22

jobs:
  unit-tests:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with:
          # Keep Node pinned to the same major used by CI/release workflows.
          node-version: ${{ env.NODE_VERSION }}
      - run: npm ci
      - run: npm run test:unit
      - uses: codecov/codecov-action@v3

  integration-tests:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with:
          node-version: ${{ env.NODE_VERSION }}
      # Pin Rust for deterministic builds (this repo uses rust-toolchain.toml).
      - uses: dtolnay/rust-toolchain@1.92.0
      - run: npm ci
      - run: npm run test:integration

  e2e-tests:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with:
          node-version: ${{ env.NODE_VERSION }}
      - run: npm ci
      - run: npx playwright install
      - run: npm run test:e2e

  performance-tests:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with:
          node-version: ${{ env.NODE_VERSION }}
      - run: npm ci
      - run: npm run test:perf
      - uses: benchmark-action/github-action-benchmark@v1
        with:
          tool: 'customSmallerIsBetter'
          output-file-path: perf-results.json
          fail-on-alert: true

  compatibility-tests:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with:
          node-version: ${{ env.NODE_VERSION }}
      - run: npm ci
      - run: npm run test:excel-compat
```

### Test Coverage Requirements

| Category | Target | Measurement |
|----------|--------|-------------|
| Formula functions | 100% | Every function tested |
| Parser grammar | 100% | Every production rule |
| File formats | 95% | Every xlsx component |
| UI components | 80% | Statement coverage |
| Overall | 85% | Line coverage |
