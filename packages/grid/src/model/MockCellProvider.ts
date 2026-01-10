import type { CellData, CellProvider, CellRange, CellStyle } from "./CellProvider";

export class MockCellProvider implements CellProvider {
  private readonly rowCount: number;
  private readonly colCount: number;

  constructor(options: { rowCount: number; colCount: number }) {
    this.rowCount = options.rowCount;
    this.colCount = options.colCount;
  }

  prefetch(_range: CellRange): void {
    // Mock provider is synchronous; real implementations can fill an internal cache here.
  }

  getCell(row: number, col: number): CellData | null {
    if (row < 0 || col < 0 || row >= this.rowCount || col >= this.colCount) return null;

    const style: CellStyle | undefined =
      row === 0
        ? { fill: "#f5f5f5", fontWeight: "600" }
        : row % 2 === 0
          ? { fill: "#ffffff" }
          : { fill: "#fcfcfc" };

    const value =
      row === 0
        ? `Column ${col + 1}`
        : col === 0
          ? row
          : `${row},${col}`;

    return { row, col, value, style };
  }
}

