import type { CellData, CellProvider, CellRange, CellStyle } from "./CellProvider";

export class MockCellProvider implements CellProvider {
  private readonly rowCount: number;
  private readonly colCount: number;

  private readonly headerStyle: CellStyle = { fontWeight: "600", textAlign: "center" };
  private readonly rowHeaderStyle: CellStyle = { fontWeight: "600", textAlign: "end" };

  constructor(options: { rowCount: number; colCount: number }) {
    this.rowCount = options.rowCount;
    this.colCount = options.colCount;
  }

  prefetch(_range: CellRange): void {
    // Mock provider is synchronous; real implementations can fill an internal cache here.
  }

  getCell(row: number, col: number): CellData | null {
    if (row < 0 || col < 0 || row >= this.rowCount || col >= this.colCount) return null;

    let style: CellStyle | undefined;
    if (row === 0 && col === 0) {
      style = this.headerStyle;
    } else if (row === 0) {
      style = this.headerStyle;
    } else if (col === 0) {
      style = this.rowHeaderStyle;
    }

    if (row === 1 && col === 1) {
      style = { ...(style ?? {}), wrapMode: "word", textAlign: "start" };
    } else if (row === 2 && col === 1) {
      style = { ...(style ?? {}), wrapMode: "word", textAlign: "start" };
    } else if (row === 3 && col === 1) {
      style = { ...(style ?? {}), wrapMode: "word", textAlign: "start" };
    }

    const value =
      row === 0
        ? `Column ${col + 1}`
        : col === 0
          ? row
          : row === 4 && col === 1
            ? true
            : row === 5 && col === 1
              ? false
              : row === 6 && col === 1
                ? "#DIV/0!"
          : row === 1 && col === 1
            ? "This is a long piece of text that should wrap in the cell (word wrap)."
            : row === 2 && col === 1
              ? "שלום world 123 — mixed RTL/LTR"
              : row === 3 && col === 1
                ? "مرحبا بالعالم hello — Arabic + English"
                : `${row},${col}`;

    const comment =
      row === 1 && col === 1
        ? { resolved: false }
        : row === 2 && col === 1
          ? { resolved: true }
          : null;

    return { row, col, value, style, comment };
  }
}
