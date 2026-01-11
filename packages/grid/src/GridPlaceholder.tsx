import type { CSSProperties } from "react";

export type GridPlaceholderProps = {
  columns?: number;
  rows?: number;
  cellSizePx?: number;
};

const containerStyle: CSSProperties = {
  border: "1px solid var(--formula-grid-line, #e5e7eb)",
  borderRadius: 8,
  overflow: "hidden",
  background: "var(--formula-grid-bg, #fff)",
};

export function GridPlaceholder({
  columns = 12,
  rows = 24,
  cellSizePx = 32,
}: GridPlaceholderProps) {
  const gridStyle: CSSProperties = {
    display: "grid",
    gridTemplateColumns: `repeat(${columns}, ${cellSizePx}px)`,
    gridTemplateRows: `repeat(${rows}, ${cellSizePx}px)`,
  };

  return (
    <div style={containerStyle} aria-label="Grid placeholder">
      <div style={gridStyle}>
        {Array.from({ length: columns * rows }).map((_, index) => (
          <div
            // eslint-disable-next-line react/no-array-index-key
            key={index}
            style={{
              borderRight: "1px solid var(--formula-grid-line, #f1f5f9)",
              borderBottom: "1px solid var(--formula-grid-line, #f1f5f9)",
            }}
          />
        ))}
      </div>
    </div>
  );
}
