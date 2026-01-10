export interface GridPresenceCursor {
  row: number;
  col: number;
}

/** Inclusive row/col endpoints. */
export interface GridPresenceRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

export interface GridPresence {
  id: string;
  name: string;
  color: string;
  cursor: GridPresenceCursor | null;
  selections: GridPresenceRange[];
}

