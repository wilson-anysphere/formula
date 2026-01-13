export type Emu = number;

export interface EmuSize {
  cx: Emu;
  cy: Emu;
}

export interface CellRef {
  /** Zero-based row index (matches DrawingML anchors). */
  row: number;
  /** Zero-based column index (matches DrawingML anchors). */
  col: number;
}

export interface CellOffset {
  xEmu: Emu;
  yEmu: Emu;
}

export interface AnchorPoint {
  cell: CellRef;
  offset: CellOffset;
}

export type Anchor =
  | { type: "oneCell"; from: AnchorPoint; size: EmuSize }
  | { type: "twoCell"; from: AnchorPoint; to: AnchorPoint }
  | { type: "absolute"; pos: CellOffset; size: EmuSize };

export type DrawingObjectKind =
  | { type: "image"; imageId: string }
  | { type: "shape"; label?: string; rawXml?: string }
  | { type: "chart"; chartId?: string; label?: string; rawXml?: string }
  | { type: "unknown"; label?: string; rawXml?: string };

export interface DrawingObject {
  id: number;
  kind: DrawingObjectKind;
  anchor: Anchor;
  /** Lower means behind. */
  zOrder: number;
  /** Optional extracted size; used for preview/handles. */
  size?: EmuSize;
}

export interface ImageEntry {
  id: string;
  bytes: Uint8Array;
  mimeType: string;
}

export interface ImageStore {
  get(id: string): ImageEntry | undefined;
  set(entry: ImageEntry): void;
}

export interface Rect {
  x: number;
  y: number;
  width: number;
  height: number;
}
