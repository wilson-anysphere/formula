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
  | { type: "shape"; label?: string; rawXml?: string; raw_xml?: string }
  | { type: "chart"; chartId?: string; label?: string; relId?: string; rawXml?: string; raw_xml?: string }
  | { type: "unknown"; label?: string; rawXml?: string; raw_xml?: string };

export interface DrawingTransform {
  /** Clockwise rotation in degrees (DrawingML uses 60000ths of a degree). */
  rotationDeg: number;
  flipH: boolean;
  flipV: boolean;
}

export interface DrawingObject {
  id: number;
  kind: DrawingObjectKind;
  anchor: Anchor;
  /** Lower means behind. */
  zOrder: number;
  /** Optional extracted size; used for preview/handles. */
  size?: EmuSize;
  /**
   * Optional preserved metadata passed through for format compatibility layers.
   *
   * For XLSX drawings this may include:
   * - `xlsx.pic_xml` (the `<xdr:pic>…</xdr:pic>` subtree for images)
   * - `xlsx.embed_rel_id`
   */
  preserved?: Record<string, string>;
  /** Optional DrawingML transform metadata (rotation / flips). */
  transform?: DrawingTransform;
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
