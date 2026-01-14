export type Emu = number;

export type DrawingObjectId = number;

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
  id: DrawingObjectId;
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
  /**
   * Optional async helpers for stores that can load/persist out-of-process
   * (e.g. IndexedDB). Callers should treat these as best-effort.
   */
  getAsync?(id: string): Promise<ImageEntry | undefined>;
  setAsync?(entry: ImageEntry): Promise<void>;
}

export interface Rect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export function createDrawingObjectId(): number {
  // We keep `DrawingObject.id` as a number for minimal disruption, but IDs must be globally unique
  // across collaborators. Generate a random 53-bit integer so it is safe to represent as a JS number.
  //
  // Collision probability: with a uniform 53-bit space, the birthday bound gives ~n^2 / 2^54.
  // Even at 10k inserted objects this is ~5e-9.
  const cryptoObj = globalThis.crypto;
  if (cryptoObj && typeof cryptoObj.getRandomValues === "function") {
    // 53 bits = 21 high bits + 32 low bits.
    const parts = new Uint32Array(2);
    cryptoObj.getRandomValues(parts);
    const high21 = parts[0] & 0x1fffff;
    const id = high21 * 2 ** 32 + parts[1];
    // Avoid `0` (useful as a sentinel in some code paths).
    if (id !== 0) return id;
  }

  // Fallback for environments without WebCrypto. `Math.random()` is not cryptographically secure,
  // but still provides sufficient entropy to make collisions extremely unlikely at our scale.
  const fallback = Math.floor(Math.random() * Number.MAX_SAFE_INTEGER);
  return fallback === 0 ? 1 : fallback;
}
