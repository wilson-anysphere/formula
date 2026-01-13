import type { Anchor, AnchorPoint, CellOffset, DrawingObject, DrawingObjectKind, EmuSize, ImageEntry, ImageStore } from "./types";

type JsonRecord = Record<string, unknown>;

function isRecord(value: unknown): value is JsonRecord {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function pick(obj: JsonRecord, keys: string[]): unknown {
  for (const key of keys) {
    if (Object.prototype.hasOwnProperty.call(obj, key)) return obj[key];
  }
  return undefined;
}

function unwrapExternallyTaggedEnum(value: unknown, context: string): { tag: string; value: unknown } {
  if (!isRecord(value)) {
    throw new Error(`${context} must be an externally-tagged enum object`);
  }
  const keys = Object.keys(value);
  if (keys.length !== 1) {
    throw new Error(`${context} must have exactly one variant key; got ${keys.length}`);
  }
  const tag = keys[0]!;
  return { tag, value: (value as Record<string, unknown>)[tag] };
}

function readNumber(value: unknown, context: string): number {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value === "bigint") return Number(value);
  if (typeof value === "string" && value.trim().length > 0) {
    const n = Number(value);
    if (Number.isFinite(n)) return n;
  }
  throw new Error(`${context} must be a number`);
}

function readOptionalNumber(value: unknown): number | undefined {
  if (value == null) return undefined;
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value === "bigint") return Number(value);
  if (typeof value === "string" && value.trim().length > 0) {
    const n = Number(value);
    return Number.isFinite(n) ? n : undefined;
  }
  return undefined;
}

function readString(value: unknown, context: string): string {
  if (typeof value === "string") return value;
  throw new Error(`${context} must be a string`);
}

function readOptionalString(value: unknown): string | undefined {
  return typeof value === "string" ? value : undefined;
}

function parseIdNumber(value: unknown): number | undefined {
  const direct = readOptionalNumber(value);
  if (direct != null) return direct;

  if (Array.isArray(value) && value.length === 1) {
    return parseIdNumber(value[0]);
  }

  if (isRecord(value) && Object.prototype.hasOwnProperty.call(value, "0")) {
    return parseIdNumber((value as JsonRecord)["0"]);
  }

  return undefined;
}

function stableHash32(input: string): number {
  // FNV-1a 32-bit.
  let hash = 0x811c9dc5;
  for (let i = 0; i < input.length; i++) {
    hash ^= input.charCodeAt(i);
    // eslint-disable-next-line no-bitwise
    hash = Math.imul(hash, 0x01000193);
  }
  // eslint-disable-next-line no-bitwise
  return hash >>> 0;
}

function stableStringify(value: unknown): string {
  try {
    return JSON.stringify(value, (_key, v) => {
      if (!isRecord(v)) return v;
      const out: Record<string, unknown> = {};
      for (const key of Object.keys(v).sort()) {
        out[key] = v[key];
      }
      return out;
    });
  } catch {
    return String(value);
  }
}

function parseDrawingObjectId(value: unknown): number {
  const parsed = parseIdNumber(value);
  if (parsed != null) return parsed;
  return stableHash32(stableStringify(value));
}

function parseImageId(value: unknown, context: string): string {
  if (typeof value === "string") return value;

  if (typeof value === "number" || typeof value === "bigint") return String(value);

  if (Array.isArray(value) && value.length === 1) {
    return parseImageId(value[0], context);
  }

  if (isRecord(value) && Object.prototype.hasOwnProperty.call(value, "0")) {
    return parseImageId((value as JsonRecord)["0"], context);
  }

  throw new Error(`${context} must be a string`);
}

function convertModelEmuSize(model: unknown, context: string): EmuSize {
  if (!isRecord(model)) throw new Error(`${context} must be an object`);
  return {
    cx: readNumber((model as JsonRecord).cx, `${context}.cx`),
    cy: readNumber((model as JsonRecord).cy, `${context}.cy`),
  };
}

function convertModelCellOffset(model: unknown, context: string): CellOffset {
  if (!isRecord(model)) throw new Error(`${context} must be an object`);
  const x = pick(model, ["x_emu", "xEmu"]);
  const y = pick(model, ["y_emu", "yEmu"]);
  return {
    xEmu: readNumber(x, `${context}.x_emu`),
    yEmu: readNumber(y, `${context}.y_emu`),
  };
}

function convertModelAnchorPoint(model: unknown, context: string): AnchorPoint {
  if (!isRecord(model)) throw new Error(`${context} must be an object`);

  const cell = (model as JsonRecord).cell;
  if (!isRecord(cell)) throw new Error(`${context}.cell must be an object`);

  return {
    cell: {
      row: readNumber((cell as JsonRecord).row, `${context}.cell.row`),
      col: readNumber((cell as JsonRecord).col, `${context}.cell.col`),
    },
    offset: convertModelCellOffset((model as JsonRecord).offset, `${context}.offset`),
  };
}

export function convertModelAnchorToUiAnchor(modelAnchorJson: unknown): Anchor {
  const { tag, value } = unwrapExternallyTaggedEnum(modelAnchorJson, "Anchor");
  if (!isRecord(value)) throw new Error(`Anchor.${tag} must be an object`);

  switch (tag) {
    case "OneCell": {
      const from = convertModelAnchorPoint((value as JsonRecord).from, "Anchor.OneCell.from");
      const ext = pick(value, ["ext", "size"]);
      const size = convertModelEmuSize(ext, "Anchor.OneCell.ext");
      return { type: "oneCell", from, size };
    }
    case "TwoCell": {
      const from = convertModelAnchorPoint((value as JsonRecord).from, "Anchor.TwoCell.from");
      const to = convertModelAnchorPoint((value as JsonRecord).to, "Anchor.TwoCell.to");
      return { type: "twoCell", from, to };
    }
    case "Absolute": {
      const pos = convertModelCellOffset((value as JsonRecord).pos, "Anchor.Absolute.pos");
      const ext = pick(value, ["ext", "size"]);
      const size = convertModelEmuSize(ext, "Anchor.Absolute.ext");
      return { type: "absolute", pos, size };
    }
    default:
      throw new Error(`Unsupported Anchor variant: ${tag}`);
  }
}

function convertModelDrawingObjectKind(model: unknown): DrawingObjectKind {
  const { tag, value } = unwrapExternallyTaggedEnum(model, "DrawingObjectKind");
  if (!isRecord(value)) throw new Error(`DrawingObjectKind.${tag} must be an object`);

  switch (tag) {
    case "Image": {
      const imageIdValue = pick(value, ["image_id", "imageId"]);
      const imageId = parseImageId(imageIdValue, "DrawingObjectKind.Image.image_id");
      return { type: "image", imageId };
    }
    case "Shape": {
      // Preserve a minimal label for debugging/selection UIs. The overlay currently
      // only cares about `kind.type`.
      return { type: "shape", label: "shape" };
    }
    case "ChartPlaceholder": {
      const relIdValue = pick(value, ["rel_id", "relId"]);
      const relId = readString(relIdValue, "DrawingObjectKind.ChartPlaceholder.rel_id");
      return { type: "chart", chartId: relId, label: `Chart (${relId})` };
    }
    case "Unknown":
      return { type: "unknown", label: "unknown" };
    default:
      return { type: "unknown", label: `unsupported:${tag}` };
  }
}

export function convertModelDrawingObjectToUiDrawingObject(modelObjJson: unknown): DrawingObject {
  if (!isRecord(modelObjJson)) throw new Error("DrawingObject must be an object");

  const id = parseDrawingObjectId((modelObjJson as JsonRecord).id);
  const anchor = convertModelAnchorToUiAnchor((modelObjJson as JsonRecord).anchor);
  const kind = convertModelDrawingObjectKind((modelObjJson as JsonRecord).kind);

  const zOrderValue = pick(modelObjJson, ["z_order", "zOrder"]);
  const zOrder = zOrderValue == null ? 0 : readNumber(zOrderValue, "DrawingObject.z_order");

  const sizeValue = pick(modelObjJson, ["size"]);
  const size =
    sizeValue == null ? undefined : convertModelEmuSize(sizeValue, "DrawingObject.size");

  return { id, kind, anchor, zOrder, size };
}

class MapImageStore implements ImageStore {
  private readonly images = new Map<string, ImageEntry>();

  get(id: string): ImageEntry | undefined {
    return this.images.get(id);
  }

  set(entry: ImageEntry): void {
    this.images.set(entry.id, entry);
  }
}

function decodeBase64ToBytes(base64: string): Uint8Array | null {
  // Browser.
  if (typeof globalThis.atob === "function") {
    const binary = globalThis.atob(base64);
    const out = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i);
    return out;
  }

  // Node.
  try {
    // eslint-disable-next-line n/no-unsupported-features/node-builtins
    const buf = Buffer.from(base64, "base64");
    return new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
  } catch {
    return null;
  }
}

function parseBytes(value: unknown, context: string): Uint8Array {
  if (value instanceof Uint8Array) return value;

  if (Array.isArray(value)) {
    const out = new Uint8Array(value.length);
    for (let i = 0; i < value.length; i++) {
      const n = readNumber(value[i], `${context}[${i}]`);
      out[i] = n & 0xff;
    }
    return out;
  }

  if (typeof value === "string") {
    const decoded = decodeBase64ToBytes(value);
    if (decoded) return decoded;
  }

  throw new Error(`${context} must be a byte array`);
}

function inferMimeTypeFromId(id: string): string {
  const ext = id.split(".").pop()?.toLowerCase();
  switch (ext) {
    case "png":
      return "image/png";
    case "jpg":
    case "jpeg":
      return "image/jpeg";
    case "gif":
      return "image/gif";
    case "bmp":
      return "image/bmp";
    case "webp":
      return "image/webp";
    case "svg":
      return "image/svg+xml";
    default:
      return "application/octet-stream";
  }
}

export function convertModelImageStoreToUiImageStore(modelImagesJson: unknown): ImageStore {
  const store = new MapImageStore();
  if (modelImagesJson == null) return store;
  if (!isRecord(modelImagesJson)) return store;

  const imagesValue = pick(modelImagesJson, ["images"]);
  const images = isRecord(imagesValue) ? imagesValue : modelImagesJson;
  if (!isRecord(images)) return store;

  for (const [imageId, data] of Object.entries(images)) {
    if (!isRecord(data)) continue;
    const bytesValue = pick(data, ["bytes"]);
    if (bytesValue == null) continue;

    const contentType = readOptionalString(pick(data, ["content_type", "contentType"]));
    const mimeType = contentType && contentType.length > 0 ? contentType : inferMimeTypeFromId(imageId);
    const bytes = parseBytes(bytesValue, `ImageStore.images[${imageId}].bytes`);

    store.set({ id: imageId, bytes, mimeType });
  }

  return store;
}

