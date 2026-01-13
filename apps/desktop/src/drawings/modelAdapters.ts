import type { Anchor, AnchorPoint, CellOffset, DrawingObject, DrawingObjectKind, EmuSize, ImageEntry, ImageStore } from "./types";
import { graphicFramePlaceholderLabel } from "./shapeRenderer";
import { parseDrawingTransformFromRawXml } from "./transform";

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

function unwrapPossiblyTaggedEnum(
  input: unknown,
  context: string,
  opts?: { tagKeys?: string[]; contentKeys?: string[] },
): { tag: string; value: unknown } {
  const tagKeys = opts?.tagKeys ?? ["kind", "type"];
  const contentKeys = opts?.contentKeys ?? ["value", "content"];

  if (!isRecord(input)) {
    throw new Error(`${context} must be an enum object`);
  }

  // Prefer externally-tagged enums: `{ Variant: {...} }`.
  const keys = Object.keys(input);
  if (keys.length === 1 && !tagKeys.includes(keys[0]!)) {
    const tag = keys[0]!;
    return { tag, value: (input as Record<string, unknown>)[tag] };
  }

  // Support `#[serde(tag = "...", content = "...")]`: `{ type: "Variant", value: {...} }`.
  for (const tagKey of tagKeys) {
    const tagVal = (input as JsonRecord)[tagKey];
    if (typeof tagVal !== "string") continue;

    for (const contentKey of contentKeys) {
      const content = (input as JsonRecord)[contentKey];
      if (isRecord(content)) return { tag: tagVal, value: content };
    }

    // Support `#[serde(tag = "...")]`: `{ kind: "Variant", ...fields }`.
    return { tag: tagVal, value: input };
  }

  throw new Error(`${context} must be an externally-tagged or internally-tagged enum object`);
}

function normalizeEnumTag(tag: string): string {
  // Normalise variants like `OneCell`, `oneCell`, `one_cell` to a stable key.
  return tag.replace(/[^A-Za-z0-9]/g, "").toLowerCase();
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

function decodeXmlEntities(value: string): string {
  // Minimal decode for the most common entities present in DrawingML object names.
  return value
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&amp;", "&");
}

function extractDrawingObjectName(rawXml: string | undefined): string | undefined {
  if (!rawXml) return undefined;

  // `cNvPr` is the non-visual properties element used for names/IDs.
  // Example: `<xdr:cNvPr id="2" name="SmartArt 1"/>`
  const match =
    /<\s*(?:[A-Za-z0-9_-]+:)?cNvPr\b[^>]*\bname\s*=\s*"([^"]+)"/.exec(rawXml) ??
    /<\s*(?:[A-Za-z0-9_-]+:)?cNvPr\b[^>]*\bname\s*=\s*'([^']+)'/.exec(rawXml);
  const name = match?.[1]?.trim();
  if (!name) return undefined;
  return decodeXmlEntities(name);
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

function parseSheetId(value: unknown): string | undefined {
  if (value == null) return undefined;
  if (typeof value === "string") {
    const trimmed = value.trim();
    return trimmed ? trimmed : undefined;
  }
  if (typeof value === "number" && Number.isFinite(value)) return String(value);
  if (typeof value === "bigint") return String(value);
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
  const { tag, value } = unwrapPossiblyTaggedEnum(modelAnchorJson, "Anchor", { tagKeys: ["kind", "type"] });
  if (!isRecord(value)) throw new Error(`Anchor.${tag} must be an object`);
  const normalized = normalizeEnumTag(tag);

  switch (normalized) {
    case "onecell": {
      const from = convertModelAnchorPoint((value as JsonRecord).from, "Anchor.OneCell.from");
      const ext = pick(value, ["ext", "size"]);
      const size = convertModelEmuSize(ext, "Anchor.OneCell.ext");
      return { type: "oneCell", from, size };
    }
    case "twocell": {
      const from = convertModelAnchorPoint((value as JsonRecord).from, "Anchor.TwoCell.from");
      const to = convertModelAnchorPoint((value as JsonRecord).to, "Anchor.TwoCell.to");
      return { type: "twoCell", from, to };
    }
    case "absolute": {
      const pos = convertModelCellOffset((value as JsonRecord).pos, "Anchor.Absolute.pos");
      const ext = pick(value, ["ext", "size"]);
      const size = convertModelEmuSize(ext, "Anchor.Absolute.ext");
      return { type: "absolute", pos, size };
    }
    default:
      throw new Error(`Unsupported Anchor variant: ${tag}`);
  }
}

function convertModelDrawingObjectKind(
  model: unknown,
  context?: { sheetId?: string; drawingObjectId?: number },
): DrawingObjectKind {
  const { tag, value } = unwrapPossiblyTaggedEnum(model, "DrawingObjectKind", { tagKeys: ["kind", "type"] });
  if (!isRecord(value)) throw new Error(`DrawingObjectKind.${tag} must be an object`);
  const normalized = normalizeEnumTag(tag);

  switch (normalized) {
    case "image": {
      const imageIdValue = pick(value, ["image_id", "imageId"]);
      const imageId = parseImageId(imageIdValue, "DrawingObjectKind.Image.image_id");
      return { type: "image", imageId };
    }
    case "shape": {
      const rawXmlValue = pick(value, ["raw_xml", "rawXml"]);
      const rawXml = readOptionalString(rawXmlValue);
      const label = extractDrawingObjectName(rawXml);
      return { type: "shape", rawXml, raw_xml: rawXml, label };
    }
    case "chartplaceholder": {
      const relIdValue = pick(value, ["rel_id", "relId", "chart_id", "chartId"]);
      const relId = readString(relIdValue, "DrawingObjectKind.ChartPlaceholder.rel_id");
      const rawXmlValue = pick(value, ["raw_xml", "rawXml"]);
      const rawXml = readOptionalString(rawXmlValue);
      const label = extractDrawingObjectName(rawXml);

      // The XLSX parser currently maps all `xdr:graphicFrame` objects to the
      // ChartPlaceholder variant, but many `graphicFrame` payloads are not charts
      // (e.g. SmartArt/diagram frames). These have `rel_id = "unknown"`.
      //
      // When the rel id is unknown, treat the object as `unknown` so overlay
      // rendering can use `graphicFramePlaceholderLabel(...)` for a stable label.
      if (relId.trim() === "" || relId === "unknown") {
        return {
          type: "unknown",
          rawXml,
          raw_xml: rawXml,
          label: label ?? graphicFramePlaceholderLabel(rawXml) ?? undefined,
        };
      }

      const sheetId = context?.sheetId;
      const objectId = context?.drawingObjectId;
      const chartId =
        sheetId && objectId != null
          ? `${sheetId}:${String(objectId)}`
          : // Back-compat: when the sheet context isn't available, fall back to the drawing rel id.
            relId;

      return { type: "chart", chartId, label: label ?? `Chart (${relId})`, rawXml, raw_xml: rawXml };
    }
    case "chart": {
      // UI/other internal representations may already use `{ type: "chart", chartId }`.
      const chartId = readOptionalString(pick(value, ["chart_id", "chartId", "rel_id", "relId"]));
      const rawXml = readOptionalString(pick(value, ["raw_xml", "rawXml"]));
      const label = readOptionalString(pick(value, ["label"])) ?? extractDrawingObjectName(rawXml);
      return { type: "chart", chartId: chartId ?? undefined, rawXml, raw_xml: rawXml, label: label ?? undefined };
    }
    case "unknown":
      {
        const rawXml = readOptionalString(pick(value, ["raw_xml", "rawXml"]));
        return {
          type: "unknown",
          rawXml,
          raw_xml: rawXml,
          label: extractDrawingObjectName(rawXml),
        };
      }
    default:
      return { type: "unknown", label: `unsupported:${tag}` };
  }
}

export function convertModelDrawingObjectToUiDrawingObject(
  modelObjJson: unknown,
  context?: { sheetId?: string | number },
): DrawingObject {
  if (!isRecord(modelObjJson)) throw new Error("DrawingObject must be an object");

  const id = parseDrawingObjectId((modelObjJson as JsonRecord).id);
  const anchor = convertModelAnchorToUiAnchor((modelObjJson as JsonRecord).anchor);
  const sheetId = parseSheetId(context?.sheetId);
  const kind = convertModelDrawingObjectKind((modelObjJson as JsonRecord).kind, { sheetId, drawingObjectId: id });

  const zOrderValue = pick(modelObjJson, ["z_order", "zOrder"]);
  const zOrder = zOrderValue == null ? 0 : readNumber(zOrderValue, "DrawingObject.z_order");

  const sizeValue = pick(modelObjJson, ["size"]);
  const size =
    sizeValue == null ? undefined : convertModelEmuSize(sizeValue, "DrawingObject.size");

  const preservedValue = pick(modelObjJson, ["preserved"]);
  let preserved: Record<string, string> | undefined;
  if (isRecord(preservedValue)) {
    const out: Record<string, string> = {};
    for (const [k, v] of Object.entries(preservedValue)) {
      if (typeof v === "string") out[k] = v;
    }
    if (Object.keys(out).length > 0) preserved = out;
  }

  const transform = (() => {
    if (kind.type === "image") {
      const picXml = preserved?.["xlsx.pic_xml"];
      if (typeof picXml !== "string" || picXml.length === 0) return undefined;
      const parsed = parseDrawingTransformFromRawXml(picXml);
      if (!parsed) return undefined;
      return parsed.rotationDeg !== 0 || parsed.flipH || parsed.flipV ? parsed : undefined;
    }

    const rawXml = (kind as any).rawXml ?? (kind as any).raw_xml;
    if (typeof rawXml !== "string" || rawXml.length === 0) return undefined;
    const parsed = parseDrawingTransformFromRawXml(rawXml);
    if (!parsed) return undefined;
    return parsed.rotationDeg !== 0 || parsed.flipH || parsed.flipV ? parsed : undefined;
  })();

  const out: DrawingObject = { id, kind, anchor, zOrder, size };
  if (preserved) out.preserved = preserved;
  if (transform) out.transform = transform;
  return out;
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
    try {
      const binary = globalThis.atob(base64);
      const out = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i);
      return out;
    } catch {
      return null;
    }
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

  if (isRecord(value)) {
    // Node's `Buffer` JSON representation: { type: "Buffer", data: number[] }.
    if ((value as JsonRecord).type === "Buffer" && Array.isArray((value as JsonRecord).data)) {
      return parseBytes((value as JsonRecord).data, `${context}.data`);
    }

    // `Uint8Array`/typed-array stringified via JSON can show up as an object with
    // numeric keys (e.g. { "0": 1, "1": 2, ... }).
    const numericKeys = Object.keys(value).filter((k) => /^\d+$/.test(k));
    if (numericKeys.length > 0) {
      numericKeys.sort((a, b) => Number(a) - Number(b));
      const maxIndex = Number(numericKeys[numericKeys.length - 1]);
      const declaredLength = readOptionalNumber((value as JsonRecord).length);
      const length =
        declaredLength != null && declaredLength >= maxIndex + 1 ? declaredLength : maxIndex + 1;

      const out = new Uint8Array(length);
      for (const k of numericKeys) {
        const idx = Number(k);
        const n = readNumber((value as JsonRecord)[k], `${context}[${k}]`);
        out[idx] = n & 0xff;
      }
      return out;
    }
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
    try {
      const bytes = parseBytes(bytesValue, `ImageStore.images[${imageId}].bytes`);
      store.set({ id: imageId, bytes, mimeType });
    } catch {
      // Best-effort: ignore malformed image payloads rather than aborting conversion
      // for the entire workbook.
      continue;
    }
  }

  return store;
}

/**
 * Convert a formula-model `Worksheet` JSON blob into the UI overlay model.
 *
 * This is a best-effort adapter: invalid drawing objects are ignored rather than
 * throwing, so a single malformed entry does not prevent the sheet from
 * rendering.
 */
export function convertModelWorksheetDrawingsToUiDrawingObjects(modelWorksheetJson: unknown): DrawingObject[] {
  if (!isRecord(modelWorksheetJson)) return [];
  const sheetId = parseSheetId(pick(modelWorksheetJson, ["id", "sheet_id", "sheetId"]));
  const drawingsValue = pick(modelWorksheetJson, ["drawings"]);
  if (!Array.isArray(drawingsValue)) return [];

  const out: DrawingObject[] = [];
  for (const obj of drawingsValue) {
    try {
      out.push(convertModelDrawingObjectToUiDrawingObject(obj, sheetId ? { sheetId } : undefined));
    } catch {
      // ignore
    }
  }
  return out;
}

/**
 * Convert a formula-model `Workbook` JSON snapshot into a structure that's easy to feed into
 * the canvas overlay (image store + per-sheet drawing objects).
 */
export function convertModelWorkbookDrawingsToUiDrawingLayer(modelWorkbookJson: unknown): {
  images: ImageStore;
  drawingsBySheetName: Record<string, DrawingObject[]>;
} {
  if (!isRecord(modelWorkbookJson)) {
    return { images: new MapImageStore(), drawingsBySheetName: {} };
  }

  const images = convertModelImageStoreToUiImageStore(pick(modelWorkbookJson, ["images"]));
  const drawingsBySheetName: Record<string, DrawingObject[]> = {};

  const sheetsValue = pick(modelWorkbookJson, ["sheets"]);
  if (Array.isArray(sheetsValue)) {
    for (const sheet of sheetsValue) {
      if (!isRecord(sheet)) continue;
      const name = readOptionalString(pick(sheet, ["name"])) ?? "";
      if (!name) continue;
      drawingsBySheetName[name] = convertModelWorksheetDrawingsToUiDrawingObjects(sheet);
    }
  } else if (isRecord(sheetsValue)) {
    // Some workbook JSON snapshots represent sheets as a keyed object rather than
    // an array (e.g. `{ sheets: { Sheet1: {...}, Sheet2: {...} } }`).
    for (const [key, sheet] of Object.entries(sheetsValue)) {
      if (!isRecord(sheet)) continue;
      const name = readOptionalString(pick(sheet, ["name"])) ?? key;
      if (!name) continue;
      drawingsBySheetName[name] = convertModelWorksheetDrawingsToUiDrawingObjects(sheet);
    }
  }

  return { images, drawingsBySheetName };
}
