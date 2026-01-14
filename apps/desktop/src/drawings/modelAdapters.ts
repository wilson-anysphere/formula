import type {
  Anchor,
  AnchorPoint,
  CellOffset,
  DrawingObject,
  DrawingObjectKind,
  DrawingTransform,
  EmuSize,
  ImageEntry,
  ImageStore,
} from "./types";
import { graphicFramePlaceholderLabel } from "./shapeRenderer";
import { parseDrawingTransformFromRawXml } from "./transform";
import { pxToEmu } from "../shared/emu.js";

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

  // Support externally-tagged enums with metadata keys alongside the variant payload, e.g.:
  // `{ sheetId: "...", Absolute: {...} }` or `{ label: "Foo", Shape: {...} }`.
  //
  // Heuristic: pick the single non-tag/non-content key whose value is an object.
  const recordKeys = keys.filter(
    (key) => !tagKeys.includes(key) && !contentKeys.includes(key) && isRecord((input as JsonRecord)[key]),
  );
  if (recordKeys.length === 1) {
    const tag = recordKeys[0]!;
    return { tag, value: (input as JsonRecord)[tag] };
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

function unwrapSingletonId(value: unknown): unknown {
  if (Array.isArray(value) && value.length === 1) {
    return unwrapSingletonId(value[0]);
  }
  if (isRecord(value) && Object.prototype.hasOwnProperty.call(value, "0")) {
    return unwrapSingletonId((value as JsonRecord)["0"]);
  }
  return value;
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

function stableHash53(input: string): number {
  // 53-bit safe integer hash derived from two independent 32-bit FNV-1a hashes.
  //
  // We intentionally avoid BigInt here so the hash remains cheap and works in older JS runtimes.
  // The collision probability is close to a uniform 53-bit space, which is dramatically safer than
  // a 32-bit hash when mapping many non-numeric drawing ids into numeric keys.
  const low32 = stableHash32(input);
  const high32 = stableHash32(`h:${input}`);
  // Use 21 high bits + 32 low bits => 53 bits.
  // eslint-disable-next-line no-bitwise
  const high21 = high32 & 0x1fffff;
  return high21 * 2 ** 32 + low32;
}

function stableStringify(value: unknown): string {
  try {
    const serialized = JSON.stringify(value, (_key, v) => {
      if (!isRecord(v)) return v;
      const out: Record<string, unknown> = {};
      for (const key of Object.keys(v).sort()) {
        out[key] = v[key];
      }
      return out;
    });
    // `JSON.stringify` returns `undefined` for inputs like `undefined` / functions / symbols.
    // Guard so callers like `stableHash32(stableStringify(...))` never throw.
    return typeof serialized === "string" ? serialized : String(value);
  } catch {
    return String(value);
  }
}

function normalizeDrawingIdForHash(unwrapped: unknown): unknown {
  if (typeof unwrapped === "string") {
    // DocumentController trims string ids before persisting; mirror that normalization here so
    // UI-layer hashed ids stay stable even if upstream snapshots include accidental whitespace.
    //
    // Defensive guard: avoid hashing arbitrarily-large strings (potential collab DoS vector).
    // For long ids we hash a stable *summary* (trimmed length + prefix/middle/suffix) instead of
    // the full string to keep hashing bounded.
    const MAX_LEN = 4096;
    if (unwrapped.length <= MAX_LEN) {
      return unwrapped.trim();
    }

    const SAMPLE = 1024;
    const TRIM_SCAN = 2048;

    // Best-effort trim without scanning the entire string: only inspect the first/last `TRIM_SCAN`
    // characters to compute trim offsets and samples. This keeps costs bounded even for huge ids.
    const startChunk = unwrapped.slice(0, TRIM_SCAN);
    const startMatch = /^\s+/.exec(startChunk);
    const start = startMatch ? startMatch[0].length : 0;

    const endChunk = unwrapped.slice(Math.max(0, unwrapped.length - TRIM_SCAN));
    const endMatch = /\s+$/.exec(endChunk);
    const endTrim = endMatch ? endMatch[0].length : 0;
    const end = Math.max(start, unwrapped.length - endTrim);

    const trimmedLen = Math.max(0, end - start);

    const prefix = unwrapped.slice(start, Math.min(end, start + SAMPLE));
    const suffix = unwrapped.slice(Math.max(start, end - SAMPLE), end);
    const midStart = start + Math.max(0, Math.floor(trimmedLen / 2) - Math.floor(SAMPLE / 2));
    const mid = unwrapped.slice(midStart, Math.min(end, midStart + SAMPLE));

    return { kind: "drawingId:longString", trimmedLen, prefix, mid, suffix };
  }
  return unwrapped;
}

function parseDrawingObjectId(value: unknown): number {
  const unwrapped = unwrapSingletonId(value);
  const parsed = (() => {
    if (typeof unwrapped === "number") return readOptionalNumber(unwrapped);
    if (typeof unwrapped === "bigint") return readOptionalNumber(unwrapped);
    if (typeof unwrapped === "string") {
      // Avoid expensive parsing work for huge string ids. Canonical safe-integer ids are at most
      // 16 digits (and a small amount of surrounding whitespace), so longer strings should be
      // treated as opaque and hashed.
      if (unwrapped.length > 64) return undefined;
      const trimmed = unwrapped.trim();
      if (!trimmed) return undefined;
      // Only accept canonical base-10 integer strings so distinct raw ids like "001" and "1"
      // do not collide in the UI layer.
      if (trimmed.length > 16) return undefined;
      if (!/^\d+$/.test(trimmed)) return undefined;
      const n = Number(trimmed);
      if (!Number.isFinite(n)) return undefined;
      // Reject non-canonical numeric strings (leading zeros, +1, 1e3, etc).
      if (String(n) !== trimmed) return undefined;
      return n;
    }
    return parseIdNumber(unwrapped);
  })();
  // Drawing object ids must fit in JS's safe integer range since the overlay/hit-test layers treat
  // them as stable numeric keys. If an upstream snapshot stores ids as strings, guard against
  // parsing an unsafe integer (e.g. "9007199254740993") by falling back to a stable hash.
  //
  // Workbook drawing ids are expected to be positive; reserve negative ids for:
  // - ChartStore canvas charts (see `chartIdToDrawingId`)
  // - hashed ids produced below (large-magnitude negative namespace)
  if (parsed != null && Number.isSafeInteger(parsed) && parsed > 0) return parsed;
  // Use a disjoint negative-id namespace for hashed ids so they cannot collide with:
  // - normal drawing ids (which are positive safe integers, including our random 53-bit ids)
  // - chart overlay ids (which use smaller-magnitude negative ids; see `chartIdToDrawingId`)
  const HASH_NAMESPACE_OFFSET = 0x200000000; // 2^33
  const maxHash = Number.MAX_SAFE_INTEGER - HASH_NAMESPACE_OFFSET;
  const normalizedForHash = normalizeDrawingIdForHash(unwrapped);
  const hashed = stableHash53(stableStringify(normalizedForHash)) % maxHash;
  return -(HASH_NAMESPACE_OFFSET + hashed);
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
  const record = model as JsonRecord;

  const readFirst = (keys: string[], fieldContext: string): number => {
    for (const key of keys) {
      const candidate = readOptionalNumber(record[key]);
      if (candidate != null) return candidate;
    }
    throw new Error(`${fieldContext} must be a number`);
  };

  return {
    cx: readFirst(["cx", "cxEmu", "widthEmu", "width_emu", "wEmu"], `${context}.cx`),
    cy: readFirst(["cy", "cyEmu", "heightEmu", "height_emu", "hEmu"], `${context}.cy`),
  };
}

function convertModelCellOffset(model: unknown, context: string): CellOffset {
  if (!isRecord(model)) throw new Error(`${context} must be an object`);
  const record = model as JsonRecord;

  const readFirst = (keys: string[], fieldContext: string): number => {
    for (const key of keys) {
      const candidate = readOptionalNumber(record[key]);
      if (candidate != null) return candidate;
    }
    throw new Error(`${fieldContext} must be a number`);
  };

  return {
    xEmu: readFirst(["x_emu", "xEmu", "dxEmu", "offsetXEmu", "offset_x_emu"], `${context}.x_emu`),
    yEmu: readFirst(["y_emu", "yEmu", "dyEmu", "offsetYEmu", "offset_y_emu"], `${context}.y_emu`),
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
      const size =
        (() => {
          const extValue = pick(value, ["ext"]);
          const sizeValue = pick(value, ["size"]);
          try {
            return convertModelEmuSize(extValue, "Anchor.OneCell.ext");
          } catch {
            return convertModelEmuSize(sizeValue, "Anchor.OneCell.size");
          }
        })();
      return { type: "oneCell", from, size };
    }
    case "twocell": {
      const from = convertModelAnchorPoint((value as JsonRecord).from, "Anchor.TwoCell.from");
      const to = convertModelAnchorPoint((value as JsonRecord).to, "Anchor.TwoCell.to");
      return { type: "twoCell", from, to };
    }
    case "absolute": {
      const pos = convertModelCellOffset((value as JsonRecord).pos, "Anchor.Absolute.pos");
      const size =
        (() => {
          const extValue = pick(value, ["ext"]);
          const sizeValue = pick(value, ["size"]);
          try {
            return convertModelEmuSize(extValue, "Anchor.Absolute.ext");
          } catch {
            return convertModelEmuSize(sizeValue, "Anchor.Absolute.size");
          }
        })();
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
  const size = (() => {
    if (sizeValue == null) return undefined;
    try {
      return convertModelEmuSize(sizeValue, "DrawingObject.size");
    } catch {
      return undefined;
    }
  })();

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

  delete(id: string): void {
    this.images.delete(id);
  }

  clear(): void {
    this.images.clear();
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
  out.sort((a, b) => a.zOrder - b.zOrder);
  return out;
}

function convertDocumentDrawingSizeToEmu(sizeJson: unknown): EmuSize | undefined {
  if (!isRecord(sizeJson)) return undefined;
  const record = sizeJson as JsonRecord;

  const readFirstNumeric = (keys: string[]): number | undefined => {
    for (const key of keys) {
      const candidate = readOptionalNumber((record as any)[key]);
      if (candidate != null) return candidate;
    }
    return undefined;
  };

  // Explicit EMU payloads (future-proof).
  const cxEmu = readFirstNumeric(["cx", "cxEmu", "widthEmu", "width_emu", "wEmu"]);
  const cyEmu = readFirstNumeric(["cy", "cyEmu", "heightEmu", "height_emu", "hEmu"]);
  if (cxEmu != null && cyEmu != null) return { cx: cxEmu, cy: cyEmu };

  // Pixel payloads (DocumentController schema).
  const widthPx = readFirstNumeric(["width", "w", "widthPx", "width_px"]);
  const heightPx = readFirstNumeric(["height", "h", "heightPx", "height_px"]);
  if (widthPx != null && heightPx != null) {
    return { cx: pxToEmu(widthPx), cy: pxToEmu(heightPx) };
  }

  return undefined;
}

function convertDocumentDrawingAnchorToUiAnchor(anchorJson: unknown, size: EmuSize | undefined): Anchor | null {
  if (!isRecord(anchorJson)) return null;
  const outer = anchorJson as JsonRecord;
  // Support DocumentController anchors (simple `{ type: "absolute", ... }`) as well as
  // formula-model/Rust enum encodings (externally tagged `{ Absolute: {...} }` and
  // internally tagged `{ type: "Absolute", value: {...} }`).
  let tag = readOptionalString(pick(outer, ["type"])) ?? "";
  let payload: JsonRecord = outer;
  try {
    const unwrapped = unwrapPossiblyTaggedEnum(outer, "Drawing.anchor", { tagKeys: ["kind", "type"] });
    tag = unwrapped.tag;
    if (!isRecord(unwrapped.value)) return null;
    payload = unwrapped.value as JsonRecord;
  } catch {
    // ignore; fall back to treating `anchorJson` as a plain DocumentController anchor object.
  }
  // Some snapshots attach metadata keys (e.g. `sheetId`) alongside an externally-tagged enum:
  // `{ sheetId: "...", Absolute: { ... } }`. Detect that encoding so we don't drop drawings or
  // incorrectly fall back to model parsing.
  if (!tag) {
    const candidateKeys = Object.keys(outer).filter((key) => key !== "sheetId" && key !== "sheet_id");
    const variantKeys = candidateKeys.filter((key) => isRecord((outer as any)[key]));
    if (variantKeys.length === 1) {
      const variant = variantKeys[0]!;
      const value = (outer as any)[variant];
      if (isRecord(value)) {
        tag = variant;
        payload = value as JsonRecord;
      }
    }
  }

  let anchorType = normalizeEnumTag(tag);
  // Some persistence layers may include the DrawingML element name in the tag
  // (e.g. `OneCellAnchor`, `absolute_anchor`). Treat those as equivalent to the
  // core variants (`oneCell`, `twoCell`, `absolute`).
  if (anchorType.endsWith("anchor")) anchorType = anchorType.slice(0, -"anchor".length);

  const resolveOffsetEmuMaybeFrom = (record: JsonRecord, axis: "x" | "y"): number | undefined => {
    const emuKeys =
      axis === "x"
        ? ["xEmu", "x_emu", "dxEmu", "offsetXEmu", "offset_x_emu"]
        : ["yEmu", "y_emu", "dyEmu", "offsetYEmu", "offset_y_emu"];
    const pxKeys = axis === "x" ? ["x", "dx", "offsetX", "offsetXPx", "offset_x"] : ["y", "dy", "offsetY", "offsetYPx", "offset_y"];

    for (const key of emuKeys) {
      const candidate = readOptionalNumber((record as any)[key]);
      if (candidate != null) return candidate;
    }
    for (const key of pxKeys) {
      const candidate = readOptionalNumber((record as any)[key]);
      if (candidate != null) return pxToEmu(candidate);
    }

    return undefined;
  };
  const resolveOffsetEmu = (axis: "x" | "y"): number =>
    resolveOffsetEmuMaybeFrom(payload, axis) ?? (payload !== outer ? resolveOffsetEmuMaybeFrom(outer, axis) : undefined) ?? 0;

  const resolvedSize =
    size ??
    // Some snapshots (or mixed back-compat encodings) store the EMU size under `ext` (formula-model
    // field name) instead of `size`. Accept both so UI-like anchors can still round-trip.
    //
    // Important: try both keys rather than `pick(["size","ext"])` so a present-but-invalid `size`
    // payload does not mask a valid `ext`.
    convertDocumentDrawingSizeToEmu(pick(payload, ["size"])) ??
    convertDocumentDrawingSizeToEmu(pick(payload, ["ext"])) ??
    // Some older/alternate encodings store size fields directly on the anchor object itself
    // (e.g. `{ type: "absolute", xEmu, yEmu, cx, cy }`). Accept those as a last resort.
    convertDocumentDrawingSizeToEmu(payload) ??
    (payload !== outer ? convertDocumentDrawingSizeToEmu(outer) : undefined) ??
    { cx: pxToEmu(100), cy: pxToEmu(100) };

  switch (anchorType) {
    case "cell": {
      const row = readNumber(pick(payload, ["row"]), "Drawing.anchor.row");
      const col = readNumber(pick(payload, ["col"]), "Drawing.anchor.col");
      return {
        type: "oneCell",
        from: { cell: { row, col }, offset: { xEmu: resolveOffsetEmu("x"), yEmu: resolveOffsetEmu("y") } },
        size: resolvedSize,
      };
    }
    // Back-compat: accept UI-like anchors persisted in a DocumentController snapshot.
    case "onecell": {
      const fromValue = pick(payload, ["from"]);
      if (!isRecord(fromValue)) return null;
      const cellValue = pick(fromValue, ["cell"]);
      if (!isRecord(cellValue)) return null;
      const row = readNumber(pick(cellValue, ["row"]), "Drawing.anchor.from.cell.row");
      const col = readNumber(pick(cellValue, ["col"]), "Drawing.anchor.from.cell.col");
      const offsetValue = pick(fromValue, ["offset"]);
      const offset: CellOffset = {
        xEmu:
          (isRecord(offsetValue) ? resolveOffsetEmuMaybeFrom(offsetValue, "x") : undefined) ??
          resolveOffsetEmuMaybeFrom(fromValue, "x") ??
          resolveOffsetEmu("x"),
        yEmu:
          (isRecord(offsetValue) ? resolveOffsetEmuMaybeFrom(offsetValue, "y") : undefined) ??
          resolveOffsetEmuMaybeFrom(fromValue, "y") ??
          resolveOffsetEmu("y"),
      };
      return { type: "oneCell", from: { cell: { row, col }, offset }, size: resolvedSize };
    }
    case "absolute": {
      // Support both DocumentController-style anchors (which may store `xEmu/yEmu` on the root)
      // and UI-like anchors (which store `pos: { xEmu, yEmu }`).
      const posValue = pick(payload, ["pos"]);
      const pos = isRecord(posValue) ? posValue : null;
      const xEmu = (pos ? resolveOffsetEmuMaybeFrom(pos, "x") : undefined) ?? resolveOffsetEmu("x");
      const yEmu = (pos ? resolveOffsetEmuMaybeFrom(pos, "y") : undefined) ?? resolveOffsetEmu("y");
      return { type: "absolute", pos: { xEmu, yEmu }, size: resolvedSize };
    }
    case "twocell": {
      const fromValue = pick(payload, ["from"]);
      const toValue = pick(payload, ["to"]);
      if (!isRecord(fromValue) || !isRecord(toValue)) return null;

      const parsePoint = (point: JsonRecord, context: string): AnchorPoint | null => {
        const cellValue = pick(point, ["cell"]);
        if (!isRecord(cellValue)) return null;
        const row = readNumber(pick(cellValue, ["row"]), `${context}.cell.row`);
        const col = readNumber(pick(cellValue, ["col"]), `${context}.cell.col`);
        const offsetValue = pick(point, ["offset"]);
        const offset: CellOffset = {
          xEmu: (isRecord(offsetValue) ? resolveOffsetEmuMaybeFrom(offsetValue, "x") : undefined) ?? resolveOffsetEmuMaybeFrom(point, "x") ?? 0,
          yEmu: (isRecord(offsetValue) ? resolveOffsetEmuMaybeFrom(offsetValue, "y") : undefined) ?? resolveOffsetEmuMaybeFrom(point, "y") ?? 0,
        };
        return { cell: { row, col }, offset };
      };

      const from = parsePoint(fromValue, "Drawing.anchor.from");
      const to = parsePoint(toValue, "Drawing.anchor.to");
      if (!from || !to) return null;
      return { type: "twoCell", from, to };
    }
    default:
      return null;
  }
}

function convertDocumentDrawingKindToUiKind(kindJson: unknown): DrawingObjectKind | null {
  if (!isRecord(kindJson)) return null;
  // If the kind is stored as an internally-tagged enum (e.g. `{ type: "Shape", value: {...} }`),
  // prefer the formula-model adapter. DocumentController drawings should store kind metadata in a
  // flat shape; treating a tagged enum as a DocumentController kind can accidentally drop nested
  // payloads like `raw_xml`.
  const contentCandidate = pick(kindJson, ["value", "content"]);
  if (isRecord(contentCandidate)) return null;
  const type = normalizeEnumTag(readOptionalString(pick(kindJson, ["type"])) ?? "");
  const rawXml = readOptionalString(pick(kindJson, ["rawXml", "raw_xml"]));
  const label = readOptionalString(pick(kindJson, ["label"]));

  switch (type) {
    case "image": {
      const imageId = readOptionalString(pick(kindJson, ["imageId", "image_id"]));
      if (!imageId) return null;
      return { type: "image", imageId };
    }
    case "shape":
      return { type: "shape", ...(label ? { label } : {}), ...(rawXml ? { rawXml } : {}) };
    case "chart": {
      const chartId = readOptionalString(pick(kindJson, ["chartId", "chart_id", "relId", "rel_id"]));
      // Mirror `convertModelDrawingObjectKind` behavior: some graphicFrames are not charts
      // (e.g. SmartArt) and surface with `chartId/relId = "unknown"`. Treat those as unknown
      // so placeholder labels can use `graphicFramePlaceholderLabel(...)`.
      if (!chartId || chartId.trim() === "" || chartId === "unknown") {
        const derived = label ?? extractDrawingObjectName(rawXml) ?? graphicFramePlaceholderLabel(rawXml) ?? undefined;
        return { type: "unknown", ...(rawXml ? { rawXml } : {}), ...(derived ? { label: derived } : {}) };
      }

      return { type: "chart", chartId, ...(label ? { label } : {}), ...(rawXml ? { rawXml } : {}) };
    }
    case "unknown":
      return { type: "unknown", ...(label ? { label } : {}), ...(rawXml ? { rawXml } : {}) };
    case "chartplaceholder": {
      const chartId = readOptionalString(pick(kindJson, ["chartId", "chart_id", "relId", "rel_id"]));
      if (!chartId || chartId.trim() === "" || chartId === "unknown") {
        const derived = label ?? extractDrawingObjectName(rawXml) ?? graphicFramePlaceholderLabel(rawXml) ?? undefined;
        return { type: "unknown", ...(rawXml ? { rawXml } : {}), ...(derived ? { label: derived } : {}) };
      }
      return { type: "chart", chartId, ...(label ? { label } : {}), ...(rawXml ? { rawXml } : {}) };
    }
    default:
      return null;
  }
}

/**
 * Convert a DocumentController sheet drawings list (or other JSON-serializable drawings array)
 * into the UI overlay model.
 *
 * This adapter understands:
 * - DocumentController's simplified anchor schema (`{ type: "cell", row, col }` + pixel size)
 * - formula-model / Rust `DrawingObject` JSON (externally- or internally-tagged enums)
 *
 * Invalid entries are ignored (best-effort).
 */
export function convertDocumentSheetDrawingsToUiDrawingObjects(
  drawingsJson: unknown,
  context?: { sheetId?: string | number },
): DrawingObject[] {
  if (!Array.isArray(drawingsJson)) return [];

  const defaultSheetId = parseSheetId(context?.sheetId);
  const out: DrawingObject[] = [];
  for (const raw of drawingsJson) {
    try {
      if (isRecord(raw)) {
        // Best-effort passthrough for metadata authored by the UI layer
        // (e.g. rotation interactions or XLSX compatibility XML).
        const preservedValue = pick(raw, ["preserved"]);
        let preserved: Record<string, string> | undefined;
        if (isRecord(preservedValue)) {
          const outPreserved: Record<string, string> = {};
          for (const [k, v] of Object.entries(preservedValue)) {
            if (typeof v === "string") outPreserved[k] = v;
          }
          if (Object.keys(outPreserved).length > 0) preserved = outPreserved;
        }

        const transformValue = pick(raw, ["transform"]);
        let transform: DrawingTransform | undefined;
        if (isRecord(transformValue)) {
          const record = transformValue as JsonRecord;
          const hasAnyTransformKey =
            Object.prototype.hasOwnProperty.call(record, "rotationDeg") ||
            Object.prototype.hasOwnProperty.call(record, "rotation_deg") ||
            Object.prototype.hasOwnProperty.call(record, "flipH") ||
            Object.prototype.hasOwnProperty.call(record, "flip_h") ||
            Object.prototype.hasOwnProperty.call(record, "flipV") ||
            Object.prototype.hasOwnProperty.call(record, "flip_v");
          if (hasAnyTransformKey) {
            const rotationRaw = pick(record, ["rotationDeg", "rotation_deg"]);
            const rotationDeg = rotationRaw === undefined ? 0 : readOptionalNumber(rotationRaw);
            if (rotationDeg != null) {
              const flipHRaw = pick(record, ["flipH", "flip_h"]);
              const flipVRaw = pick(record, ["flipV", "flip_v"]);
              const flipH = flipHRaw === undefined ? false : flipHRaw;
              const flipV = flipVRaw === undefined ? false : flipVRaw;
              if (typeof flipH === "boolean" && typeof flipV === "boolean") {
                transform = { rotationDeg, flipH, flipV };
              }
            }
          }
        }

        const anchorValue = pick(raw, ["anchor"]);
        const anchorSheetId = (() => {
          if (!isRecord(anchorValue)) return undefined;
          return parseSheetId(pick(anchorValue, ["sheetId", "sheet_id"]));
        })();
        const sheetId = anchorSheetId ?? defaultSheetId;

        const kindValue = pick(raw, ["kind"]);
        const kind = convertDocumentDrawingKindToUiKind(kindValue);
        if (kind) {
          const id = parseDrawingObjectId(pick(raw, ["id"]));
          const zOrder = readOptionalNumber(pick(raw, ["zOrder", "z_order"])) ?? 0;
          const size = convertDocumentDrawingSizeToEmu(pick(raw, ["size"]));
          const anchor = convertDocumentDrawingAnchorToUiAnchor(anchorValue, size);
          if (anchor) {
            const derivedTransform =
              transform ??
              (() => {
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

            const obj: DrawingObject = { id, kind, anchor, zOrder, ...(size ? { size } : {}) };
            if (preserved) obj.preserved = preserved;
            if (derivedTransform) obj.transform = derivedTransform;
            out.push(obj);
            continue;
          }

          // Some DocumentController drawings use the simplified `{ kind: { type: ... } }` encoding
          // but store anchors in the formula-model/Rust enum shape (e.g. `{ Absolute: {...} }`),
          // such as imported chart objects hydrated from the XLSX backend.
          //
          // If the anchor doesn't match the DocumentController schema, fall through to the
          // formula-model adapter rather than dropping the object entirely.
        }

        // Some DocumentController drawings store `kind` in the formula-model/Rust enum encoding
        // (e.g. `{ Image: {...} }`) while storing anchors in the simplified DocumentController
        // schema (including legacy `{ type: "cell" }` anchors). The formula-model adapter cannot
        // parse legacy anchors, so try a mixed-mode conversion: model kind + DocumentController
        // anchor.
        {
          const id = parseDrawingObjectId(pick(raw, ["id"]));
          const zOrder = readOptionalNumber(pick(raw, ["zOrder", "z_order"])) ?? 0;
          const size = convertDocumentDrawingSizeToEmu(pick(raw, ["size"]));
          const anchor = convertDocumentDrawingAnchorToUiAnchor(anchorValue, size);
          if (anchor) {
            try {
              const modelKind = convertModelDrawingObjectKind(kindValue, { sheetId: sheetId ?? undefined, drawingObjectId: id });
              const label = isRecord(kindValue) ? readOptionalString(pick(kindValue, ["label"])) : undefined;
              const patchedKind: DrawingObjectKind =
                label &&
                (modelKind.type === "shape" || modelKind.type === "chart" || modelKind.type === "unknown") &&
                !(typeof (modelKind as any).label === "string" && String((modelKind as any).label).trim() !== "")
                  ? ({ ...modelKind, label } as DrawingObjectKind)
                  : modelKind;

              const derivedTransform =
                transform ??
                (() => {
                  if (patchedKind.type === "image") {
                    const picXml = preserved?.["xlsx.pic_xml"];
                    if (typeof picXml !== "string" || picXml.length === 0) return undefined;
                    const parsed = parseDrawingTransformFromRawXml(picXml);
                    if (!parsed) return undefined;
                    return parsed.rotationDeg !== 0 || parsed.flipH || parsed.flipV ? parsed : undefined;
                  }

                  const rawXml = (patchedKind as any).rawXml ?? (patchedKind as any).raw_xml;
                  if (typeof rawXml !== "string" || rawXml.length === 0) return undefined;
                  const parsed = parseDrawingTransformFromRawXml(rawXml);
                  if (!parsed) return undefined;
                  return parsed.rotationDeg !== 0 || parsed.flipH || parsed.flipV ? parsed : undefined;
                })();

              const obj: DrawingObject = { id, kind: patchedKind, anchor, zOrder, ...(size ? { size } : {}) };
              if (preserved) obj.preserved = preserved;
              if (derivedTransform) obj.transform = derivedTransform;
              out.push(obj);
              continue;
            } catch {
              // ignore
            }
          }
        }

        // If the drawing doesn't match the DocumentController schema, fall through to the
        // formula-model adapter but preserve any known sheet context so chart placeholders can
        // construct stable chart ids.
        const base = convertModelDrawingObjectToUiDrawingObject(raw, sheetId ? { sheetId } : undefined);
        // Preserve metadata even when parsing via the formula-model adapter (best-effort).
        // This ensures UI-authored rotation/preserved XML survives when DocumentController stores
        // drawings in a model-like shape.
        const merged: DrawingObject =
          preserved || transform
            ? { ...base, ...(preserved ? { preserved } : {}), ...(transform ? { transform } : {}) }
            : base;
        out.push(merged);
        continue;
      }

      // Fallback to formula-model conversion (externally-tagged enums).
      out.push(convertModelDrawingObjectToUiDrawingObject(raw));
    } catch {
      // ignore
    }
  }

  out.sort((a, b) => a.zOrder - b.zOrder);
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
