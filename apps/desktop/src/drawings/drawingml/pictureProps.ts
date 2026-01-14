export type DrawingMLPictureCrop = {
  /** Crop from left, DrawingML percent in [0, 100000]. */
  l: number;
  /** Crop from top, DrawingML percent in [0, 100000]. */
  t: number;
  /** Crop from right, DrawingML percent in [0, 100000]. */
  r: number;
  /** Crop from bottom, DrawingML percent in [0, 100000]. */
  b: number;
};

export type DrawingMLPictureOutline = {
  /**
   * Line width in DrawingML EMUs (1pt = 12700 EMU).
   *
   * Convert at render time via `emuToPx(…, zoom)` so it scales with sheet zoom.
   */
  widthEmu?: number;
  /** CSS color string (currently supports `solidFill/srgbClr`). */
  color?: string;
};

export type DrawingMLPictureProps = {
  crop?: DrawingMLPictureCrop;
  outline?: DrawingMLPictureOutline;
};

const PIC_XML_CACHE_MAX = 256;

// LRU cache keyed by the raw `xlsx.pic_xml` string.
const picXmlPropsCache = new Map<string, DrawingMLPictureProps | null>();

function clampNumber(value: number, min: number, max: number): number {
  if (value < min) return min;
  if (value > max) return max;
  return value;
}

function normalizeHexColor(value: string | null): string | undefined {
  if (value == null) return undefined;
  const raw = value.trim();
  if (!raw) return undefined;
  const hex = raw.replace(/^#/, "");
  if (!/^[0-9a-fA-F]{6,8}$/.test(hex)) return undefined;
  // If alpha is included, drop it. DrawingML uses the first 6 digits for RGB.
  return `#${hex.slice(0, 6).toUpperCase()}`;
}

function getAttrFromTag(tagXml: string, name: string): string | null {
  const re = new RegExp(`\\b${name}\\s*=\\s*(?:\"([^\"]*)\"|'([^']*)')`, "i");
  const m = re.exec(tagXml);
  return m ? (m[1] ?? m[2] ?? null) : null;
}

function extractFirstElementXml(xml: string, localName: string): string | null {
  const tag = localName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const prefix = "(?:[A-Za-z_][\\w.-]*:)?";

  const selfClosing = new RegExp(`<\\s*${prefix}${tag}\\b[^>]*\\/\\s*>`, "i");
  const sc = selfClosing.exec(xml);
  if (sc) return sc[0];

  const full = new RegExp(
    `<\\s*${prefix}${tag}\\b[^>]*>[\\s\\S]*?<\\/\\s*${prefix}${tag}\\s*>`,
    "i",
  );
  const m = full.exec(xml);
  return m ? m[0] : null;
}

function parsePct(value: string | null): number {
  if (value == null) return 0;
  const n = Number.parseInt(value, 10);
  if (!Number.isFinite(n)) return 0;
  return clampNumber(n, 0, 100_000);
}

function parseIntAttr(value: string | null): number | undefined {
  if (value == null) return undefined;
  const n = Number.parseInt(value, 10);
  return Number.isFinite(n) ? n : undefined;
}

function parsePictureCropFromXml(xml: string): DrawingMLPictureCrop | undefined {
  const srcRectTag = extractFirstElementXml(xml, "srcRect");
  if (!srcRectTag) return undefined;
  return {
    l: parsePct(getAttrFromTag(srcRectTag, "l")),
    t: parsePct(getAttrFromTag(srcRectTag, "t")),
    r: parsePct(getAttrFromTag(srcRectTag, "r")),
    b: parsePct(getAttrFromTag(srcRectTag, "b")),
  };
}

function parsePictureOutlineFromXml(xml: string): DrawingMLPictureOutline | undefined {
  const spPrXml = extractFirstElementXml(xml, "spPr") ?? xml;
  const lnXml = extractFirstElementXml(spPrXml, "ln");
  if (!lnXml) return undefined;

  // Extract `w` from the ln start tag.
  const lnStart = /<[^>]+>/.exec(lnXml)?.[0] ?? lnXml;
  const widthEmu = parseIntAttr(getAttrFromTag(lnStart, "w"));

  // Support the common `solidFill/srgbClr` case (ignore theme colors for now).
  const srgbTag = extractFirstElementXml(lnXml, "srgbClr");
  const color = normalizeHexColor(srgbTag ? getAttrFromTag(srgbTag, "val") : null);

  // Best-effort: only treat this as an outline when we can resolve a concrete color.
  if (!color) return undefined;

  const out: DrawingMLPictureOutline = { color };
  if (widthEmu !== undefined) out.widthEmu = widthEmu;
  return out;
}

function parseDrawingMLPicturePropsFromPicXml(picXml: string): DrawingMLPictureProps | null {
  const xml = String(picXml ?? "");
  if (!xml.trim()) return null;
  try {
    const crop = parsePictureCropFromXml(xml);
    const outline = parsePictureOutlineFromXml(xml);
    if (!crop && !outline) return null;
    const out: DrawingMLPictureProps = {};
    if (crop) out.crop = crop;
    if (outline) out.outline = outline;
    return out;
  } catch {
    return null;
  }
}

/**
 * Cached (LRU) parser for preserved `xlsx.pic_xml` fragments.
 *
 * These fragments are preserved verbatim for XLSX round-tripping; parsing them
 * on every render quickly becomes hot-path work when many pictures exist.
 */
export function getDrawingMLPicturePropsFromPicXml(picXml: string): DrawingMLPictureProps | null {
  const key = String(picXml ?? "");
  if (!key) return null;

  if (picXmlPropsCache.has(key)) {
    const cached = picXmlPropsCache.get(key) ?? null;
    // Refresh LRU.
    picXmlPropsCache.delete(key);
    picXmlPropsCache.set(key, cached);
    return cached;
  }

  const parsed = parseDrawingMLPicturePropsFromPicXml(key);
  picXmlPropsCache.set(key, parsed);

  // LRU eviction.
  if (picXmlPropsCache.size > PIC_XML_CACHE_MAX) {
    const oldest = picXmlPropsCache.keys().next().value as string | undefined;
    if (oldest !== undefined) picXmlPropsCache.delete(oldest);
  }
  return parsed;
}

