import opentype from "opentype.js";

/**
 * @typedef {import("../font.js").FontSpec} FontSpec
 */

/**
 * @typedef {{
 *   key: string,
 *   family: string,
 *   weight: string | number,
 *   style: string,
 *   upem: number,
 *   ascentRatio: number,
 *   descentRatio: number,
 *   hbBlob: { ptr: number, destroy: () => void },
 *   hbFace: { ptr: number, upem: number, destroy: () => void },
 *   hbFont: { ptr: number, setScale: (x: number, y: number) => void, destroy: () => void },
 *   otFont: any,
 *   supportsText: (text: string) => boolean,
 * }} HarfBuzzFontFace
 */

/**
 * @param {FontSpec | Omit<FontSpec, "sizePx">} spec
 */
function normalizeFaceSpec(spec) {
  return {
    family: spec.family,
    weight: spec.weight ?? 400,
    style: spec.style ?? "normal",
  };
}

/**
 * @param {ReturnType<typeof normalizeFaceSpec>} spec
 */
function faceKey(spec) {
  return `${spec.style}|${spec.weight}|${spec.family}`;
}

/**
 * @param {ArrayBuffer | ArrayBufferView} data
 * @returns {ArrayBuffer}
 */
function toArrayBuffer(data) {
  if (data instanceof ArrayBuffer) return data;
  if (ArrayBuffer.isView(data)) {
    return data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
  }
  throw new TypeError("Expected ArrayBuffer or ArrayBufferView");
}

/**
 * HarfBuzz font fallback should ignore formatting/codepoints that aren't expected to map to glyphs.
 *
 * This is not a complete `Default_Ignorable_Code_Point` implementation, but it covers the characters
 * commonly encountered in spreadsheet cell text (emoji variation selectors, joiners, bidi marks).
 *
 * @param {number} cp
 */
function isIgnorableForGlyphFallback(cp) {
  // Variation Selectors + Variation Selectors Supplement.
  if ((cp >= 0xfe00 && cp <= 0xfe0f) || (cp >= 0xe0100 && cp <= 0xe01ef)) return true;
  // Join controls.
  if (cp === 0x200c || cp === 0x200d) return true; // ZWNJ, ZWJ
  // Common bidi marks / format controls.
  if (
    cp === 0x200e || // LRM
    cp === 0x200f || // RLM
    (cp >= 0x202a && cp <= 0x202e) || // embedding/override
    (cp >= 0x2066 && cp <= 0x2069) // isolates
  )
    return true;
  return false;
}

export class HarfBuzzFontManager {
  /**
   * @param {any} hb HarfBuzz instance returned from `loadHarfBuzz()`.
   */
  constructor(hb) {
    this.hb = hb;

    /** @type {Map<string, HarfBuzzFontFace>} */
    this.facesByKey = new Map();
    /** @type {Map<string, HarfBuzzFontFace[]>} */
    this.facesByFamily = new Map();

    /** @type {string[]} */
    this.fallbackFamilies = [];

    // Incremented when font data / fallback configuration changes.
    this.version = 0;
  }

  /**
   * Load a font from raw bytes.
   *
   * The font is cached by `{family, weight, style}` (size is intentionally excluded).
   *
   * @param {ArrayBuffer | ArrayBufferView} data
   * @param {Omit<FontSpec, "sizePx">} spec
   * @returns {HarfBuzzFontFace}
   */
  loadFont(data, spec) {
    const normalized = normalizeFaceSpec(spec);
    const key = faceKey(normalized);

    const existing = this.facesByKey.get(key);
    if (existing) {
      // Replace in-place, but free previous WASM allocations to avoid leaking if consumers reload fonts.
      existing.hbFont.destroy();
      existing.hbFace.destroy();
      existing.hbBlob.destroy();
    }

    const arrayBuffer = toArrayBuffer(data);
    const otFont = opentype.parse(arrayBuffer);

    const blob = this.hb.createBlob(arrayBuffer);
    const face = this.hb.createFace(blob, 0);
    const font = this.hb.createFont(face);
    font.setScale(face.upem, face.upem);

    const asc = otFont.ascender;
    const desc = otFont.descender;
    const denom = asc - desc;
    const ascentRatio = Number.isFinite(denom) && denom !== 0 ? asc / denom : 0.8;
    const descentRatio = Number.isFinite(denom) && denom !== 0 ? -desc / denom : 0.2;

    /** @type {Map<number, boolean>} */
    const glyphSupportCache = new Map();

    /** @type {HarfBuzzFontFace} */
    const faceObj = {
      key,
      family: normalized.family,
      weight: normalized.weight,
      style: normalized.style,
      upem: face.upem,
      ascentRatio,
      descentRatio,
      hbBlob: blob,
      hbFace: face,
      hbFont: font,
      otFont,
      supportsText: (text) => {
        for (const ch of text) {
          const cp = ch.codePointAt(0);
          if (cp === undefined) continue;
          if (isIgnorableForGlyphFallback(cp)) continue;

          const cached = glyphSupportCache.get(cp);
          if (cached !== undefined) {
            if (!cached) return false;
            continue;
          }

          const gid = otFont.charToGlyphIndex(ch);
          const ok = gid !== 0;
          glyphSupportCache.set(cp, ok);
          if (!ok) return false;
        }
        return true;
      },
    };

    this.facesByKey.set(key, faceObj);

    const familyList = this.facesByFamily.get(normalized.family) ?? [];
    // Replace an existing face with the same key to keep insertion order stable.
    const existingIdx = familyList.findIndex((f) => f.key === key);
    if (existingIdx >= 0) familyList[existingIdx] = faceObj;
    else familyList.push(faceObj);
    this.facesByFamily.set(normalized.family, familyList);

    this.version++;
    return faceObj;
  }

  /**
   * Configure the global fallback family order.
   *
   * @param {string[]} families
   */
  setFallbackFamilies(families) {
    this.fallbackFamilies = [...families];
    this.version++;
  }

  /**
   * @param {FontSpec | Omit<FontSpec, "sizePx">} spec
   * @returns {HarfBuzzFontFace}
   */
  getFace(spec) {
    const normalized = normalizeFaceSpec(spec);
    const key = faceKey(normalized);
    const exact = this.facesByKey.get(key);
    if (exact) return exact;

    const family = this.facesByFamily.get(normalized.family);
    if (family && family.length) return family[0];

    throw new Error(`HarfBuzz font not loaded: ${normalized.family} (${normalized.style} ${normalized.weight})`);
  }

  /**
   * @param {FontSpec | Omit<FontSpec, "sizePx">} spec
   * @returns {HarfBuzzFontFace[]}
   */
  getFallbackFaces(spec) {
    const primary = this.getFace(spec);
    /** @type {HarfBuzzFontFace[]} */
    const out = [];

    for (const family of this.fallbackFamilies) {
      if (family === primary.family) continue;
      const faces = this.facesByFamily.get(family);
      if (!faces || faces.length === 0) continue;

      // Prefer the same style/weight if present for more predictable fallback.
      const normalized = normalizeFaceSpec({ ...spec, family });
      const wantedKey = faceKey(normalized);
      const exact = this.facesByKey.get(wantedKey);
      out.push(exact ?? faces[0]);
    }

    return out;
  }

  destroy() {
    for (const face of this.facesByKey.values()) {
      // Order matters: destroy font -> face -> blob.
      face.hbFont.destroy();
      face.hbFace.destroy();
      face.hbBlob.destroy();
    }
    this.facesByKey.clear();
    this.facesByFamily.clear();
    this.version++;
  }
}
