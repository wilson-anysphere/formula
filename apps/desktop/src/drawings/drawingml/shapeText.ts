export type ShapeTextRun = {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSizePt?: number;
  fontFamily?: string;
  color?: string;
};

export type ShapeTextLayout = {
  textRuns: ShapeTextRun[];
  alignment?: "left" | "center" | "right";
  vertical?: "top" | "middle" | "bottom";
  wrap?: boolean;
  /** Optional text insets from `<a:bodyPr lIns/tIns/rIns/bIns>` (DrawingML EMUs). */
  insetLeftEmu?: number;
  insetTopEmu?: number;
  insetRightEmu?: number;
  insetBottomEmu?: number;
};

type RunStyle = Omit<ShapeTextRun, "text">;

type BulletDef =
  | { kind: "none" }
  | { kind: "char"; char: string }
  | { kind: "auto"; type: string; startAt: number };

function parseParagraphLevel(value: string | null): number {
  if (value == null) return 0;
  const n = Number.parseInt(value, 10);
  if (!Number.isFinite(n) || n < 0) return 0;
  return Math.min(8, n);
}

function parseStartAt(value: string | null): number {
  if (value == null) return 1;
  const n = Number.parseInt(value, 10);
  return Number.isFinite(n) && n > 0 ? n : 1;
}

const KNOWN_NAMESPACE_URIS: Record<string, string> = {
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  xdr: "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
  a14: "http://schemas.microsoft.com/office/drawing/2010/main",
  a15: "http://schemas.microsoft.com/office/drawing/2012/main",
  a16: "http://schemas.microsoft.com/office/drawing/2014/main",
};

function mergeStyle(base: RunStyle, patch: RunStyle): RunStyle {
  // Shallow merge, but only apply defined values so callers can pass `{}`.
  const out: RunStyle = { ...base };
  for (const [k, v] of Object.entries(patch) as Array<[keyof RunStyle, RunStyle[keyof RunStyle]]>) {
    if (v !== undefined) {
      (out as any)[k] = v;
    }
  }
  return out;
}

function alphaIndex(num: number, upper: boolean): string {
  if (!Number.isFinite(num) || num <= 0) return "";
  let n = Math.trunc(num);
  let out = "";
  while (n > 0) {
    n -= 1;
    const ch = String.fromCharCode((n % 26) + (upper ? 65 : 97));
    out = ch + out;
    n = Math.floor(n / 26);
  }
  return out;
}

function romanNumeral(num: number): string {
  if (!Number.isFinite(num) || num <= 0) return "";
  let n = Math.trunc(num);
  if (n > 3999) return String(n);
  const numerals: Array<[number, string]> = [
    [1000, "M"],
    [900, "CM"],
    [500, "D"],
    [400, "CD"],
    [100, "C"],
    [90, "XC"],
    [50, "L"],
    [40, "XL"],
    [10, "X"],
    [9, "IX"],
    [5, "V"],
    [4, "IV"],
    [1, "I"],
  ];
  let out = "";
  for (const [value, glyph] of numerals) {
    while (n >= value) {
      out += glyph;
      n -= value;
    }
  }
  return out;
}

function formatAutoNumber(type: string, num: number): string {
  const n = Math.max(1, Math.trunc(num));
  switch (type) {
    case "arabicPeriod":
      return `${n}.`;
    case "arabicParenR":
      return `${n})`;
    case "arabicParenBoth":
      return `(${n})`;
    case "arabicPlain":
      return `${n}`;
    case "alphaLcPeriod":
      return `${alphaIndex(n, false)}.`;
    case "alphaUcPeriod":
      return `${alphaIndex(n, true)}.`;
    case "alphaLcParenR":
      return `${alphaIndex(n, false)})`;
    case "alphaUcParenR":
      return `${alphaIndex(n, true)})`;
    case "alphaLcParenBoth":
      return `(${alphaIndex(n, false)})`;
    case "alphaUcParenBoth":
      return `(${alphaIndex(n, true)})`;
    case "romanLcPeriod":
      return `${romanNumeral(n).toLowerCase()}.`;
    case "romanUcPeriod":
      return `${romanNumeral(n)}.`;
    case "romanLcParenR":
      return `${romanNumeral(n).toLowerCase()})`;
    case "romanUcParenR":
      return `${romanNumeral(n)})`;
    case "romanLcParenBoth":
      return `(${romanNumeral(n).toLowerCase()})`;
    case "romanUcParenBoth":
      return `(${romanNumeral(n)})`;
    default:
      return `${n}.`;
  }
}

function parseBoolAttr(value: string | null): boolean | undefined {
  if (value == null) return undefined;
  const normalized = value.trim().toLowerCase();
  if (normalized === "1" || normalized === "true" || normalized === "yes" || normalized === "on") return true;
  if (normalized === "0" || normalized === "false" || normalized === "no" || normalized === "off") return false;
  return undefined;
}

function parseUnderlineAttr(value: string | null): boolean | undefined {
  if (value == null) return undefined;
  const normalized = value.trim().toLowerCase();
  if (normalized === "none" || normalized === "0" || normalized === "false" || normalized === "off") return false;
  return true;
}

function parseFontSizePt(sz: string | null): number | undefined {
  if (sz == null) return undefined;
  const n = Number.parseInt(sz, 10);
  if (!Number.isFinite(n) || n <= 0) return undefined;
  // DrawingML `sz` is in 1/100 of a point.
  return n / 100;
}

function parseNonNegativeInt(value: string | null): number | undefined {
  if (value == null) return undefined;
  const n = Number.parseInt(value, 10);
  if (!Number.isFinite(n) || n < 0) return undefined;
  return n;
}

function normalizeHexColor(value: string | null): string | undefined {
  if (value == null) return undefined;
  const raw = value.trim();
  if (!raw) return undefined;
  const hex = raw.replace(/^#/, "");
  if (!/^[0-9a-fA-F]{6,8}$/.test(hex)) return undefined;
  // If alpha is included, drop it. (DrawingML uses the first 6 digits for RGB.)
  return `#${hex.slice(0, 6).toUpperCase()}`;
}

function parseAlignment(value: string | null): ShapeTextLayout["alignment"] | undefined {
  if (value == null) return undefined;
  switch (value.trim().toLowerCase()) {
    case "l":
    case "left":
      return "left";
    case "ctr":
    case "center":
    case "centre":
      return "center";
    case "r":
    case "right":
      return "right";
    default:
      return undefined;
  }
}

function parseVertical(value: string | null): ShapeTextLayout["vertical"] | undefined {
  if (value == null) return undefined;
  switch (value.trim().toLowerCase()) {
    case "t":
    case "top":
      return "top";
    case "ctr":
    case "center":
    case "centre":
      return "middle";
    case "b":
    case "bottom":
      return "bottom";
    default:
      return undefined;
  }
}

function parseWrap(value: string | null): boolean | undefined {
  if (value == null) return undefined;
  const normalized = value.trim().toLowerCase();
  if (normalized === "none") return false;
  return true;
}

function findFirstByLocalName(root: ParentNode, localName: string): Element | null {
  const walker = (node: ParentNode): Element | null => {
    for (const child of Array.from(node.childNodes)) {
      if (child.nodeType !== 1) continue;
      const el = child as Element;
      if (el.localName === localName) return el;
      const found = walker(el);
      if (found) return found;
    }
    return null;
  };
  return walker(root);
}

function collectTextInDescendants(root: ParentNode, localName: string): string {
  let out = "";
  const walk = (node: ParentNode) => {
    for (const child of Array.from(node.childNodes)) {
      if (child.nodeType === 3) continue;
      if (child.nodeType !== 1) continue;
      const el = child as Element;
      if (el.localName === localName) {
        out += el.textContent ?? "";
      } else {
        walk(el);
      }
    }
  };
  walk(root);
  return out;
}

function parseBulletDefFromDom(pPr: Element): BulletDef | null {
  if (findFirstByLocalName(pPr, "buNone")) return { kind: "none" };

  const buAuto = findFirstByLocalName(pPr, "buAutoNum");
  if (buAuto) {
    const type = buAuto.getAttribute("type")?.trim() || "arabicPeriod";
    const startAt = parseStartAt(buAuto.getAttribute("startAt"));
    return { kind: "auto", type, startAt };
  }

  const buChar = findFirstByLocalName(pPr, "buChar");
  const bullet = buChar?.getAttribute("char")?.trim();
  if (bullet) return { kind: "char", char: bullet };
  return null;
}

function parseListStyleBulletsFromDom(lstStyle: Element | null): Map<number, BulletDef> {
  const bullets = new Map<number, BulletDef>();
  if (!lstStyle) return bullets;
  for (let level = 0; level < 9; level += 1) {
    const lvlPPr = findFirstByLocalName(lstStyle, `lvl${level + 1}pPr`);
    if (!lvlPPr) continue;
    const def = parseBulletDefFromDom(lvlPPr);
    if (def) bullets.set(level, def);
  }
  return bullets;
}

function parseListStyleRunStylesFromDom(lstStyle: Element | null): Map<number, RunStyle> {
  const styles = new Map<number, RunStyle>();
  if (!lstStyle) return styles;
  for (let level = 0; level < 9; level += 1) {
    const lvlPPr = findFirstByLocalName(lstStyle, `lvl${level + 1}pPr`);
    if (!lvlPPr) continue;
    const defRPr = findFirstByLocalName(lvlPPr, "defRPr");
    if (!defRPr) continue;
    const style = parseRunStyleFromDom(defRPr);
    if (Object.keys(style).length > 0) styles.set(level, style);
  }
  return styles;
}

function parseRunStyleFromDom(el: Element | null): RunStyle {
  if (!el) return {};
  const bold = parseBoolAttr(el.getAttribute("b"));
  const italic = parseBoolAttr(el.getAttribute("i"));
  const underline = parseUnderlineAttr(el.getAttribute("u"));
  const fontSizePt = parseFontSizePt(el.getAttribute("sz"));

  let fontFamily: string | undefined;
  const latin = findFirstByLocalName(el, "latin");
  const typeface = latin?.getAttribute("typeface");
  if (typeface && typeface.trim()) {
    fontFamily = typeface.trim();
  }

  let color: string | undefined;
  const solidFill = findFirstByLocalName(el, "solidFill");
  if (solidFill) {
    const srgb = findFirstByLocalName(solidFill, "srgbClr");
    color = normalizeHexColor(srgb?.getAttribute("val"));
  }

  const style: RunStyle = {};
  if (bold !== undefined) style.bold = bold;
  if (italic !== undefined) style.italic = italic;
  if (underline !== undefined) style.underline = underline;
  if (fontSizePt !== undefined) style.fontSizePt = fontSizePt;
  if (fontFamily !== undefined) style.fontFamily = fontFamily;
  if (color !== undefined) style.color = color;
  return style;
}

function buildXmlWrapper(rawXml: string): string {
  // Raw snippets are extracted from a larger DrawingML document and typically do *not* include
  // namespace declarations. DOMParser in XML mode rejects unbound prefixes ("xdr:", "a:", etc),
  // so we synthesize a wrapper element that binds every prefix we see in the payload.
  const prefixes = new Set<string>();
  const tagPrefixRe = /<\/?\s*([A-Za-z_][\w.-]*):[A-Za-z_][\w.-]*/g;
  const attrPrefixRe = /\s([A-Za-z_][\w.-]*):[A-Za-z_][\w.-]*\s*=/g;
  for (const match of rawXml.matchAll(tagPrefixRe)) {
    const prefix = match[1];
    if (prefix && prefix !== "xml" && prefix !== "xmlns") prefixes.add(prefix);
  }
  for (const match of rawXml.matchAll(attrPrefixRe)) {
    const prefix = match[1];
    // `xmlns:` attributes declare namespaces and use the reserved `xmlns` prefix; attempting to
    // bind it in our wrapper would produce invalid XML (`xmlns:xmlns="..."`). Skip it.
    if (prefix && prefix !== "xml" && prefix !== "xmlns") prefixes.add(prefix);
  }

  const attrs: string[] = [];
  for (const prefix of prefixes) {
    const uri = KNOWN_NAMESPACE_URIS[prefix] ?? `urn:formula:drawingml:${prefix}`;
    attrs.push(`xmlns:${prefix}="${uri}"`);
  }
  const decls = attrs.length ? ` ${attrs.join(" ")}` : "";
  return `<root${decls}>${rawXml}</root>`;
}

function parseShapeTextDom(rawXml: string): ShapeTextLayout | null {
  const wrapped = buildXmlWrapper(rawXml);
  const doc = new DOMParser().parseFromString(wrapped, "application/xml");
  if (doc.getElementsByTagName("parsererror").length > 0) return null;

  const txBody = findFirstByLocalName(doc, "txBody");
  if (!txBody) return { textRuns: [] };

  const bodyPr = findFirstByLocalName(txBody, "bodyPr");
  const vertical = parseVertical(bodyPr?.getAttribute("anchor") ?? null);
  const wrap = parseWrap(bodyPr?.getAttribute("wrap") ?? null);
  const insetLeftEmu = parseNonNegativeInt(bodyPr?.getAttribute("lIns") ?? null);
  const insetTopEmu = parseNonNegativeInt(bodyPr?.getAttribute("tIns") ?? null);
  const insetRightEmu = parseNonNegativeInt(bodyPr?.getAttribute("rIns") ?? null);
  const insetBottomEmu = parseNonNegativeInt(bodyPr?.getAttribute("bIns") ?? null);

  const shapeDefaultStyle = parseRunStyleFromDom(findFirstByLocalName(txBody, "defRPr"));

  const paragraphs = Array.from(txBody.childNodes).filter(
    (n): n is Element => n.nodeType === 1 && (n as Element).localName === "p",
  );

  const lstStyle = findFirstByLocalName(txBody, "lstStyle");
  const listBullets = parseListStyleBulletsFromDom(lstStyle);
  const listRunStyles = parseListStyleRunStylesFromDom(lstStyle);

  const textRuns: ShapeTextRun[] = [];
  let alignment: ShapeTextLayout["alignment"] | undefined;
  const autoNumbers = new Map<number, { type: string; next: number }>();
  let prevLevel = 0;

  for (let pi = 0; pi < paragraphs.length; pi += 1) {
    const p = paragraphs[pi];
    const pPr = Array.from(p.childNodes).find((n): n is Element => n.nodeType === 1 && (n as Element).localName === "pPr") ?? null;
    const level = parseParagraphLevel(pPr?.getAttribute("lvl") ?? null);
    if (!alignment) {
      const pPrAlign = pPr?.getAttribute("algn") ?? null;
      alignment = parseAlignment(pPrAlign);
      if (!alignment && pPr) {
        const algnChild = findFirstByLocalName(pPr, "algn");
        alignment = parseAlignment(algnChild?.getAttribute("val") ?? null);
      }
    }

    const levelStyle = listRunStyles.get(level) ?? {};
    const paraDefaultStyle = mergeStyle(
      mergeStyle(shapeDefaultStyle, levelStyle),
      parseRunStyleFromDom(findFirstByLocalName(pPr ?? p, "defRPr")),
    );
    // Best-effort bullet handling: prepend a bullet character or auto-numbering prefix when
    // the paragraph defines a bullet (either directly on `<a:pPr>` or inherited via `<a:lstStyle>`).
    if (level < prevLevel) {
      for (const key of Array.from(autoNumbers.keys())) {
        if (key > level) autoNumbers.delete(key);
      }
    }
    prevLevel = level;

    let bulletDef: BulletDef | null = pPr ? parseBulletDefFromDom(pPr) : null;
    if (bulletDef == null) bulletDef = listBullets.get(level) ?? null;
    const indent = level > 0 ? "  ".repeat(level) : "";
    if (bulletDef?.kind === "auto") {
      let state = autoNumbers.get(level);
      if (!state || state.type !== bulletDef.type) {
        state = { type: bulletDef.type, next: bulletDef.startAt };
      }
      const prefix = formatAutoNumber(bulletDef.type, state.next);
      state.next += 1;
      autoNumbers.set(level, state);
      if (prefix) textRuns.push({ text: `${indent}${prefix} `, ...paraDefaultStyle });
    } else {
      autoNumbers.delete(level);
      if (bulletDef?.kind === "char") {
        textRuns.push({ text: `${indent}${bulletDef.char} `, ...paraDefaultStyle });
      } else if (indent) {
        // Even when bullets are disabled (e.g. `<a:buNone/>`), the paragraph level still implies
        // indentation. Preserve a minimal amount of indentation so nested list-like structures
        // remain readable.
        textRuns.push({ text: indent, ...paraDefaultStyle });
      }
    }

    for (const child of Array.from(p.childNodes)) {
      if (child.nodeType !== 1) continue;
      const el = child as Element;

      if (el.localName === "br") {
        textRuns.push({ text: "\n", ...paraDefaultStyle });
        continue;
      }

      if (el.localName === "tab") {
        textRuns.push({ text: "\t", ...paraDefaultStyle });
        continue;
      }

      if (el.localName !== "r" && el.localName !== "fld") continue;
      const rPr = findFirstByLocalName(el, "rPr");
      const style = mergeStyle(paraDefaultStyle, parseRunStyleFromDom(rPr));
      const text = collectTextInDescendants(el, "t");
      if (text.length === 0) continue;
      textRuns.push({ text, ...style });
    }

    if (pi !== paragraphs.length - 1) {
      textRuns.push({ text: "\n", ...shapeDefaultStyle });
    }
  }

  return { textRuns, alignment, vertical, wrap, insetLeftEmu, insetTopEmu, insetRightEmu, insetBottomEmu };
}

function decodeXmlEntities(text: string): string {
  return String(text)
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&")
    .replace(/&#(\d+);/g, (match, code) => {
      const cp = Number.parseInt(code, 10);
      if (!Number.isFinite(cp)) return match;
      try {
        return String.fromCodePoint(cp);
      } catch {
        return match;
      }
    })
    .replace(/&#x([0-9a-fA-F]+);/g, (match, code) => {
      const cp = Number.parseInt(code, 16);
      if (!Number.isFinite(cp)) return match;
      try {
        return String.fromCodePoint(cp);
      } catch {
        return match;
      }
    });
}

function getAttrFromTag(tagXml: string, name: string): string | null {
  const re = new RegExp(`\\b${name}\\s*=\\s*(?:\"([^\"]*)\"|'([^']*)')`, "i");
  const m = re.exec(tagXml);
  return m ? (m[1] ?? m[2] ?? null) : null;
}

function extractFirstElementXml(xml: string, localName: string): string | null {
  const tag = localName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const selfClosing = new RegExp(`<[^>]*?\\b${tag}\\b[^>]*/>`, "i");
  const sc = selfClosing.exec(xml);
  if (sc) return sc[0];

  const full = new RegExp(`<[^>]*?\\b${tag}\\b[^>]*>[\\s\\S]*?<\\/[^>]*?\\b${tag}\\b\\s*>`, "i");
  const m = full.exec(xml);
  return m ? m[0] : null;
}

function extractAllElementInner(xml: string, localName: string): string[] {
  const tag = localName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const re = new RegExp(`<[^>]*?\\b${tag}\\b[^>]*>([\\s\\S]*?)<\\/[^>]*?\\b${tag}\\b\\s*>`, "gi");
  const out: string[] = [];
  for (const m of xml.matchAll(re)) {
    out.push(m[1] ?? "");
  }
  return out;
}

function parseRunStyleFromXmlSnippet(xml: string | null): RunStyle {
  if (!xml) return {};
  const tagOpenMatch = /<[^>]+>/.exec(xml);
  const open = tagOpenMatch ? tagOpenMatch[0] : xml;

  const bold = parseBoolAttr(getAttrFromTag(open, "b"));
  const italic = parseBoolAttr(getAttrFromTag(open, "i"));
  const underline = parseUnderlineAttr(getAttrFromTag(open, "u"));
  const fontSizePt = parseFontSizePt(getAttrFromTag(open, "sz"));

  const latinMatch = /<[^>]*?\blatin\b[^>]*?\btypeface\s*=\s*(?:"([^"]*)"|'([^']*)')/i.exec(xml);
  const fontFamily = latinMatch ? (latinMatch[1] ?? latinMatch[2] ?? "").trim() : "";

  const srgbMatch = /<[^>]*?\bsrgbClr\b[^>]*?\bval\s*=\s*(?:"([^"]*)"|'([^']*)')/i.exec(xml);
  const color = normalizeHexColor(srgbMatch ? (srgbMatch[1] ?? srgbMatch[2] ?? null) : null);

  const style: RunStyle = {};
  if (bold !== undefined) style.bold = bold;
  if (italic !== undefined) style.italic = italic;
  if (underline !== undefined) style.underline = underline;
  if (fontSizePt !== undefined) style.fontSizePt = fontSizePt;
  if (fontFamily) style.fontFamily = fontFamily;
  if (color !== undefined) style.color = color;
  return style;
}

function parseBulletDefFromXml(pPrXml: string | null): BulletDef | null {
  if (!pPrXml) return null;
  if (extractFirstElementXml(pPrXml, "buNone")) return { kind: "none" };

  const buAutoXml = extractFirstElementXml(pPrXml, "buAutoNum");
  if (buAutoXml) {
    const buAutoOpen = /<[^>]+>/.exec(buAutoXml)?.[0] ?? buAutoXml;
    const type = (getAttrFromTag(buAutoOpen, "type") ?? "").trim() || "arabicPeriod";
    const startAt = parseStartAt(getAttrFromTag(buAutoOpen, "startAt"));
    return { kind: "auto", type, startAt };
  }

  const buCharXml = extractFirstElementXml(pPrXml, "buChar");
  if (buCharXml) {
    const buCharOpen = /<[^>]+>/.exec(buCharXml)?.[0] ?? buCharXml;
    const bulletRaw = getAttrFromTag(buCharOpen, "char");
    const bullet = bulletRaw ? decodeXmlEntities(bulletRaw) : "";
    const trimmed = bullet.trim();
    if (trimmed) return { kind: "char", char: trimmed };
  }

  return null;
}

function parseListStyleBulletsFromXml(lstStyleXml: string | null): Map<number, BulletDef> {
  const bullets = new Map<number, BulletDef>();
  if (!lstStyleXml) return bullets;
  for (let level = 0; level < 9; level += 1) {
    const lvlXml = extractFirstElementXml(lstStyleXml, `lvl${level + 1}pPr`);
    if (!lvlXml) continue;
    const def = parseBulletDefFromXml(lvlXml);
    if (def) bullets.set(level, def);
  }
  return bullets;
}

function parseListStyleRunStylesFromXml(lstStyleXml: string | null): Map<number, RunStyle> {
  const styles = new Map<number, RunStyle>();
  if (!lstStyleXml) return styles;
  for (let level = 0; level < 9; level += 1) {
    const lvlXml = extractFirstElementXml(lstStyleXml, `lvl${level + 1}pPr`);
    if (!lvlXml) continue;
    const defRPrXml = extractFirstElementXml(lvlXml, "defRPr");
    if (!defRPrXml) continue;
    const style = parseRunStyleFromXmlSnippet(defRPrXml);
    if (Object.keys(style).length > 0) styles.set(level, style);
  }
  return styles;
}

function parseShapeTextFallback(rawXml: string): ShapeTextLayout | null {
  const txBodyMatch = /<(?:[A-Za-z_][\w.-]*:)?txBody\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?txBody\s*>/i.exec(rawXml);
  if (!txBodyMatch) return { textRuns: [] };
  const txBodyXml = txBodyMatch[0];
  const txBodyInner = txBodyMatch[1] ?? "";

  const bodyPrXml = extractFirstElementXml(txBodyInner, "bodyPr");
  const bodyPrOpen = bodyPrXml ? /<[^>]+>/.exec(bodyPrXml)?.[0] ?? bodyPrXml : "";
  const vertical = parseVertical(getAttrFromTag(bodyPrOpen, "anchor"));
  const wrap = parseWrap(getAttrFromTag(bodyPrOpen, "wrap"));
  const insetLeftEmu = parseNonNegativeInt(getAttrFromTag(bodyPrOpen, "lIns"));
  const insetTopEmu = parseNonNegativeInt(getAttrFromTag(bodyPrOpen, "tIns"));
  const insetRightEmu = parseNonNegativeInt(getAttrFromTag(bodyPrOpen, "rIns"));
  const insetBottomEmu = parseNonNegativeInt(getAttrFromTag(bodyPrOpen, "bIns"));

  const shapeDefaultStyle = parseRunStyleFromXmlSnippet(extractFirstElementXml(txBodyXml, "defRPr"));
  const lstStyleXml = extractFirstElementXml(txBodyXml, "lstStyle");
  const listBullets = parseListStyleBulletsFromXml(lstStyleXml);
  const listRunStyles = parseListStyleRunStylesFromXml(lstStyleXml);

  const paragraphsInner = extractAllElementInner(txBodyInner, "p");
  const textRuns: ShapeTextRun[] = [];
  let alignment: ShapeTextLayout["alignment"] | undefined;
  const autoNumbers = new Map<number, { type: string; next: number }>();
  let prevLevel = 0;

  for (let pi = 0; pi < paragraphsInner.length; pi += 1) {
    const pInner = paragraphsInner[pi] ?? "";

    const pPrXml = extractFirstElementXml(pInner, "pPr");
    const pPrOpen = pPrXml ? /<[^>]+>/.exec(pPrXml)?.[0] ?? pPrXml : null;
    const level = parseParagraphLevel(getAttrFromTag(pPrOpen ?? "", "lvl"));
    if (!alignment) {
      alignment = parseAlignment(getAttrFromTag(pPrOpen ?? "", "algn"));
      if (!alignment && pPrXml) {
        const algnXml = extractFirstElementXml(pPrXml, "algn");
        const algnOpen = algnXml ? /<[^>]+>/.exec(algnXml)?.[0] ?? algnXml : "";
        alignment = parseAlignment(getAttrFromTag(algnOpen, "val"));
      }
    }

    const levelStyle = listRunStyles.get(level) ?? {};
    const paraDefaultStyle = mergeStyle(
      mergeStyle(shapeDefaultStyle, levelStyle),
      parseRunStyleFromXmlSnippet(extractFirstElementXml(pPrXml ?? "", "defRPr")),
    );
    // Best-effort bullet handling (see DOMParser path above).
    if (level < prevLevel) {
      for (const key of Array.from(autoNumbers.keys())) {
        if (key > level) autoNumbers.delete(key);
      }
    }
    prevLevel = level;

    let bulletDef: BulletDef | null = parseBulletDefFromXml(pPrXml);
    if (bulletDef == null) bulletDef = listBullets.get(level) ?? null;
    const indent = level > 0 ? "  ".repeat(level) : "";
    if (bulletDef?.kind === "auto") {
      let state = autoNumbers.get(level);
      if (!state || state.type !== bulletDef.type) {
        state = { type: bulletDef.type, next: bulletDef.startAt };
      }
      const prefix = formatAutoNumber(bulletDef.type, state.next);
      state.next += 1;
      autoNumbers.set(level, state);
      if (prefix) textRuns.push({ text: `${indent}${prefix} `, ...paraDefaultStyle });
    } else {
      autoNumbers.delete(level);
      if (bulletDef?.kind === "char") {
        textRuns.push({ text: `${indent}${bulletDef.char} `, ...paraDefaultStyle });
      } else if (indent) {
        textRuns.push({ text: indent, ...paraDefaultStyle });
      }
    }

    // Walk runs + breaks/tabs in order within the paragraph.
    const nodeRe =
      /<(?:[A-Za-z_][\w.-]*:)?(r|fld)\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?\1\s*>|<(?:[A-Za-z_][\w.-]*:)?br\b[^>]*\/>|<(?:[A-Za-z_][\w.-]*:)?br\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?br\s*>|<(?:[A-Za-z_][\w.-]*:)?tab\b[^>]*\/>|<(?:[A-Za-z_][\w.-]*:)?tab\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?tab\s*>/gi;
    for (const match of pInner.matchAll(nodeRe)) {
      const chunk = match[0] ?? "";
      if (/<(?:[A-Za-z_][\w.-]*:)?br\b/i.test(chunk)) {
        textRuns.push({ text: "\n", ...paraDefaultStyle });
        continue;
      }

      if (/<(?:[A-Za-z_][\w.-]*:)?tab\b/i.test(chunk)) {
        textRuns.push({ text: "\t", ...paraDefaultStyle });
        continue;
      }

      const runPrXml = extractFirstElementXml(chunk, "rPr");
      const style = mergeStyle(paraDefaultStyle, parseRunStyleFromXmlSnippet(runPrXml));
      const texts = extractAllElementInner(chunk, "t").map((t) => decodeXmlEntities(t));
      const text = texts.join("");
      if (text.length === 0) continue;
      textRuns.push({ text, ...style });
    }

    if (pi !== paragraphsInner.length - 1) {
      textRuns.push({ text: "\n", ...shapeDefaultStyle });
    }
  }

  return { textRuns, alignment, vertical, wrap, insetLeftEmu, insetTopEmu, insetRightEmu, insetBottomEmu };
}

/**
 * Best-effort DrawingML `<xdr:txBody>` parser for shapes (text boxes, labeled shapes).
 *
 * This is intentionally minimal — it focuses on extracting user-visible text + a small
 * subset of run styles so we can render basic Excel-like shape labels on the overlay canvas.
 */
export function parseDrawingMLShapeText(rawXml: string): ShapeTextLayout | null {
  const xml = String(rawXml ?? "");
  if (xml.trim() === "") return null;
  try {
    if (typeof DOMParser !== "undefined") {
      const parsed = parseShapeTextDom(xml);
      if (parsed) return parsed;
      // Fall back to regex parsing if the DOM path fails (e.g. unexpected namespace prefixes).
    }
    return parseShapeTextFallback(xml);
  } catch {
    return null;
  }
}
