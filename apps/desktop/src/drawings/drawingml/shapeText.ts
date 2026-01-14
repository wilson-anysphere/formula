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
};

type RunStyle = Omit<ShapeTextRun, "text">;

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

function parseBoolAttr(value: string | null): boolean | undefined {
  if (value == null) return undefined;
  const normalized = value.trim().toLowerCase();
  if (normalized === "1" || normalized === "true") return true;
  if (normalized === "0" || normalized === "false") return false;
  // DrawingML sometimes uses "on"/"off".
  if (normalized === "on") return true;
  if (normalized === "off") return false;
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
    const scheme = findFirstByLocalName(solidFill, "schemeClr");
    color = normalizeHexColor(srgb?.getAttribute("val")) ?? (scheme?.getAttribute("val")?.trim() ? undefined : undefined);
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

  const shapeDefaultStyle = parseRunStyleFromDom(findFirstByLocalName(txBody, "defRPr"));

  const paragraphs = Array.from(txBody.childNodes).filter(
    (n): n is Element => n.nodeType === 1 && (n as Element).localName === "p",
  );

  const textRuns: ShapeTextRun[] = [];
  let alignment: ShapeTextLayout["alignment"] | undefined;

  for (let pi = 0; pi < paragraphs.length; pi += 1) {
    const p = paragraphs[pi];
    const pPr = Array.from(p.childNodes).find((n): n is Element => n.nodeType === 1 && (n as Element).localName === "pPr") ?? null;
    if (!alignment) {
      const pPrAlign = pPr?.getAttribute("algn") ?? null;
      alignment = parseAlignment(pPrAlign);
      if (!alignment && pPr) {
        const algnChild = findFirstByLocalName(pPr, "algn");
        alignment = parseAlignment(algnChild?.getAttribute("val") ?? null);
      }
    }

    const paraDefaultStyle = mergeStyle(shapeDefaultStyle, parseRunStyleFromDom(findFirstByLocalName(pPr ?? p, "defRPr")));

    for (const child of Array.from(p.childNodes)) {
      if (child.nodeType !== 1) continue;
      const el = child as Element;

      if (el.localName === "br") {
        textRuns.push({ text: "\n", ...paraDefaultStyle });
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

  return { textRuns, alignment, vertical, wrap };
}

function decodeXmlEntities(text: string): string {
  return String(text)
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&")
    .replace(/&#(\d+);/g, (_, code) => String.fromCharCode(Number(code)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, code) => String.fromCharCode(Number.parseInt(code, 16)));
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

function parseShapeTextFallback(rawXml: string): ShapeTextLayout | null {
  const txBodyMatch = /<(?:[A-Za-z_][\w.-]*:)?txBody\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?txBody\s*>/i.exec(rawXml);
  if (!txBodyMatch) return { textRuns: [] };
  const txBodyXml = txBodyMatch[0];
  const txBodyInner = txBodyMatch[1] ?? "";

  const bodyPrXml = extractFirstElementXml(txBodyInner, "bodyPr");
  const bodyPrOpen = bodyPrXml ? /<[^>]+>/.exec(bodyPrXml)?.[0] ?? bodyPrXml : "";
  const vertical = parseVertical(getAttrFromTag(bodyPrOpen, "anchor"));
  const wrap = parseWrap(getAttrFromTag(bodyPrOpen, "wrap"));

  const shapeDefaultStyle = parseRunStyleFromXmlSnippet(extractFirstElementXml(txBodyXml, "defRPr"));

  const paragraphsInner = extractAllElementInner(txBodyInner, "p");
  const textRuns: ShapeTextRun[] = [];
  let alignment: ShapeTextLayout["alignment"] | undefined;

  for (let pi = 0; pi < paragraphsInner.length; pi += 1) {
    const pInner = paragraphsInner[pi] ?? "";

    const pPrXml = extractFirstElementXml(pInner, "pPr");
    const pPrOpen = pPrXml ? /<[^>]+>/.exec(pPrXml)?.[0] ?? pPrXml : null;
    if (!alignment) {
      alignment = parseAlignment(getAttrFromTag(pPrOpen ?? "", "algn"));
      if (!alignment && pPrXml) {
        const algnXml = extractFirstElementXml(pPrXml, "algn");
        const algnOpen = algnXml ? /<[^>]+>/.exec(algnXml)?.[0] ?? algnXml : "";
        alignment = parseAlignment(getAttrFromTag(algnOpen, "val"));
      }
    }

    const paraDefaultStyle = mergeStyle(shapeDefaultStyle, parseRunStyleFromXmlSnippet(extractFirstElementXml(pPrXml ?? "", "defRPr")));

    // Walk runs + breaks in order within the paragraph.
    const nodeRe =
      /<(?:[A-Za-z_][\w.-]*:)?(r|fld)\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?\1\s*>|<(?:[A-Za-z_][\w.-]*:)?br\b[^>]*\/>|<(?:[A-Za-z_][\w.-]*:)?br\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?br\s*>/gi;
    for (const match of pInner.matchAll(nodeRe)) {
      const chunk = match[0] ?? "";
      if (/(?:^|<)[^>]*?\bbr\b/i.test(chunk)) {
        textRuns.push({ text: "\n", ...paraDefaultStyle });
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

  return { textRuns, alignment, vertical, wrap };
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
