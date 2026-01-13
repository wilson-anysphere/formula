/**
 * Helpers for parsing / classifying DrawingML payloads stored on drawing objects.
 *
 * The drawing overlay renders a small subset of shapes (rect/ellipse/line/etc).
 * We keep parsing lightweight and cache results since these helpers can be
 * called frequently while rendering.
 */

/**
 * Returns true when the raw DrawingML payload contains a `graphicFrame` element.
 *
 * In SpreadsheetDrawing, `xdr:graphicFrame` is used for a variety of embedded
 * objects (charts, SmartArt/diagrams, etc).
 */
export function isGraphicFrame(rawXml: string | null | undefined): boolean {
  if (!rawXml) return false;
  // Match `<xdr:graphicFrame ...>` as well as other namespace prefixes.
  return /<\s*(?:[A-Za-z0-9_-]+:)?graphicFrame\b/.test(rawXml);
}

/**
 * Best-effort classification for diagram-based graphic frames (SmartArt).
 */
export function isSmartArtGraphicFrame(rawXml: string | null | undefined): boolean {
  if (!rawXml) return false;
  if (!isGraphicFrame(rawXml)) return false;
  // SmartArt uses `a:graphicData uri=".../diagram"` and/or `dgm:*` elements.
  return rawXml.includes("drawingml/2006/diagram") || /<\s*dgm:/.test(rawXml);
}

/**
 * Placeholder label for unsupported graphic frames.
 */
export function graphicFramePlaceholderLabel(rawXml: string | null | undefined): "SmartArt" | "GraphicFrame" | null {
  if (!rawXml) return null;
  if (!isGraphicFrame(rawXml)) return null;
  return isSmartArtGraphicFrame(rawXml) ? "SmartArt" : "GraphicFrame";
}

export type ShapePresetGeometry = "rect" | "roundRect" | "ellipse" | "line";

export type ShapeFill = { type: "none" } | { type: "solid"; color: string };

export interface ShapeStroke {
  color: string;
  /** DrawingML EMU line width (1pt = 12700). */
  widthEmu: number;
}

export interface ShapeRenderSpec {
  geometry: { type: ShapePresetGeometry };
  fill: ShapeFill;
  stroke?: ShapeStroke;
  /** Best-effort first line of text from `<xdr:txBody>`. */
  label?: string;
}

const DRAWINGML_NAMESPACES = {
  xdr: "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
};

const WRAP_PREFIX = `<root xmlns:xdr="${DRAWINGML_NAMESPACES.xdr}" xmlns:a="${DRAWINGML_NAMESPACES.a}" xmlns:r="${DRAWINGML_NAMESPACES.r}">`;
const WRAP_SUFFIX = "</root>";

const DEFAULT_LINE_WIDTH_EMU = 9_525; // 1px at 96DPI (914400 / 96)

const SHAPE_SPEC_CACHE = new Map<string, ShapeRenderSpec | null>();
const SHAPE_SPEC_CACHE_MAX = 200;

type SimpleXmlNode = SimpleXmlElement | SimpleXmlText;

interface SimpleXmlElement {
  kind: "element";
  name: string;
  attributes: Record<string, string>;
  children: SimpleXmlNode[];
}

interface SimpleXmlText {
  kind: "text";
  text: string;
}

type XmlElementLike = Element | SimpleXmlElement;

function isDomElement(node: XmlElementLike): node is Element {
  return typeof (node as any).getAttribute === "function";
}

function localName(node: XmlElementLike): string {
  if (isDomElement(node)) {
    // DOMParser in XML mode sets `localName`.
    return node.localName ?? node.tagName;
  }
  const idx = node.name.indexOf(":");
  return idx >= 0 ? node.name.slice(idx + 1) : node.name;
}

function getAttribute(node: XmlElementLike, name: string): string | null {
  if (isDomElement(node)) return node.getAttribute(name);
  return node.attributes[name] ?? null;
}

function childElements(node: XmlElementLike): XmlElementLike[] {
  if (isDomElement(node)) {
    const out: Element[] = [];
    for (const child of Array.from(node.childNodes)) {
      if (child.nodeType === 1) out.push(child as Element);
    }
    return out;
  }
  return node.children.filter((c): c is SimpleXmlElement => c.kind === "element");
}

function textContent(node: XmlElementLike): string {
  if (isDomElement(node)) return node.textContent ?? "";
  let out = "";
  const stack: SimpleXmlNode[] = [...node.children];
  while (stack.length) {
    const next = stack.shift()!;
    if (next.kind === "text") out += next.text;
    else stack.unshift(...next.children);
  }
  return out;
}

function findFirstDescendantByLocalName(root: XmlElementLike, name: string): XmlElementLike | null {
  const queue: XmlElementLike[] = [root];
  while (queue.length) {
    const node = queue.shift()!;
    if (localName(node) === name) return node;
    queue.unshift(...childElements(node));
  }
  return null;
}

function wrapRawXml(rawXml: string): string {
  // Raw DrawingML snippets frequently omit namespace declarations because they're
  // inherited from the drawing root (`xdr:wsDr`). Wrap the fragment with the
  // required namespace bindings so `DOMParser` can parse it.
  return `${WRAP_PREFIX}${rawXml}${WRAP_SUFFIX}`;
}

function parseWithDomParser(xml: string): Element | null {
  const DOMParserCtor = (globalThis as any).DOMParser as typeof DOMParser | undefined;
  if (typeof DOMParserCtor !== "function") return null;
  try {
    const parser = new DOMParserCtor();
    const doc = parser.parseFromString(xml, "application/xml");
    const err = doc.getElementsByTagName("parsererror")[0];
    if (err) return null;
    return doc.documentElement;
  } catch {
    return null;
  }
}

function decodeEntities(text: string): string {
  return text
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function parseAttributes(text: string): Record<string, string> {
  const attrs: Record<string, string> = {};
  const re = /([^\s=/>]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'>/]+)))?/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text))) {
    const key = m[1]!;
    const value = m[2] ?? m[3] ?? m[4] ?? "";
    attrs[key] = decodeEntities(value);
  }
  return attrs;
}

function parseWithFallback(xml: string): SimpleXmlElement | null {
  const len = xml.length;
  let i = 0;
  const stack: SimpleXmlElement[] = [];
  let root: SimpleXmlElement | null = null;

  const pushText = (text: string) => {
    if (!stack.length) return;
    const decoded = decodeEntities(text);
    stack[stack.length - 1]!.children.push({ kind: "text", text: decoded });
  };

  while (i < len) {
    const lt = xml.indexOf("<", i);
    if (lt === -1) {
      pushText(xml.slice(i));
      break;
    }

    if (lt > i) {
      pushText(xml.slice(i, lt));
    }

    // Handle special constructs.
    if (xml.startsWith("<!--", lt)) {
      const end = xml.indexOf("-->", lt + 4);
      i = end === -1 ? len : end + 3;
      continue;
    }
    if (xml.startsWith("<?", lt)) {
      const end = xml.indexOf("?>", lt + 2);
      i = end === -1 ? len : end + 2;
      continue;
    }
    if (xml.startsWith("<![CDATA[", lt)) {
      const end = xml.indexOf("]]>", lt + 9);
      const content = end === -1 ? xml.slice(lt + 9) : xml.slice(lt + 9, end);
      pushText(content);
      i = end === -1 ? len : end + 3;
      continue;
    }

    const gt = xml.indexOf(">", lt + 1);
    if (gt === -1) break;
    const rawTag = xml.slice(lt + 1, gt);

    // End tag.
    if (rawTag.startsWith("/")) {
      if (stack.length) stack.pop();
      i = gt + 1;
      continue;
    }

    const selfClosing = rawTag.endsWith("/");
    const tagContent = selfClosing ? rawTag.slice(0, -1) : rawTag;
    const match = /^\s*([^\s/>]+)([\s\S]*)$/.exec(tagContent);
    if (!match) {
      i = gt + 1;
      continue;
    }

    const name = match[1]!;
    const attrsText = match[2] ?? "";
    const el: SimpleXmlElement = {
      kind: "element",
      name,
      attributes: parseAttributes(attrsText),
      children: [],
    };

    if (stack.length) stack[stack.length - 1]!.children.push(el);
    else root = el;

    if (!selfClosing) stack.push(el);
    i = gt + 1;
  }

  return root;
}

function parseXmlRoot(rawXml: string): XmlElementLike | null {
  const wrapped = wrapRawXml(rawXml);
  const dom = parseWithDomParser(wrapped);
  if (dom) return dom;
  return parseWithFallback(wrapped);
}

function normalizePresetGeometry(prst: string): ShapePresetGeometry | null {
  switch (prst) {
    case "rect":
    case "roundRect":
    case "ellipse":
    case "line":
      return prst;
    default:
      return null;
  }
}

function parseSrgbColor(node: XmlElementLike): string | null {
  const srgb = findFirstDescendantByLocalName(node, "srgbClr");
  if (!srgb) return null;
  const val = getAttribute(srgb, "val");
  if (!val) return null;
  const hex = val.trim();
  if (!/^[0-9a-fA-F]{6}$/.test(hex)) return null;
  return `#${hex.toUpperCase()}`;
}

function parseFill(spPr: XmlElementLike): ShapeFill {
  const children = childElements(spPr);
  if (children.some((c) => localName(c) === "noFill")) return { type: "none" };
  const solid = children.find((c) => localName(c) === "solidFill");
  const color = solid ? parseSrgbColor(solid) : null;
  if (color) return { type: "solid", color };
  return { type: "none" };
}

function parseStroke(spPr: XmlElementLike): ShapeStroke | undefined {
  const ln = childElements(spPr).find((c) => localName(c) === "ln");
  // Shapes default to having a 1px black outline if not specified. This keeps
  // minimally-specified shapes visible and matches typical Excel defaults.
  if (!ln) return { color: "#000000", widthEmu: DEFAULT_LINE_WIDTH_EMU };
  if (childElements(ln).some((c) => localName(c) === "noFill")) return undefined;

  const widthAttr = getAttribute(ln, "w");
  let widthEmu = DEFAULT_LINE_WIDTH_EMU;
  if (widthAttr != null) {
    const parsedWidth = Number.parseInt(widthAttr, 10);
    widthEmu = Number.isFinite(parsedWidth) && parsedWidth >= 0 ? parsedWidth : DEFAULT_LINE_WIDTH_EMU;
  }
  if (widthEmu === 0) return undefined;

  const solid = childElements(ln).find((c) => localName(c) === "solidFill");
  const color = (solid ? parseSrgbColor(solid) : null) ?? "#000000";
  return { color, widthEmu };
}

function parseLabel(root: XmlElementLike): string | undefined {
  const txBody = findFirstDescendantByLocalName(root, "txBody");
  if (!txBody) return undefined;
  const p = findFirstDescendantByLocalName(txBody, "p");
  if (!p) return undefined;

  // Collect `<a:t>` values inside the first paragraph. Stop on explicit `<a:br>`.
  const parts: string[] = [];
  const queue: XmlElementLike[] = [p];
  while (queue.length) {
    const node = queue.shift()!;
    const name = localName(node);
    if (name === "br") break;
    if (name === "t") {
      const value = textContent(node);
      if (value) parts.push(value);
    }
    queue.unshift(...childElements(node));
  }
  const text = parts.join("").trim();
  if (!text) return undefined;
  return text.split(/\r?\n/, 1)[0]!.trim();
}

function cacheResult(rawXml: string, spec: ShapeRenderSpec | null): ShapeRenderSpec | null {
  // Keep the cache bounded (simple FIFO eviction is sufficient here).
  if (SHAPE_SPEC_CACHE.size >= SHAPE_SPEC_CACHE_MAX) {
    const firstKey = SHAPE_SPEC_CACHE.keys().next().value as string | undefined;
    if (firstKey) SHAPE_SPEC_CACHE.delete(firstKey);
  }
  SHAPE_SPEC_CACHE.set(rawXml, spec);
  return spec;
}

/**
 * Best-effort parsing of a DrawingML `<xdr:sp>` payload into a canvas-friendly render spec.
 *
 * Returns `null` for unsupported or unknown shapes so callers can fall back to
 * placeholder rendering.
 */
export function parseShapeRenderSpec(rawXml: string): ShapeRenderSpec | null {
  if (typeof rawXml !== "string" || rawXml.trim().length === 0) return null;
  if (SHAPE_SPEC_CACHE.has(rawXml)) return SHAPE_SPEC_CACHE.get(rawXml)!;

  const root = parseXmlRoot(rawXml);
  if (!root) return cacheResult(rawXml, null);

  const spPr = findFirstDescendantByLocalName(root, "spPr");
  if (!spPr) return cacheResult(rawXml, null);

  const prstGeom = findFirstDescendantByLocalName(spPr, "prstGeom");
  const prst = prstGeom ? getAttribute(prstGeom, "prst") : null;
  if (!prst) return cacheResult(rawXml, null);

  const geometry = normalizePresetGeometry(prst);
  if (!geometry) return cacheResult(rawXml, null);

  return cacheResult(rawXml, {
    geometry: { type: geometry },
    fill: parseFill(spPr),
    stroke: parseStroke(spPr),
    label: parseLabel(root),
  });
}

