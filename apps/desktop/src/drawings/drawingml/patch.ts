/**
 * Best-effort DrawingML patch helpers.
 *
 * We intentionally avoid full XML parsing here (to keep the desktop renderer
 * Node/DOM free and lightweight) and instead do tolerant-but-guarded string
 * patching:
 * - Match elements by local-name (optional namespace prefix).
 * - Only rewrite specific attributes on specific elements.
 * - No-op when the target element/attribute isn't present.
 *
 * These helpers are used to keep preserved DrawingML fragments (`raw_xml`,
 * `xlsx.pic_xml`, etc.) in sync when UI edits anchors/sizes or duplicates
 * objects. This maintains round-trip fidelity for XLSX export.
 */

function formatEmu(n: number): string {
  // DrawingML EMU values are integers. UI code may compute floats; normalize.
  return String(Math.round(n));
}

function patchAttr(tag: string, attrName: string, attrValue: string): string {
  // Replace only when the attribute already exists.
  // Handles both single and double quoted values.
  // Guardrail: require the attribute name to be preceded by whitespace (or be at
  // the very start of the matched string). This avoids accidentally patching
  // namespaced attributes like `r:id="..."` or namespace declarations like
  // `xmlns:x="..."` when targeting `id`/`x`/etc.
  const re = new RegExp(`(^|\\s)${attrName}=(['"])([^'"]*)\\2`);
  if (!re.test(tag)) return tag;
  return tag.replace(re, `$1${attrName}=$2${attrValue}$2`);
}

/**
 * Updates `<*:cNvPr id="…">` IDs inside an object fragment.
 */
export function patchNvPrId(xml: string, newId: number): string {
  const id = String(Math.trunc(newId));
  // Match a start tag for an element whose local name is `cNvPr`.
  // Example matches:
  //   <xdr:cNvPr id="2" name="Picture 1"/>
  //   <cNvPr id='2' />
  const cNvPrTagRe = /<(?:[A-Za-z_][\w.-]*:)?cNvPr\b[^>]*>/g;
  let changed = false;
  const out = xml.replace(cNvPrTagRe, (tag) => {
    const patched = patchAttr(tag, "id", id);
    if (patched !== tag) changed = true;
    return patched;
  });
  return changed ? out : xml;
}

function patchFirstInXfrm(xml: string, localName: "ext" | "off", patch: (tag: string) => string): string {
  // Find the first `<*:xfrm>…</*:xfrm>` element; we avoid trying to patch every
  // xfrm in the fragment because some fragments (e.g. group shapes) may contain
  // multiple transforms.
  const xfrmRe = /<(?:[A-Za-z_][\w.-]*:)?xfrm\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?xfrm>/;
  const m = xfrmRe.exec(xml);
  if (!m) return xml;

  const xfrmXml = m[0];
  const tagRe = new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${localName}\\b[^>]*\\/?>`);
  const tagMatch = tagRe.exec(xfrmXml);
  if (!tagMatch) return xml;

  const tag = tagMatch[0];
  const patchedTag = patch(tag);
  if (patchedTag === tag) return xml;

  const patchedXfrm = xfrmXml.replace(tag, patchedTag);
  return xml.slice(0, m.index) + patchedXfrm + xml.slice(m.index + xfrmXml.length);
}

/**
 * Updates `<a:ext cx="…" cy="…"/>` under the first `<*:xfrm>` found in the
 * fragment. No-op when the relevant node/attrs are missing.
 */
export function patchXfrmExt(xml: string, cxEmu: number, cyEmu: number): string {
  const cx = formatEmu(cxEmu);
  const cy = formatEmu(cyEmu);
  return patchFirstInXfrm(xml, "ext", (tag) => {
    let out = tag;
    out = patchAttr(out, "cx", cx);
    out = patchAttr(out, "cy", cy);
    return out;
  });
}

/**
 * Updates `<a:off x="…" y="…"/>` under the first `<*:xfrm>` found in the
 * fragment. No-op when the relevant node/attrs are missing.
 */
export function patchXfrmOff(xml: string, xEmu: number, yEmu: number): string {
  const x = formatEmu(xEmu);
  const y = formatEmu(yEmu);
  return patchFirstInXfrm(xml, "off", (tag) => {
    let out = tag;
    out = patchAttr(out, "x", x);
    out = patchAttr(out, "y", y);
    return out;
  });
}

export function extractXfrmOff(xml: string): { xEmu: number; yEmu: number } | null {
  const xfrmRe = /<(?:[A-Za-z_][\w.-]*:)?xfrm\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?xfrm>/;
  const m = xfrmRe.exec(xml);
  if (!m) return null;
  const xfrmXml = m[0];
  const offRe = /<(?:[A-Za-z_][\w.-]*:)?off\b[^>]*\/?>/;
  const offMatch = offRe.exec(xfrmXml);
  if (!offMatch) return null;
  const tag = offMatch[0];
  // Guardrails: avoid matching namespaced attributes like `r:x="..."`.
  const xMatch = /(^|\s)x=(['"])(-?\d+)\2/.exec(tag);
  const yMatch = /(^|\s)y=(['"])(-?\d+)\2/.exec(tag);
  if (!xMatch || !yMatch) return null;
  const xEmu = Number.parseInt(xMatch[3]!, 10);
  const yEmu = Number.parseInt(yMatch[3]!, 10);
  if (!Number.isFinite(xEmu) || !Number.isFinite(yEmu)) return null;
  return { xEmu, yEmu };
}

function patchFirstBlock(xml: string, localName: string, patch: (block: string) => string): string {
  const blockRe = new RegExp(
    `<(?:[A-Za-z_][\\w.-]*:)?${localName}\\b[^>]*>[\\s\\S]*?<\\/(?:[A-Za-z_][\\w.-]*:)?${localName}>`,
  );
  const m = blockRe.exec(xml);
  if (!m) return xml;
  const block = m[0];
  const patched = patch(block);
  if (patched === block) return xml;
  return xml.slice(0, m.index) + patched + xml.slice(m.index + block.length);
}

function patchFirstChildTextInt(parentXml: string, childLocalName: string, value: number): string {
  const n = String(Math.trunc(value));
  const re = new RegExp(
    `(<(?:[A-Za-z_][\\w.-]*:)?${childLocalName}\\b[^>]*>\\s*)(-?\\d+)(\\s*<\\/(?:[A-Za-z_][\\w.-]*:)?${childLocalName}>)`,
  );
  if (!re.test(parentXml)) return parentXml;
  return parentXml.replace(re, `$1${n}$3`);
}

function patchFirstChildTextEmu(parentXml: string, childLocalName: string, value: number): string {
  const n = formatEmu(value);
  const re = new RegExp(
    `(<(?:[A-Za-z_][\\w.-]*:)?${childLocalName}\\b[^>]*>\\s*)(-?\\d+)(\\s*<\\/(?:[A-Za-z_][\\w.-]*:)?${childLocalName}>)`,
  );
  if (!re.test(parentXml)) return parentXml;
  return parentXml.replace(re, `$1${n}$3`);
}

/**
 * Patch an anchor `<from>` / `<to>` block inside a full anchor subtree (e.g. when
 * preserving an unknown DrawingML anchor verbatim).
 *
 * No-op when the target block or expected children are not present.
 */
export function patchAnchorPoint(
  xml: string,
  which: "from" | "to",
  point: { col: number; row: number; colOffEmu: number; rowOffEmu: number },
): string {
  return patchFirstBlock(xml, which, (block) => {
    let out = block;
    out = patchFirstChildTextInt(out, "col", point.col);
    out = patchFirstChildTextEmu(out, "colOff", point.colOffEmu);
    out = patchFirstChildTextInt(out, "row", point.row);
    out = patchFirstChildTextEmu(out, "rowOff", point.rowOffEmu);
    return out;
  });
}

/**
 * Patch the first `<*:pos x="…" y="…"/>` element. Intended for absoluteAnchor
 * payloads where the full anchor subtree is preserved.
 */
export function patchAnchorPos(xml: string, xEmu: number, yEmu: number): string {
  const posRe = /<(?:[A-Za-z_][\w.-]*:)?pos\b[^>]*\/?>/;
  const m = posRe.exec(xml);
  if (!m) return xml;
  const tag = m[0];
  let patched = tag;
  patched = patchAttr(patched, "x", formatEmu(xEmu));
  patched = patchAttr(patched, "y", formatEmu(yEmu));
  if (patched === tag) return xml;
  return xml.slice(0, m.index) + patched + xml.slice(m.index + tag.length);
}

/**
 * Patch the first `<*:ext cx="…" cy="…"/>` element. Intended for oneCellAnchor /
 * absoluteAnchor payloads where the full anchor subtree is preserved.
 */
export function patchAnchorExt(xml: string, cxEmu: number, cyEmu: number): string {
  const extRe = /<(?:[A-Za-z_][\w.-]*:)?ext\b[^>]*\/?>/;
  const m = extRe.exec(xml);
  if (!m) return xml;
  const tag = m[0];
  let patched = tag;
  patched = patchAttr(patched, "cx", formatEmu(cxEmu));
  patched = patchAttr(patched, "cy", formatEmu(cyEmu));
  if (patched === tag) return xml;
  return xml.slice(0, m.index) + patched + xml.slice(m.index + tag.length);
}
