import type { DrawingObject } from "./types";
import { patchNvPrId } from "./drawingml/patch";

function patchNvPrNameSuffix(xml: string, newId: number): string {
  const idStr = String(Math.trunc(newId));
  const cNvPrTagRe = /<(?:[A-Za-z_][\w.-]*:)?cNvPr\b[^>]*>/g;
  let changed = false;
  const out = xml.replace(cNvPrTagRe, (tag) => {
    const nameRe = /(^|\s)name=(['"])([^'"]*)\2/;
    const m = nameRe.exec(tag);
    if (!m) return tag;
    const name = m[3] ?? "";
    const next = name.replace(/\d+$/, idStr);
    if (next === name) return tag;
    changed = true;
    return tag.replace(nameRe, `$1name=$2${next}$2`);
  });
  return changed ? out : xml;
}

function patchInnerXmlForNewId(xml: string, newId: number): string {
  let out = patchNvPrId(xml, newId);
  out = patchNvPrNameSuffix(out, newId);
  return out;
}

/**
 * Duplicates a drawing object, assigning a new `id` and keeping any preserved
 * DrawingML fragments consistent (e.g. `<xdr:cNvPr id="…">`).
 */
export function duplicateDrawingObject<T extends DrawingObject>(object: T, newId: number): T {
  let next: DrawingObject = { ...object, id: newId };

  if (object.kind.type === "image") {
    const picXml = object.preserved?.["xlsx.pic_xml"];
    if (typeof picXml === "string") {
      const patched = patchInnerXmlForNewId(picXml, newId);
      if (patched !== picXml) {
        next = {
          ...next,
          preserved: {
            ...(object.preserved ?? {}),
            "xlsx.pic_xml": patched,
          },
        };
      }
    }
    return next as T;
  }

  const kindAny = object.kind as any;
  const rawXml: unknown = kindAny.rawXml ?? kindAny.raw_xml;
  if (typeof rawXml === "string") {
    const patched = patchInnerXmlForNewId(rawXml, newId);
    if (patched !== rawXml) {
      next = {
        ...next,
        kind: { ...kindAny, rawXml: patched, raw_xml: patched },
      };
    }
  }

  return next as T;
}
