function hasOwn(obj: unknown, key: string): boolean {
  return Boolean(obj) && Object.prototype.hasOwnProperty.call(obj as object, key);
}

function readPreferredKey(obj: unknown, camelKey: string, snakeKey: string): unknown {
  if (hasOwn(obj, camelKey)) return (obj as any)[camelKey];
  if (hasOwn(obj, snakeKey)) return (obj as any)[snakeKey];
  return undefined;
}

export function getStyleNumberFormat(style: unknown): string | null {
  const raw = readPreferredKey(style, "numberFormat", "number_format");
  if (typeof raw !== "string") return null;
  const trimmed = raw.trim();
  if (trimmed === "") return null;
  // Treat "General" (Excel default) as equivalent to clearing the number format.
  if (trimmed.toLowerCase() === "general") return null;
  return raw;
}

export function getStyleWrapText(style: unknown): boolean {
  const alignment = (style as any)?.alignment;
  const raw = readPreferredKey(alignment, "wrapText", "wrap_text");
  return raw === true;
}

export function getStyleFillFgColor(style: unknown): unknown {
  const fill = (style as any)?.fill;

  // Excel / OOXML-ish styles usually express fill using `fgColor`, but formula-model uses
  // snake_case, and some older encodings store the primary fill under `background`.
  if (hasOwn(fill, "fgColor")) return (fill as any).fgColor;
  if (hasOwn(fill, "fg_color")) return (fill as any).fg_color;
  if (hasOwn(fill, "background")) return (fill as any).background;
  if (hasOwn(fill, "bgColor")) return (fill as any).bgColor;
  if (hasOwn(fill, "bg_color")) return (fill as any).bg_color;

  return undefined;
}

export function getStyleFillBgColor(style: unknown): unknown {
  const fill = (style as any)?.fill;
  if (hasOwn(fill, "bgColor")) return (fill as any).bgColor;
  if (hasOwn(fill, "bg_color")) return (fill as any).bg_color;
  if (hasOwn(fill, "background")) return (fill as any).background;
  return undefined;
}

export function getStyleFontSizePt(style: unknown): number | null {
  const font = (style as any)?.font;

  // Prefer `font.size` (pt) when present so user edits override imported `size_100pt`.
  if (hasOwn(font, "size")) {
    const raw = (font as any).size;
    return typeof raw === "number" && Number.isFinite(raw) ? raw : null;
  }

  if (hasOwn(font, "size_100pt")) {
    const raw = (font as any).size_100pt;
    if (typeof raw !== "number" || !Number.isFinite(raw)) return null;
    // formula-model / XLSX import serializes font sizes in 1/100th of a point.
    return raw / 100;
  }

  // Backward-compatible fallbacks (pre-style-table OOXML-ish shims).
  const legacy = readPreferredKey(style, "fontSize", "font_size");
  return typeof legacy === "number" && Number.isFinite(legacy) ? legacy : null;
}
