export const SR_ONLY_STYLE: Record<string, string> = {
  position: "absolute",
  width: "1px",
  height: "1px",
  padding: "0",
  margin: "-1px",
  overflow: "hidden",
  clip: "rect(0, 0, 0, 0)",
  whiteSpace: "nowrap",
  border: "0"
};

Object.freeze(SR_ONLY_STYLE);

export function applySrOnlyStyle(el: HTMLElement): void {
  for (const [key, value] of Object.entries(SR_ONLY_STYLE)) {
    // `CSSStyleDeclaration` is not indexable by arbitrary string keys in TS, but
    // these are all valid inline style properties.
    (el.style as unknown as Record<string, string>)[key] = value;
  }
}
