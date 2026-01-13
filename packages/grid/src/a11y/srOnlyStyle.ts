export const SR_ONLY_STYLE = Object.freeze({
  position: "absolute",
  width: "1px",
  height: "1px",
  padding: "0px",
  margin: "-1px",
  overflow: "hidden",
  clip: "rect(0px, 0px, 0px, 0px)",
  whiteSpace: "nowrap",
  border: "0px"
} as const);

export function applySrOnlyStyle(el: HTMLElement): void {
  for (const [key, value] of Object.entries(SR_ONLY_STYLE)) {
    // `CSSStyleDeclaration` is not indexable by arbitrary string keys in TS, but
    // these are all valid inline style properties.
    (el.style as unknown as Record<string, string>)[key] = value;
  }
}
