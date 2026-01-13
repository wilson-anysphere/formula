// Prefer a dedicated design token instead of hardcoding hex colors so hyperlinks
// remain legible across light/dark/high-contrast themes.
export const DEFAULT_HYPERLINK_COLOR = "var(--link)";

/**
 * Return a CSS style object for rendering hyperlink text.
 *
 * Consumers (canvas renderer, cell editor, etc.) can map this onto their
 * rendering primitives.
 */
export function hyperlinkTextStyle() {
  return {
    color: DEFAULT_HYPERLINK_COLOR,
    textDecoration: "underline",
  };
}
