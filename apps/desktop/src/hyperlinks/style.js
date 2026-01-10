// Prefer a design token instead of hardcoding hex colors; downstream themes can
// map `--accent` appropriately.
export const DEFAULT_HYPERLINK_COLOR = "var(--accent)";

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
