/**
 * Default font families used by {@link CanvasGridRenderer}.
 *
 * The shared grid renderer is used across desktop and web surfaces. Keep the
 * renderer defaults aligned with "system UI" typography, and allow host apps
 * to opt into monospace cell rendering via `CanvasGridRenderer` constructor
 * options (`defaultCellFontFamily` / `defaultHeaderFontFamily`).
 */

/**
 * Default font family for UI chrome (and as the grid renderer's baseline when
 * no per-cell style overrides are provided).
 */
export const DEFAULT_GRID_FONT_FAMILY = "system-ui";

/**
 * Monospace stack used by the desktop spreadsheet UI for cell content.
 *
 * Keep this list conservative (no quoted families) so it can be embedded into
 * `CanvasRenderingContext2D.font` strings without additional escaping.
 */
export const DEFAULT_GRID_MONOSPACE_FONT_FAMILY = "ui-monospace, SFMono-Regular, Menlo, Consolas, monospace";

