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
export const DEFAULT_GRID_FONT_FAMILY = '-apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif';

/**
 * Monospace stack used by the desktop spreadsheet UI for cell content.
 */
export const DEFAULT_GRID_MONOSPACE_FONT_FAMILY = 'ui-monospace, SFMono-Regular, "SF Mono", Menlo, Consolas, monospace';
