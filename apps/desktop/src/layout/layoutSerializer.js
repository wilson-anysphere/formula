import { createDefaultLayout } from "./layoutState.js";
import { normalizeLayout } from "./layoutNormalization.js";

/**
 * @param {unknown} layout
 * @param {{ panelRegistry?: Record<string, unknown>, primarySheetId?: string | null }} [options]
 */
export function serializeLayout(layout, options = {}) {
  return JSON.stringify(normalizeLayout(layout, options));
}

/**
 * @param {string} json
 * @param {{ panelRegistry?: Record<string, unknown>, primarySheetId?: string | null }} [options]
 */
export function deserializeLayout(json, options = {}) {
  try {
    const parsed = JSON.parse(json);
    return normalizeLayout(parsed, options);
  } catch {
    return createDefaultLayout({ primarySheetId: options.primarySheetId });
  }
}

