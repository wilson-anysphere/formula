export const LAYOUT_STATE_VERSION = 1;

export const DOCK_SIDES = /** @type {const} */ (["left", "right", "bottom"]);
export const SPLIT_DIRECTIONS = /** @type {const} */ (["none", "vertical", "horizontal"]);
export const SPLIT_PANES = /** @type {const} */ (["primary", "secondary"]);

export const DEFAULT_DOCK_SIZES = Object.freeze({
  left: 320,
  right: 360,
  bottom: 240,
});

export const DEFAULT_FLOATING_RECT = Object.freeze({
  x: 80,
  y: 80,
  width: 420,
  height: 560,
});

export const DEFAULT_SPLIT_RATIO = 0.5;
export const DEFAULT_ACTIVE_SPLIT_PANE = "primary";
