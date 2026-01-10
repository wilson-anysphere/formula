/**
 * @typedef {Object} RichTextRunStyle
 * @property {boolean=} bold
 * @property {boolean=} italic
 * @property {'single' | 'double' | 'single_accounting' | 'double_accounting' | 'none'=} underline
 * @property {string=} color Engine color string in `#AARRGGBB` form (alpha-first).
 * @property {string=} font Font family name
 * @property {number=} size_100pt Font size in 1/100 points (e.g. 1100 = 11pt)
 */

/**
 * @typedef {Object} RichTextRun
 * @property {number} start Inclusive start offset (Unicode code point index)
 * @property {number} end Exclusive end offset (Unicode code point index)
 * @property {RichTextRunStyle=} style
 */

/**
 * @typedef {Object} RichText
 * @property {string} text
 * @property {RichTextRun[]=} runs
 */

export {};
