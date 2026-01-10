/**
 * @typedef {Object} RichTextRunStyle
 * @property {boolean=} bold
 * @property {boolean=} italic
 * @property {boolean | 'single' | 'double' | 'singleAccounting' | 'doubleAccounting' | 'none'=} underline
 * @property {string=} color CSS color string (`#RRGGBB`, `rgba(...)`, etc)
 * @property {string=} font Font family name
 * @property {number=} size Font size in pixels
 */

/**
 * @typedef {Object} RichTextRun
 * @property {number} start Inclusive start offset (JS string index)
 * @property {number} end Exclusive end offset (JS string index)
 * @property {RichTextRunStyle=} style
 */

/**
 * @typedef {Object} RichText
 * @property {string} text
 * @property {RichTextRun[]=} runs
 */

export {};

