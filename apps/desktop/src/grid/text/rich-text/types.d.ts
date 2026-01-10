export type Underline =
  | "single"
  | "double"
  | "single_accounting"
  | "double_accounting"
  | "none";

export interface RichTextRunStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: Underline;
  /**
   * Engine color string in `#AARRGGBB` form (alpha-first).
   */
  color?: string;
  font?: string;
  /**
   * Font size in 1/100 points (e.g. 1100 = 11pt).
   */
  size_100pt?: number;
}

export interface RichTextRun {
  /**
   * Inclusive start offset (Unicode code point index).
   */
  start: number;
  /**
   * Exclusive end offset (Unicode code point index).
   */
  end: number;
  style?: RichTextRunStyle;
}

export interface RichText {
  text: string;
  runs?: RichTextRun[];
}

