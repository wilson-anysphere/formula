export type PresenceRendererOptions = {
  selectionFillAlpha?: number;
  selectionStrokeAlpha?: number;
  cursorStrokeWidth?: number;
  badgePaddingX?: number;
  badgePaddingY?: number;
  badgeOffsetX?: number;
  badgeOffsetY?: number;
  font?: string;
};

export class PresenceRenderer {
  [key: string]: any;
  constructor(options?: PresenceRendererOptions);
  clear(ctx: CanvasRenderingContext2D): void;
  render(ctx: CanvasRenderingContext2D, presences: any[], options: { getCellRect: (...args: any[]) => any }): void;
}

