export interface Rect {
  x: number;
  y: number;
  width: number;
  height: number;
}

function normalizeRect(rect: Rect): Rect | null {
  if (!Number.isFinite(rect.x) || !Number.isFinite(rect.y)) return null;
  if (!Number.isFinite(rect.width) || !Number.isFinite(rect.height)) return null;
  if (rect.width <= 0 || rect.height <= 0) return null;
  return rect;
}

function rectsOverlap(a: Rect, b: Rect): boolean {
  return (
    a.x < b.x + b.width &&
    a.x + a.width > b.x &&
    a.y < b.y + b.height &&
    a.y + a.height > b.y
  );
}

function unionRect(a: Rect, b: Rect): Rect {
  const x1 = Math.min(a.x, b.x);
  const y1 = Math.min(a.y, b.y);
  const x2 = Math.max(a.x + a.width, b.x + b.width);
  const y2 = Math.max(a.y + a.height, b.y + b.height);

  return { x: x1, y: y1, width: x2 - x1, height: y2 - y1 };
}

export class DirtyRegionTracker {
  private dirty: Rect[] = [];

  markDirty(rect: Rect): void {
    const normalized = normalizeRect(rect);
    if (!normalized) return;

    let merged = normalized;
    for (let i = 0; i < this.dirty.length; ) {
      const existing = this.dirty[i];
      if (rectsOverlap(existing, merged)) {
        merged = unionRect(existing, merged);
        this.dirty.splice(i, 1);
        continue;
      }
      i++;
    }

    this.dirty.push(merged);
  }

  drain(): Rect[] {
    const drained = this.dirty;
    this.dirty = [];
    return drained;
  }

  clear(): void {
    this.dirty = [];
  }
}
