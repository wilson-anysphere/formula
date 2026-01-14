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

export class DirtyRegionTracker {
  private dirty: Rect[] = [];

  markDirty(rect: Rect): void {
    const normalized = normalizeRect(rect);
    if (!normalized) return;

    let merged: Rect = normalized;
    for (let i = 0; i < this.dirty.length; ) {
      const existing = this.dirty[i];
      if (rectsOverlap(existing, merged)) {
        const x1 = Math.min(existing.x, merged.x);
        const y1 = Math.min(existing.y, merged.y);
        const x2 = Math.max(existing.x + existing.width, merged.x + merged.width);
        const y2 = Math.max(existing.y + existing.height, merged.y + merged.height);

        existing.x = x1;
        existing.y = y1;
        existing.width = x2 - x1;
        existing.height = y2 - y1;

        this.dirty.splice(i, 1);
        merged = existing;
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
    this.dirty.length = 0;
  }
}
