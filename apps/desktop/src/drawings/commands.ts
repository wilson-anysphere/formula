import { createDrawingObjectId, type Anchor, type DrawingObject } from "./types";
import { pxToEmu } from "./overlay";

/**
 * Normalize drawing z-order to a dense 0..n-1 range.
 *
 * Ordering is defined by:
 *  1) ascending `zOrder`
 *  2) ascending `id` as a deterministic tie-breaker
 *
 * The returned array preserves the input array order (only zOrder values change).
 */
export function normalizeZOrder(objects: DrawingObject[]): DrawingObject[] {
  if (objects.length <= 1) return objects;

  const ordered = [...objects].sort((a, b) => a.zOrder - b.zOrder || a.id - b.id);
  const zOrderById = new Map<number, number>();
  for (let i = 0; i < ordered.length; i += 1) {
    zOrderById.set(ordered[i]!.id, i);
  }

  let changed = false;
  const next = objects.map((obj) => {
    const z = zOrderById.get(obj.id);
    if (z == null) return obj;
    if (obj.zOrder === z) return obj;
    changed = true;
    return { ...obj, zOrder: z };
  });

  return changed ? next : objects;
}

export function deleteSelected(objects: DrawingObject[], id: number): DrawingObject[] {
  const next = objects.filter((o) => o.id !== id);
  if (next.length === objects.length) return objects;
  return normalizeZOrder(next);
}

export type DuplicateSelectedResult = { objects: DrawingObject[]; duplicatedId: number };

export function duplicateSelected(
  objects: DrawingObject[],
  id: number,
  opts: { offsetPx?: number } = {},
): DuplicateSelectedResult | null {
  const source = objects.find((o) => o.id === id);
  if (!source) return null;

  const offsetPx = typeof opts.offsetPx === "number" && Number.isFinite(opts.offsetPx) ? opts.offsetPx : 10;
  const dxEmu = pxToEmu(offsetPx);
  const dyEmu = pxToEmu(offsetPx);

  const normalized = normalizeZOrder(objects);
  // Drawing ids must be globally unique across collaborators; avoid deterministic counters like
  // `max + 1` that can collide when two users duplicate simultaneously.
  let nextId = createDrawingObjectId();
  while (normalized.some((o) => o.id === nextId)) {
    nextId = createDrawingObjectId();
  }
  const clone: DrawingObject = {
    ...deepCloneDrawingObject(source),
    id: nextId,
    anchor: shiftAnchor(source.anchor, dxEmu, dyEmu),
    zOrder: normalized.length, // top
  };

  const merged = [...normalized, clone];
  return { objects: normalizeZOrder(merged), duplicatedId: nextId };
}

export function bringToFront(objects: DrawingObject[], id: number): DrawingObject[] {
  return reorderZOrder(objects, id, "toFront");
}

export function sendToBack(objects: DrawingObject[], id: number): DrawingObject[] {
  return reorderZOrder(objects, id, "toBack");
}

export function bringForward(objects: DrawingObject[], id: number): DrawingObject[] {
  return reorderZOrder(objects, id, "forward");
}

export function sendBackward(objects: DrawingObject[], id: number): DrawingObject[] {
  return reorderZOrder(objects, id, "backward");
}

type ReorderMode = "toFront" | "toBack" | "forward" | "backward";

function reorderZOrder(objects: DrawingObject[], id: number, mode: ReorderMode): DrawingObject[] {
  if (objects.length <= 1) return objects;

  const normalized = normalizeZOrder(objects);
  const stack = [...normalized].sort((a, b) => a.zOrder - b.zOrder || a.id - b.id);
  const index = stack.findIndex((o) => o.id === id);
  if (index === -1) return objects;

  if (mode === "toFront") {
    if (index === stack.length - 1) return normalized;
    const [picked] = stack.splice(index, 1);
    stack.push(picked!);
  } else if (mode === "toBack") {
    if (index === 0) return normalized;
    const [picked] = stack.splice(index, 1);
    stack.unshift(picked!);
  } else if (mode === "forward") {
    if (index === stack.length - 1) return normalized;
    const tmp = stack[index]!;
    stack[index] = stack[index + 1]!;
    stack[index + 1] = tmp;
  } else if (mode === "backward") {
    if (index === 0) return normalized;
    const tmp = stack[index]!;
    stack[index] = stack[index - 1]!;
    stack[index - 1] = tmp;
  }

  const zOrderById = new Map<number, number>();
  for (let i = 0; i < stack.length; i += 1) {
    zOrderById.set(stack[i]!.id, i);
  }

  let changed = false;
  const reordered = normalized.map((obj) => {
    const z = zOrderById.get(obj.id);
    if (z == null) return obj;
    if (obj.zOrder === z) return obj;
    changed = true;
    return { ...obj, zOrder: z };
  });

  return changed ? reordered : normalized;
}

function deepCloneDrawingObject(obj: DrawingObject): DrawingObject {
  // DrawingObject is plain data (no functions/classes) so structured cloning is safe here.
  // Using JSON stringify avoids relying on structuredClone availability in older runtimes.
  return JSON.parse(JSON.stringify(obj)) as DrawingObject;
}

function shiftAnchor(anchor: Anchor, dxEmu: number, dyEmu: number): Anchor {
  switch (anchor.type) {
    case "oneCell":
      return {
        ...anchor,
        from: {
          ...anchor.from,
          offset: {
            xEmu: anchor.from.offset.xEmu + dxEmu,
            yEmu: anchor.from.offset.yEmu + dyEmu,
          },
        },
      };
    case "twoCell":
      return {
        ...anchor,
        from: {
          ...anchor.from,
          offset: {
            xEmu: anchor.from.offset.xEmu + dxEmu,
            yEmu: anchor.from.offset.yEmu + dyEmu,
          },
        },
        to: {
          ...anchor.to,
          offset: {
            xEmu: anchor.to.offset.xEmu + dxEmu,
            yEmu: anchor.to.offset.yEmu + dyEmu,
          },
        },
      };
    case "absolute":
      return {
        ...anchor,
        pos: {
          xEmu: anchor.pos.xEmu + dxEmu,
          yEmu: anchor.pos.yEmu + dyEmu,
        },
      };
  }
}
