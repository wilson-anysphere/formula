import type { SlicerViewModel, TimelineViewModel } from "./types";

export function toggleSlicerSelection(
  slicer: SlicerViewModel,
  key: string,
): SlicerViewModel {
  const selected = new Set(slicer.selectedKeys);
  if (selected.has(key)) {
    selected.delete(key);
  } else {
    selected.add(key);
  }
  return { ...slicer, selectedKeys: Array.from(selected) };
}

export function clearSlicerSelection(slicer: SlicerViewModel): SlicerViewModel {
  return { ...slicer, selectedKeys: [] };
}

export function setTimelineRange(
  timeline: TimelineViewModel,
  start?: string,
  end?: string,
): TimelineViewModel {
  return { ...timeline, start, end };
}

