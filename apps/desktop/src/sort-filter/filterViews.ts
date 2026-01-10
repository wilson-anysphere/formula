import type { AutoFilter, FilterViewId } from "./types";

export class FilterViews {
  private readonly views = new Map<FilterViewId, Map<string, AutoFilter>>();

  setFilter(viewId: FilterViewId, filter: AutoFilter) {
    if (!this.views.has(viewId)) {
      this.views.set(viewId, new Map());
    }
    this.views.get(viewId)!.set(filter.rangeA1, filter);
  }

  clearFilter(viewId: FilterViewId, rangeA1: string) {
    const view = this.views.get(viewId);
    if (!view) return;
    view.delete(rangeA1);
    if (view.size === 0) this.views.delete(viewId);
  }

  getFilter(viewId: FilterViewId, rangeA1: string): AutoFilter | undefined {
    return this.views.get(viewId)?.get(rangeA1);
  }
}

