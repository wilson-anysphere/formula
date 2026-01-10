export class OutlineAxis {
  constructor() {
    /** @type {Map<number, {level:number, hidden:{user:boolean, outline:boolean, filter:boolean}, collapsed:boolean}>} */
    this.entries = new Map();
  }

  /**
   * @param {number} index
   */
  entry(index) {
    return (
      this.entries.get(index) ?? {
        level: 0,
        hidden: { user: false, outline: false, filter: false },
        collapsed: false,
      }
    );
  }

  /**
   * @param {number} index
   */
  entryMut(index) {
    let entry = this.entries.get(index);
    if (!entry) {
      entry = {
        level: 0,
        hidden: { user: false, outline: false, filter: false },
        collapsed: false,
      };
      this.entries.set(index, entry);
    }
    return entry;
  }

  clearOutlineHidden() {
    for (const entry of this.entries.values()) {
      entry.hidden.outline = false;
    }
  }
}

export function isHidden(hidden) {
  return hidden.user || hidden.outline || hidden.filter;
}

export function groupDetailRange(axis, summaryIndex, summaryLevel, summaryAfterDetails) {
  const targetLevel = Math.min(summaryLevel + 1, 7);
  if (targetLevel <= 0 || targetLevel > 7) return null;

  if (summaryAfterDetails) {
    if (summaryIndex <= 1) return null;
    let cursor = summaryIndex - 1;
    if (axis.entry(cursor).level < targetLevel) return null;
    while (cursor > 0 && axis.entry(cursor).level >= targetLevel) {
      cursor -= 1;
      if (cursor === 0) break;
    }
    const start = axis.entry(cursor).level >= targetLevel ? 1 : cursor + 1;
    const end = summaryIndex - 1;
    return [start, end, targetLevel];
  }

  let cursor = summaryIndex + 1;
  if (axis.entry(cursor).level < targetLevel) return null;
  while (axis.entry(cursor).level >= targetLevel) {
    cursor += 1;
    if (cursor >= Number.MAX_SAFE_INTEGER) break;
  }
  return [summaryIndex + 1, cursor - 1, targetLevel];
}

export class Outline {
  constructor() {
    this.pr = { summaryBelow: true, summaryRight: true, showOutlineSymbols: true };
    this.rows = new OutlineAxis();
    this.cols = new OutlineAxis();
  }

  toggleRowGroup(summaryIndex) {
    const entry = this.rows.entryMut(summaryIndex);
    entry.collapsed = !entry.collapsed;
    this.recomputeOutlineHiddenRows();
  }

  toggleColGroup(summaryIndex) {
    const entry = this.cols.entryMut(summaryIndex);
    entry.collapsed = !entry.collapsed;
    this.recomputeOutlineHiddenCols();
  }

  groupRows(start, end) {
    for (let i = start; i <= end; i += 1) {
      const entry = this.rows.entryMut(i);
      entry.level = Math.min(entry.level + 1, 7);
    }
  }

  ungroupRows(start, end) {
    for (let i = start; i <= end; i += 1) {
      const entry = this.rows.entryMut(i);
      entry.level = Math.max(entry.level - 1, 0);
      if (entry.level === 0) entry.collapsed = false;
    }
    this.recomputeOutlineHiddenRows();
  }

  groupCols(start, end) {
    for (let i = start; i <= end; i += 1) {
      const entry = this.cols.entryMut(i);
      entry.level = Math.min(entry.level + 1, 7);
    }
  }

  ungroupCols(start, end) {
    for (let i = start; i <= end; i += 1) {
      const entry = this.cols.entryMut(i);
      entry.level = Math.max(entry.level - 1, 0);
      if (entry.level === 0) entry.collapsed = false;
    }
    this.recomputeOutlineHiddenCols();
  }

  recomputeOutlineHiddenRows() {
    this._recomputeAxis(this.rows, this.pr.summaryBelow);
  }

  recomputeOutlineHiddenCols() {
    this._recomputeAxis(this.cols, this.pr.summaryRight);
  }

  _recomputeAxis(axis, summaryAfterDetails) {
    axis.clearOutlineHidden();

    for (const [summaryIndex, entry] of axis.entries.entries()) {
      if (!entry.collapsed) continue;
      const range = groupDetailRange(axis, summaryIndex, entry.level, summaryAfterDetails);
      if (!range) continue;
      const [start, end] = range;
      for (let i = start; i <= end; i += 1) {
        axis.entryMut(i).hidden.outline = true;
      }
    }
  }
}
