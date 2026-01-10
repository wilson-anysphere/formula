export type PivotTableId = string;
export type SlicerId = string;
export type TimelineId = string;

export type SlicerItem = {
  key: string;
  label: string;
};

export type SlicerViewModel = {
  id: SlicerId;
  name: string;
  field: string;
  items: SlicerItem[];
  /**
   * Empty array means "all selected" (Excel slicer default).
   */
  selectedKeys: string[];
  connectedPivots: PivotTableId[];
};

export type TimelineViewModel = {
  id: TimelineId;
  name: string;
  field: string;
  start?: string;
  end?: string;
  connectedPivots: PivotTableId[];
};

