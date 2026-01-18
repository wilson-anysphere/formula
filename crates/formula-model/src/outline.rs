use std::collections::BTreeMap;

use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(default)]
pub struct OutlinePr {
    pub summary_below: bool,
    pub summary_right: bool,
    pub show_outline_symbols: bool,
}

impl Default for OutlinePr {
    fn default() -> Self {
        Self {
            summary_below: true,
            summary_right: true,
            show_outline_symbols: true,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default, Serialize, Deserialize)]
#[serde(default)]
pub struct HiddenState {
    /// The row/column is explicitly hidden (eg via "Hide row").
    ///
    /// In OOXML this is stored as `hidden="1"`, but that same attribute is also
    /// used for outline-collapsed rows/columns. When reading XLSX we apply a
    /// heuristic to infer whether a hidden row/column is hidden because of an
    /// outline group collapse.
    pub user: bool,
    /// The row/column is hidden because it is inside a collapsed outline group.
    pub outline: bool,
    /// Reserved for filter hidden (Task 61 integration).
    pub filter: bool,
}

impl HiddenState {
    #[must_use]
    pub fn is_hidden(self) -> bool {
        self.user || self.outline || self.filter
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default, Serialize, Deserialize)]
#[serde(default)]
pub struct OutlineEntry {
    /// 0-7, as defined by the OOXML spec (`outlineLevel`).
    pub level: u8,
    pub hidden: HiddenState,
    /// Indicates this row/column is a collapsed summary row/column (`collapsed="1"`).
    pub collapsed: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Default, Serialize, Deserialize)]
#[serde(default)]
pub struct OutlineAxis {
    #[serde(default, skip_serializing_if = "BTreeMap::is_empty")]
    entries: BTreeMap<u32, OutlineEntry>,
}

impl OutlineAxis {
    #[must_use]
    pub fn entry(&self, index: u32) -> OutlineEntry {
        self.entries.get(&index).copied().unwrap_or_default()
    }

    #[must_use]
    pub fn get(&self, index: u32) -> Option<&OutlineEntry> {
        self.entries.get(&index)
    }

    pub fn entry_mut(&mut self, index: u32) -> &mut OutlineEntry {
        self.entries.entry(index).or_default()
    }

    pub fn iter(&self) -> impl Iterator<Item = (u32, &OutlineEntry)> {
        self.entries.iter().map(|(k, v)| (*k, v))
    }

    pub fn iter_mut(&mut self) -> impl Iterator<Item = (u32, &mut OutlineEntry)> {
        self.entries.iter_mut().map(|(k, v)| (*k, v))
    }

    pub fn clear_outline_hidden(&mut self) {
        for (_, entry) in self.iter_mut() {
            entry.hidden.outline = false;
        }
    }

    /// Sets whether the row/column at `index` is user-hidden (eg via "Hide row").
    ///
    /// This toggles only the `user` hidden bit, preserving outline/filter hidden state.
    /// When clearing (`hidden = false`), the entry is removed entirely if it becomes
    /// the default (no outline level, no hidden bits, not collapsed) to keep the
    /// serialized representation compact.
    pub fn set_user_hidden(&mut self, index: u32, hidden: bool) {
        if hidden {
            self.entry_mut(index).hidden.user = true;
            return;
        }

        let Some(entry) = self.entries.get_mut(&index) else {
            return;
        };
        entry.hidden.user = false;
        if *entry == OutlineEntry::default() {
            self.entries.remove(&index);
        }
    }

    /// Sets whether the row/column at `index` is hidden by an AutoFilter.
    ///
    /// This toggles only the `filter` hidden bit, preserving user/outline hidden state.
    /// When clearing (`hidden = false`), the entry is removed entirely if it becomes
    /// the default (no outline level, no hidden bits, not collapsed) to keep the
    /// serialized representation compact.
    pub fn set_filter_hidden(&mut self, index: u32, hidden: bool) {
        if hidden {
            self.entry_mut(index).hidden.filter = true;
            return;
        }

        let Some(entry) = self.entries.get_mut(&index) else {
            return;
        };
        entry.hidden.filter = false;
        if *entry == OutlineEntry::default() {
            self.entries.remove(&index);
        }
    }

    /// Clears filter-hidden flags for any stored entries within `[start, end]`.
    ///
    /// This does not create new entries for rows/columns that were not previously
    /// tracked in the outline map.
    pub fn clear_filter_hidden_range(&mut self, start: u32, end: u32) {
        if start > end {
            return;
        }
        let key_count = self
            .entries
            .range(start..=end)
            .filter(|(_, v)| v.hidden.filter)
            .count();
        let mut keys: Vec<u32> = Vec::new();
        if keys.try_reserve_exact(key_count).is_err() {
            debug_assert!(
                false,
                "allocation failed (outline clear filter hidden range, count={key_count})"
            );
            return;
        }
        for (k, v) in self.entries.range(start..=end) {
            if v.hidden.filter {
                keys.push(*k);
            }
        }
        for key in keys {
            self.set_filter_hidden(key, false);
        }
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Default, Serialize, Deserialize)]
#[serde(default)]
pub struct Outline {
    pub pr: OutlinePr,
    pub rows: OutlineAxis,
    pub cols: OutlineAxis,
}

impl Outline {
    /// Collapses or expands the outline group controlled by `summary_index`.
    ///
    /// The `summary_index` is the row/column where Excel renders the +/- box.
    /// This is the row below the detail block when `summary_below = true` (the
    /// default), and the row above the detail block when `summary_below = false`.
    pub fn toggle_row_group(&mut self, summary_index: u32) -> bool {
        let entry = self.rows.entry_mut(summary_index);
        entry.collapsed = !entry.collapsed;
        self.recompute_outline_hidden_rows();
        true
    }

    pub fn toggle_col_group(&mut self, summary_index: u32) -> bool {
        let entry = self.cols.entry_mut(summary_index);
        entry.collapsed = !entry.collapsed;
        self.recompute_outline_hidden_cols();
        true
    }

    pub fn group_rows(&mut self, start: u32, end: u32) {
        for index in start..=end {
            let entry = self.rows.entry_mut(index);
            entry.level = (entry.level + 1).min(7);
        }
        self.recompute_outline_hidden_rows();
    }

    pub fn ungroup_rows(&mut self, start: u32, end: u32) {
        for index in start..=end {
            let entry = self.rows.entry_mut(index);
            entry.level = entry.level.saturating_sub(1);
            if entry.level == 0 {
                entry.collapsed = false;
            }
        }
        self.recompute_outline_hidden_rows();
    }

    pub fn group_cols(&mut self, start: u32, end: u32) {
        for index in start..=end {
            let entry = self.cols.entry_mut(index);
            entry.level = (entry.level + 1).min(7);
        }
        self.recompute_outline_hidden_cols();
    }

    pub fn ungroup_cols(&mut self, start: u32, end: u32) {
        for index in start..=end {
            let entry = self.cols.entry_mut(index);
            entry.level = entry.level.saturating_sub(1);
            if entry.level == 0 {
                entry.collapsed = false;
            }
        }
        self.recompute_outline_hidden_cols();
    }

    /// Recomputes which rows are hidden because of collapsed outline groups.
    ///
    /// This does not modify user/filter hidden flags.
    pub fn recompute_outline_hidden_rows(&mut self) {
        let summary_below = self.pr.summary_below;
        self.rows.clear_outline_hidden();

        let collapsed_count = self.rows.iter().filter(|(_, entry)| entry.collapsed).count();
        let mut collapsed_summaries: Vec<(u32, u8)> = Vec::new();
        if collapsed_summaries
            .try_reserve_exact(collapsed_count)
            .is_err()
        {
            debug_assert!(
                false,
                "allocation failed (outline collapsed row summaries, count={collapsed_count})"
            );
            return;
        }
        for (index, entry) in self.rows.iter() {
            if entry.collapsed {
                collapsed_summaries.push((index, entry.level));
            }
        }

        for (summary_index, summary_level) in collapsed_summaries {
            let Some((start, end, _level)) =
                group_detail_range(&self.rows, summary_index, summary_level, summary_below)
            else {
                continue;
            };
            for index in start..=end {
                self.rows.entry_mut(index).hidden.outline = true;
            }
        }
    }

    pub fn recompute_outline_hidden_cols(&mut self) {
        let summary_right = self.pr.summary_right;
        self.cols.clear_outline_hidden();

        let collapsed_count = self.cols.iter().filter(|(_, entry)| entry.collapsed).count();
        let mut collapsed_summaries: Vec<(u32, u8)> = Vec::new();
        if collapsed_summaries
            .try_reserve_exact(collapsed_count)
            .is_err()
        {
            debug_assert!(
                false,
                "allocation failed (outline collapsed col summaries, count={collapsed_count})"
            );
            return;
        }
        for (index, entry) in self.cols.iter() {
            if entry.collapsed {
                collapsed_summaries.push((index, entry.level));
            }
        }

        for (summary_index, summary_level) in collapsed_summaries {
            let Some((start, end, _level)) =
                group_detail_range(&self.cols, summary_index, summary_level, summary_right)
            else {
                continue;
            };
            for index in start..=end {
                self.cols.entry_mut(index).hidden.outline = true;
            }
        }
    }

    /// Finds the next visible row after `start` in the given `direction`.
    ///
    /// This is intended for selection/navigation logic so the UI can skip
    /// outline-hidden rows (and, eventually, filter-hidden rows).
    #[must_use]
    pub fn next_visible_row(&self, start: u32, direction: i32, max_row: u32) -> Option<u32> {
        next_visible_index(&self.rows, start, direction, max_row)
    }

    /// Finds the next visible column after `start` in the given `direction`.
    #[must_use]
    pub fn next_visible_col(&self, start: u32, direction: i32, max_col: u32) -> Option<u32> {
        next_visible_index(&self.cols, start, direction, max_col)
    }
}

fn group_detail_range(
    axis: &OutlineAxis,
    summary_index: u32,
    summary_level: u8,
    summary_after_details: bool,
) -> Option<(u32, u32, u8)> {
    let target_level = summary_level.saturating_add(1);
    if target_level == 0 || target_level > 7 {
        return None;
    }

    if summary_after_details {
        if summary_index <= 1 {
            return None;
        }
        let mut cursor = summary_index - 1;
        if axis.entry(cursor).level < target_level {
            return None;
        }
        while cursor > 0 && axis.entry(cursor).level >= target_level {
            cursor = cursor.saturating_sub(1);
            if cursor == 0 {
                break;
            }
        }
        let start = if axis.entry(cursor).level >= target_level {
            1
        } else {
            cursor + 1
        };
        let end = summary_index - 1;
        Some((start, end, target_level))
    } else {
        let mut cursor = summary_index + 1;
        if axis.entry(cursor).level < target_level {
            return None;
        }
        while axis.entry(cursor).level >= target_level {
            cursor = cursor.saturating_add(1);
            if cursor == u32::MAX {
                break;
            }
        }
        let start = summary_index + 1;
        let end = cursor - 1;
        Some((start, end, target_level))
    }
}

fn next_visible_index(axis: &OutlineAxis, start: u32, direction: i32, max: u32) -> Option<u32> {
    let dir = match direction {
        -1 => -1i64,
        1 => 1i64,
        _ => return None,
    };
    let mut cursor = start as i64 + dir;
    while cursor >= 1 && cursor <= max as i64 {
        let index = cursor as u32;
        if !axis.entry(index).hidden.is_hidden() {
            return Some(index);
        }
        cursor += dir;
    }
    None
}
