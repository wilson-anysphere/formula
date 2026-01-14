use std::fmt;

/// A compressed formatting segment applied to a single column.
///
/// This mirrors the document model's `formatRunsByCol` representation:
/// - Each run covers rows `[start_row, end_row_exclusive)`.
/// - `style_id` references an entry in the workbook [`formula_model::StyleTable`] (`0` is the
///   default/empty style).
///
/// When computing effective formatting, these runs have precedence:
/// `sheet < col < row < range-run < cell`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct FormatRun {
    pub start_row: u32,
    pub end_row_exclusive: u32,
    pub style_id: u32,
}

impl fmt::Display for FormatRun {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            f,
            "FormatRun{{start_row={}, end_row_exclusive={}, style_id={}}}",
            self.start_row, self.end_row_exclusive, self.style_id
        )
    }
}

/// Returns the style id for `row` in a sorted run list.
///
/// Runs are expected to be sorted by `start_row` (as in `Engine::set_format_runs_by_col`).
/// If multiple runs overlap, the *last* matching run wins to preserve deterministic behavior.
///
/// Returns `0` when no run covers `row` (or when only `style_id=0` runs match).
pub fn style_id_for_row_in_runs(runs: Option<&[FormatRun]>, row: u32) -> u32 {
    let Some(runs) = runs else {
        return 0;
    };
    let mut style_id = 0;
    for run in runs {
        if row < run.start_row {
            break;
        }
        if row >= run.end_row_exclusive {
            continue;
        }
        if run.style_id != 0 {
            style_id = run.style_id;
        }
    }
    style_id
}
