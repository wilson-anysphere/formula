use std::fmt;

/// A compressed formatting segment for a column in a sheet.
///
/// The segment covers the half-open row interval `[start_row, end_row_exclusive)`.
///
/// `style_id` references an entry in the engine's style table and represents a patch that
/// participates in effective style merge semantics.
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

/// Returns the style id for `row` in a sorted, non-overlapping run list.
///
/// Returns `0` when no run covers `row`.
pub fn style_id_for_row_in_runs(runs: Option<&[FormatRun]>, row: u32) -> u32 {
    let Some(runs) = runs else {
        return 0;
    };
    if runs.is_empty() {
        return 0;
    }

    // Runs are sorted by start_row. Find the last run whose start_row <= row.
    let mut lo: usize = 0;
    let mut hi: usize = runs.len();
    while lo < hi {
        let mid = lo + (hi - lo) / 2;
        if runs[mid].start_row <= row {
            lo = mid + 1;
        } else {
            hi = mid;
        }
    }

    let idx = lo.saturating_sub(1);
    let run = &runs[idx];
    if row >= run.start_row && row < run.end_row_exclusive {
        run.style_id
    } else {
        0
    }
}
