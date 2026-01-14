use std::collections::BTreeSet;

use crate::functions::{FunctionContext, Reference, SheetId};
use crate::value::ErrorKind;

/// Returns the Excel 1-based sheet number for an internal sheet id.
pub fn sheet_number(ctx: &dyn FunctionContext, sheet_id: &SheetId) -> Result<f64, ErrorKind> {
    match sheet_id {
        SheetId::Local(id) => ctx
            .sheet_order_index(*id)
            .map(|idx| (idx as f64) + 1.0)
            .ok_or(ErrorKind::NA),
        // We don't have external workbook sheet ordering information, so match Excel's
        // "not available" style error.
        SheetId::External(_) => Err(ErrorKind::NA),
    }
}

/// Returns the sheet number Excel should use for a multi-area reference.
///
/// `SHEET` returns a single scalar number even when passed a 3D reference like
/// `Sheet1:Sheet3!A1` (which expands to multiple sheet-local areas). We approximate Excel by
/// returning the first sheet in tab order.
pub fn sheet_number_for_references(
    ctx: &dyn FunctionContext,
    references: &[Reference],
) -> Result<f64, ErrorKind> {
    let mut min_tab_index: Option<usize> = None;
    for r in references {
        if let SheetId::Local(id) = &r.sheet_id {
            if let Some(tab_index) = ctx.sheet_order_index(*id) {
                min_tab_index = Some(match min_tab_index {
                    Some(existing) => existing.min(tab_index),
                    None => tab_index,
                });
            }
        }
    }

    match min_tab_index {
        Some(idx) => Ok((idx as f64) + 1.0),
        None => Err(ErrorKind::NA),
    }
}

/// Counts the number of distinct sheets referenced by `references`.
pub fn count_distinct_sheets(references: &[Reference]) -> usize {
    let mut set = BTreeSet::new();
    for r in references {
        set.insert(r.sheet_id.clone());
    }
    set.len()
}

/// Normalizes a stored formula string so it matches Excel's `FORMULATEXT` output convention of
/// including the leading `=`.
pub fn normalize_formula_text(formula: &str) -> String {
    // Match `CELL("contents")` behavior: preserve any leading whitespace, but ensure the first
    // *non-whitespace* character is `=` by inserting one if missing.
    if formula.trim_start().starts_with('=') {
        return formula.to_string();
    }
    format!("={formula}")
}
