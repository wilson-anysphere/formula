use std::collections::BTreeSet;

use crate::functions::{Reference, SheetId};
use crate::value::ErrorKind;

/// Returns the Excel 1-based sheet number for an internal sheet id.
pub fn sheet_number(sheet_id: &SheetId) -> Result<f64, ErrorKind> {
    match sheet_id {
        SheetId::Local(id) => Ok((*id as f64) + 1.0),
        // We don't have external workbook sheet ordering information, so match Excel's
        // "not available" style error.
        SheetId::External(_) => Err(ErrorKind::NA),
    }
}

/// Returns the sheet number Excel should use for a multi-area reference.
///
/// `SHEET` returns a single scalar number even when passed a 3D reference like
/// `Sheet1:Sheet3!A1` (which expands to multiple sheet-local areas). We approximate Excel by
/// returning the first (lowest) local sheet id, which corresponds to workbook order.
pub fn sheet_number_for_references(references: &[Reference]) -> Result<f64, ErrorKind> {
    let mut min_local: Option<usize> = None;
    for r in references {
        if let SheetId::Local(id) = &r.sheet_id {
            min_local = Some(match min_local {
                Some(existing) => existing.min(*id),
                None => *id,
            });
        }
    }

    match min_local {
        Some(id) => Ok((id as f64) + 1.0),
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
    if formula.starts_with('=') {
        return formula.to_string();
    }
    format!("={formula}")
}
