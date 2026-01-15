use std::fmt::Write;

use formula_model::push_column_label as push_column_label_model;

/// Format helpers intended for diagnostics and developer tooling.
pub fn format_a1(row: u32, col: u32) -> String {
    // XLSB row/col indices are 0-based, matching `formula-model::CellRef`.
    //
    // Delegate formatting to `CellRef::to_a1` so row/col arithmetic is done in u64 and
    // remains robust for very large indices (e.g. u32::MAX) without debug overflow panics.
    formula_model::CellRef::new(row, col).to_a1()
}

/// Convert a 0-based column index to an Excel column label and append it to `out`.
pub fn push_column_label(col: u32, out: &mut String) {
    push_column_label_model(col, out);
}

/// Format bytes as an uppercase hex string, separated by spaces.
pub fn format_hex(bytes: &[u8]) -> String {
    let mut out = String::with_capacity(bytes.len().saturating_mul(3));
    for (i, b) in bytes.iter().enumerate() {
        if i > 0 {
            out.push(' ');
        }
        let _ = write!(&mut out, "{:02X}", b);
    }
    out
}
