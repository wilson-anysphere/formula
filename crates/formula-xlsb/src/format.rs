use formula_model::push_column_label as push_column_label_model;

/// Format helpers intended for diagnostics and developer tooling.
pub fn format_a1(row: u32, col: u32) -> String {
    // XLSB row/col indices are 0-based, matching `formula-model::CellRef`.
    //
    // Delegate formatting to `formula-model` so row arithmetic is done in u64 and remains robust
    // for very large indices (e.g. u32::MAX) without overflow panics.
    let mut out = String::new();
    formula_model::push_a1_cell_ref(row, col, false, false, &mut out);
    out
}

/// Convert a 0-based column index to an Excel column label and append it to `out`.
pub fn push_column_label(col: u32, out: &mut String) {
    push_column_label_model(col, out);
}

/// Format bytes as an uppercase hex string, separated by spaces.
pub fn format_hex(bytes: &[u8]) -> String {
    let mut out = String::new();
    let _ = out.try_reserve(bytes.len().saturating_mul(3));
    const HEX: &[u8; 16] = b"0123456789ABCDEF";
    for (i, b) in bytes.iter().enumerate() {
        if i > 0 {
            out.push(' ');
        }
        out.push(HEX[(b >> 4) as usize] as char);
        out.push(HEX[(b & 0x0F) as usize] as char);
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn format_hex_matches_uppercase_two_digit_bytes() {
        assert_eq!(format_hex(&[]), "");
        assert_eq!(format_hex(&[0x00]), "00");
        assert_eq!(format_hex(&[0x01, 0xAB, 0xFF]), "01 AB FF");
    }

    #[test]
    fn format_a1_matches_cell_ref_formatting() {
        assert_eq!(format_a1(0, 0), "A1");
        assert_eq!(format_a1(1_048_575, 0), "A1048576");
        assert_eq!(format_a1(u32::MAX, 0), "A4294967296");
    }
}
