use std::fmt::Write;

/// Format helpers intended for diagnostics and developer tooling.
pub fn format_a1(row: u32, col: u32) -> String {
    let mut out = String::new();
    push_column_label(col, &mut out);
    // XLSB row/col indices are 0-based.
    let _ = write!(&mut out, "{}", row + 1);
    out
}

/// Convert a 0-based column index to an Excel column label and append it to `out`.
pub fn push_column_label(mut col: u32, out: &mut String) {
    // Excel column labels are 1-based.
    col += 1;
    let mut buf = [0u8; 10];
    let mut i = 0usize;
    while col > 0 {
        let rem = ((col - 1) % 26) as u8;
        buf[i] = b'A' + rem;
        i += 1;
        col = (col - 1) / 26;
    }
    for ch in buf[..i].iter().rev() {
        out.push(*ch as char);
    }
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

