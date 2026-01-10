use formula_model::Range;
use std::fmt::Write as _;

#[must_use]
pub fn write_merge_cells_section(merges: &[Range]) -> String {
    if merges.is_empty() {
        return String::new();
    }

    let mut out = String::new();
    let _ = writeln!(out, r#"<mergeCells count="{}">"#, merges.len());
    for merge in merges {
        let _ = writeln!(
            out,
            r#"  <mergeCell ref="{}"/>"#,
            merge
        );
    }
    out.push_str("</mergeCells>\n");
    out
}

/// Write a minimal worksheet XML containing sheet data and merge cells.
///
/// This is *not* a full-fidelity round-trip serializer; it is intended for tests and
/// for the merge-cells subsystem in isolation.
#[must_use]
pub fn write_worksheet_xml(merges: &[Range]) -> String {
    let mut out = String::new();
    out.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    out.push('\n');
    out.push_str(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
    );
    out.push('\n');
    out.push_str("  <sheetData/>\n");
    if !merges.is_empty() {
        out.push_str("  ");
        out.push_str(&write_merge_cells_section(merges).replace('\n', "\n  "));
        // `replace` adds trailing indentation; normalize.
        out = out.replace("\n  </mergeCells>", "\n</mergeCells>");
        if !out.ends_with('\n') {
            out.push('\n');
        }
    }
    out.push_str("</worksheet>\n");
    out
}
