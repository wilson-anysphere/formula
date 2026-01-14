use formula_format::cell_format_code;

#[test]
fn cell_format_code_returns_n_for_unclassified_formats() {
    // Excel returns `N` for fraction formats (built-ins 12/13: `# ?/?` and `# ??/??`).
    assert_eq!(cell_format_code(Some("__builtin_numFmtId:12")), "N");
    assert_eq!(cell_format_code(Some("__builtin_numFmtId:13")), "N");

    // Literal-only format strings (no numeric placeholders) are also non-classifiable.
    assert_eq!(cell_format_code(Some(r#""foo""#)), "N");
}

#[test]
fn cell_format_code_still_classifies_common_numeric_families() {
    assert_eq!(cell_format_code(Some("General")), "G");
    assert_eq!(cell_format_code(Some("0.00")), "F2");
    assert_eq!(cell_format_code(Some("$#,##0.00")), "C2");
    assert_eq!(cell_format_code(Some("0.00%")), "P2");
    assert_eq!(cell_format_code(Some("0.00E+00")), "S2");
    assert_eq!(cell_format_code(Some("@")), "@");
}

