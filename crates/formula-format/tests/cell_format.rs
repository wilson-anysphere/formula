use formula_format::{
    builtin_format_code, cell_format_code, BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX,
};

#[test]
fn cell_format_code_text_builtin_49_and_at_literal() {
    // Built-in Text format.
    assert_eq!(builtin_format_code(49), Some("@"));
    assert_eq!(
        cell_format_code(Some(builtin_format_code(49).unwrap())),
        "@"
    );
    assert_eq!(cell_format_code(Some("@")), "@");

    // Placeholder variant used by some importers.
    assert_eq!(
        cell_format_code(Some(&format!("{BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX}49"))),
        "@"
    );
}

#[test]
fn cell_format_code_fraction_builtins_12_13() {
    // Excel's `CELL("format")` returns `N` for numeric formats that don't fit the standard
    // fixed/currency/percent/scientific/date/time/text/general families. Fractions are one of
    // these non-classifiable numeric families (built-ins 12/13: `# ?/?` and `# ??/??`).
    assert_eq!(builtin_format_code(12), Some("# ?/?"));
    assert_eq!(builtin_format_code(13), Some("# ??/??"));

    assert_eq!(
        cell_format_code(Some(builtin_format_code(12).unwrap())),
        "N"
    );
    assert_eq!(
        cell_format_code(Some(builtin_format_code(13).unwrap())),
        "N"
    );

    assert_eq!(
        cell_format_code(Some(&format!("{BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX}12"))),
        "N"
    );
    assert_eq!(
        cell_format_code(Some(&format!("{BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX}13"))),
        "N"
    );
}

#[test]
fn cell_format_code_accounting_builtins_41_through_44() {
    // Accounting formats without currency still use `#,##0`-style grouping, so Excel classifies
    // them as numbers (`N*`) rather than fixed (`F*`).
    assert_eq!(
        cell_format_code(Some(builtin_format_code(41).unwrap())),
        "N0"
    );
    assert_eq!(
        cell_format_code(Some(builtin_format_code(43).unwrap())),
        "N2"
    );

    // Accounting formats with currency use the currency ("C") classification.
    assert_eq!(
        cell_format_code(Some(builtin_format_code(42).unwrap())),
        "C0"
    );
    assert_eq!(
        cell_format_code(Some(builtin_format_code(44).unwrap())),
        "C2"
    );

    // Placeholder variants.
    assert_eq!(
        cell_format_code(Some(&format!("{BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX}41"))),
        "N0"
    );
    assert_eq!(
        cell_format_code(Some(&format!("{BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX}42"))),
        "C0"
    );
    assert_eq!(
        cell_format_code(Some(&format!("{BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX}43"))),
        "N2"
    );
    assert_eq!(
        cell_format_code(Some(&format!("{BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX}44"))),
        "C2"
    );
}
