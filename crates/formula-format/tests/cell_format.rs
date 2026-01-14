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
    // Excel's CELL("format") does not have a dedicated "fraction" code; fraction formats are
    // classified as fixed numbers with 0 decimal places.
    assert_eq!(builtin_format_code(12), Some("# ?/?"));
    assert_eq!(builtin_format_code(13), Some("# ??/??"));

    assert_eq!(
        cell_format_code(Some(builtin_format_code(12).unwrap())),
        "F0"
    );
    assert_eq!(
        cell_format_code(Some(builtin_format_code(13).unwrap())),
        "F0"
    );

    assert_eq!(
        cell_format_code(Some(&format!("{BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX}12"))),
        "F0"
    );
    assert_eq!(
        cell_format_code(Some(&format!("{BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX}13"))),
        "F0"
    );
}

#[test]
fn cell_format_code_accounting_builtins_41_through_44() {
    // Accounting formats without currency use the fixed ("F") classification.
    assert_eq!(
        cell_format_code(Some(builtin_format_code(41).unwrap())),
        "F0"
    );
    assert_eq!(
        cell_format_code(Some(builtin_format_code(43).unwrap())),
        "F2"
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
        "F0"
    );
    assert_eq!(
        cell_format_code(Some(&format!("{BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX}42"))),
        "C0"
    );
    assert_eq!(
        cell_format_code(Some(&format!("{BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX}43"))),
        "F2"
    );
    assert_eq!(
        cell_format_code(Some(&format!("{BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX}44"))),
        "C2"
    );
}
