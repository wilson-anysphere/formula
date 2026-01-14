use formula_format::{cell_format_info, FormatOptions, Locale};

#[test]
fn cell_color_and_parentheses_match_excel_semantics() {
    let options = FormatOptions::default();

    // --- 1) One-section formats ---
    //
    // Excel's CELL("color") / CELL("parentheses") flags are about *explicit negative sections*.
    // With only one section, Excel auto-prefixes '-' for negatives and reports 0/0 even if the
    // section contains a color token or parentheses literals.
    assert_eq!(cell_format_info(Some("[Red]0"), &options).color, 0);
    assert_eq!(cell_format_info(Some("(0)"), &options).parentheses, 0);

    // --- 2) Two-section formats ---
    let info = cell_format_info(Some("0;(0)"), &options);
    assert_eq!(info.parentheses, 1);
    assert_eq!(info.color, 0);

    let info = cell_format_info(Some("0;[Red]0"), &options);
    assert_eq!(info.color, 1);
    assert_eq!(info.parentheses, 0);

    let info = cell_format_info(Some("0;[Red](0)"), &options);
    assert_eq!(info.color, 1);
    assert_eq!(info.parentheses, 1);

    // --- 3) Conditional sections ---
    //
    // When any section has a condition, Excel evaluates them in-order to pick the section.
    // These cases ensure the *negative* value branch is used when computing the flags.
    assert_eq!(cell_format_info(Some("[<0][Red]0;0"), &options).color, 1);
    assert_eq!(cell_format_info(Some("[>=0]0;[Red]0"), &options).color, 1);

    // --- 4) Parentheses inside quotes should NOT count ---
    assert_eq!(
        cell_format_info(Some(r#"0;"(neg)"0"#), &options).parentheses,
        0
    );

    // --- 5) Escaped parentheses should NOT count ---
    assert_eq!(
        cell_format_info(Some(r#"0;\(0\)"#), &options).parentheses,
        0
    );

    // --- 6) Bracket tokens containing parentheses-like chars shouldn’t be mis-detected ---
    //
    // Currency symbols in `[$...-...]` tokens can legally contain parentheses. Excel's
    // `CELL("parentheses")` should not treat those as "negative in parentheses".
    assert_eq!(
        cell_format_info(Some(r#"0;[$(USD)-409]0"#), &options).parentheses,
        0
    );

    // --- 7) Underscore / fill layout token operands should NOT count ---
    assert_eq!(
        cell_format_info(Some("0;0_(0_)"), &options).parentheses,
        0
    );
    assert_eq!(
        cell_format_info(Some("0;0*(0*)"), &options).parentheses,
        0
    );
}

#[test]
fn cell_format_info_resolves_builtin_placeholders() {
    let options = FormatOptions {
        locale: Locale::en_us(),
        ..FormatOptions::default()
    };

    // Built-in format 6 is `$#,##0_);[Red]($#,##0)` (currency with red parentheses negatives).
    let info = cell_format_info(Some("__builtin_numFmtId:6"), &options);
    assert_eq!(info.color, 1);
    assert_eq!(info.parentheses, 1);

    // Unknown built-ins should behave like General.
    let info = cell_format_info(Some("__builtin_numFmtId:999"), &options);
    assert_eq!(info.color, 0);
    assert_eq!(info.parentheses, 0);
}
