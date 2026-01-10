use formula_format::builtin_format_code;

#[test]
fn builtin_format_ids_cover_common_ooxml_cases() {
    assert_eq!(builtin_format_code(0), Some("General"));
    assert_eq!(builtin_format_code(14), Some("m/d/yyyy"));
    assert_eq!(builtin_format_code(49), Some("@"));
    assert_eq!(builtin_format_code(999), None);
}

