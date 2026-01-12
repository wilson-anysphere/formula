use formula_xls::sanitize_sheet_name;

#[test]
fn strips_embedded_nuls() {
    let existing = Vec::new();
    let sanitized = sanitize_sheet_name("She\0et1", 1, &existing);
    assert_eq!(sanitized, "Sheet1");
}

#[test]
fn replaces_invalid_characters() {
    let existing = Vec::new();
    let sanitized = sanitize_sheet_name(r"A:B\C/D?E*F[G]H", 1, &existing);
    assert_eq!(sanitized, "A_B_C_D_E_F_G_H");
}

#[test]
fn truncates_to_excel_max_sheet_name_length() {
    let existing = Vec::new();
    let long = "a".repeat(40);
    let sanitized = sanitize_sheet_name(&long, 1, &existing);
    assert_eq!(sanitized, "a".repeat(31));
    assert_eq!(sanitized.encode_utf16().count(), 31);
}

#[test]
fn dedupes_name_collisions() {
    let existing = vec!["Sheet".to_string(), "Sheet (2)".to_string()];
    let sanitized = sanitize_sheet_name("Sheet", 1, &existing);
    assert_eq!(sanitized, "Sheet (3)");
}

