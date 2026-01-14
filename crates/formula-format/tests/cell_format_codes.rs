use formula_format::cell_format_code;

#[test]
fn cell_format_date_time_builtin_placeholders_match_excel_doc_codes() {
    // Mappings for 14–22 come from Microsoft Support `CELL` docs:
    // https://support.microsoft.com/en-us/office/cell-function-51bd39a5-f338-4dbe-a33f-955d67c2b2cf
    let cases: &[(u16, &str)] = &[
        (14, "D4"),
        (15, "D1"),
        (16, "D2"),
        (17, "D3"),
        (18, "D7"),
        (19, "D6"),
        (20, "D9"),
        (21, "D8"),
        (22, "D4"),
        // Durations (best-effort).
        (45, "D8"),
        (46, "D8"),
        (47, "D8"),
        // Locale-reserved ids (best-effort).
        (27, "D4"),
        (28, "D4"),
        (29, "D4"),
        (30, "D4"),
        (31, "D4"),
        (32, "D8"),
        (33, "D8"),
        (34, "D8"),
        (35, "D8"),
        (36, "D8"),
        // Excel-reserved locale ids (best-effort).
        (50, "D4"),
        (51, "D4"),
        (52, "D4"),
        (53, "D4"),
        (54, "D4"),
        (55, "D4"),
        (56, "D4"),
        (57, "D4"),
        (58, "D4"),
    ];

    for (id, expected) in cases {
        let fmt = format!("__builtin_numFmtId:{id}");
        assert_eq!(
            cell_format_code(Some(&fmt)),
            *expected,
            "placeholder {fmt:?} should map to {expected}"
        );
    }
}

#[test]
fn cell_format_date_time_explicit_patterns_match_excel_doc_codes() {
    let cases: &[(&str, &str)] = &[
        ("m/d/yyyy", "D4"),
        ("m/d/yy h:mm", "D4"),
        ("d-mmm-yy", "D1"),
        ("dd-mmm-yy", "D1"),
        ("d-mmm", "D2"),
        ("dd-mmm", "D2"),
        ("mmm-yy", "D3"),
        ("mm/dd", "D5"),
        ("h:mm AM/PM", "D7"),
        ("h:mm:ss AM/PM", "D6"),
        ("h:mm", "D9"),
        ("h:mm:ss", "D8"),
        // Durations / elapsed time.
        ("mm:ss", "D8"),
        ("[h]:mm:ss", "D8"),
        ("mm:ss.0", "D8"),
    ];

    for (pattern, expected) in cases {
        assert_eq!(
            cell_format_code(Some(pattern)),
            *expected,
            "pattern {pattern:?} should map to {expected}"
        );
    }
}
