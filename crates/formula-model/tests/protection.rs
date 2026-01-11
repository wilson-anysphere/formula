use formula_model::{
    hash_legacy_password, verify_legacy_password, CellRef, Protection, Range,
    SheetProtectionAction, Style, StyleTable, Worksheet,
};

#[test]
fn legacy_password_hash_matches_known_vectors() {
    // Published examples (Excel legacy sheet/workbook protection hash).
    // See e.g. XlsxWriter/OpenXML docs.
    let cases = [
        ("password", 0x83AF),
        ("test", 0xCBEB),
        ("1234", 0xCC3D),
        ("", 0xCE4B),
    ];

    for (pw, expected) in cases {
        assert_eq!(hash_legacy_password(pw), expected, "pw={pw:?}");
        assert!(verify_legacy_password(pw, expected));
        assert!(!verify_legacy_password("wrong", expected));
    }
}

#[test]
fn sheet_protection_editable_cell_semantics() {
    let mut styles = StyleTable::new();
    let locked_style_id = styles.intern(Style {
        protection: Some(Protection {
            locked: true,
            hidden: false,
        }),
        ..Style::default()
    });
    let unlocked_style_id = styles.intern(Style {
        protection: Some(Protection {
            locked: false,
            hidden: false,
        }),
        ..Style::default()
    });

    let a1 = CellRef::new(0, 0);
    let b1 = CellRef::new(0, 1);
    let c1 = CellRef::new(0, 2);

    // Unprotected sheet: locked/unlocked flags do not block editing.
    let mut sheet = Worksheet::new(1, "Sheet1");
    sheet.set_style_id(a1, locked_style_id);
    sheet.set_style_id(b1, unlocked_style_id);

    assert!(sheet.is_cell_editable(a1, &styles));
    assert!(sheet.is_cell_editable(b1, &styles));

    // Protected sheet: locked cells are not editable; unlocked cells are.
    sheet.sheet_protection.enabled = true;
    assert!(!sheet.is_cell_editable(a1, &styles));
    assert!(sheet.is_cell_editable(b1, &styles));

    // Default style (no stored cell) is locked in Excel.
    assert!(!sheet.is_cell_editable(c1, &styles));
}

#[test]
fn merged_cells_use_anchor_cell_protection() {
    let mut styles = StyleTable::new();
    let locked_style_id = styles.intern(Style {
        protection: Some(Protection {
            locked: true,
            hidden: false,
        }),
        ..Style::default()
    });
    let unlocked_style_id = styles.intern(Style {
        protection: Some(Protection {
            locked: false,
            hidden: false,
        }),
        ..Style::default()
    });

    let mut sheet = Worksheet::new(1, "Sheet1");
    sheet
        .merge_range(Range::new(CellRef::new(0, 0), CellRef::new(0, 1)))
        .unwrap();

    sheet.sheet_protection.enabled = true;

    // Anchor is locked, non-anchor is unlocked -> not editable (Excel uses anchor cell).
    sheet.set_style_id(CellRef::new(0, 0), locked_style_id);
    sheet.set_style_id(CellRef::new(0, 1), unlocked_style_id);
    assert!(!sheet.is_cell_editable(CellRef::new(0, 1), &styles));

    // Anchor is unlocked, non-anchor is locked -> editable (anchor decides).
    sheet.set_style_id(CellRef::new(0, 0), unlocked_style_id);
    sheet.set_style_id(CellRef::new(0, 1), locked_style_id);
    assert!(sheet.is_cell_editable(CellRef::new(0, 1), &styles));
}

#[test]
fn sheet_protection_can_perform_respects_flags() {
    let mut sheet = Worksheet::new(1, "Sheet1");
    sheet.sheet_protection.enabled = true;

    assert!(!sheet.can_perform(SheetProtectionAction::FormatCells));
    sheet.sheet_protection.format_cells = true;
    assert!(sheet.can_perform(SheetProtectionAction::FormatCells));
}
