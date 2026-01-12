use desktop::file_io::Workbook;
use desktop::state::{AppState, AppStateError};

fn loaded_state_with_two_sheets() -> (AppState, String, String) {
    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet("Sheet1".to_string());
    workbook.add_sheet("Sheet2".to_string());
    let sheet1_id = workbook.sheets[0].id.clone();
    let sheet2_id = workbook.sheets[1].id.clone();

    let mut state = AppState::new();
    state.load_workbook(workbook);
    (state, sheet1_id, sheet2_id)
}

#[test]
fn add_sheet_rejects_invalid_character() {
    let (mut state, _sheet1_id, _sheet2_id) = loaded_state_with_two_sheets();
    let err = state
        .add_sheet("Bad/Name".to_string(), None, None, None)
        .expect_err("expected invalid sheet name error");
    match err {
        AppStateError::WhatIf(msg) => {
            assert!(
                msg.contains("invalid character"),
                "expected invalid character error, got {msg:?}"
            );
        }
        other => panic!("expected WhatIf error, got {other:?}"),
    }
}

#[test]
fn add_sheet_rejects_empty_string() {
    let (mut state, _sheet1_id, _sheet2_id) = loaded_state_with_two_sheets();
    for name in ["   ", "\n\t"] {
        let err = state
            .add_sheet(name.to_string(), None, None, None)
            .expect_err("expected empty sheet name error");
        match err {
            AppStateError::WhatIf(msg) => assert!(
                msg.contains("cannot be blank"),
                "expected blank name error, got {msg:?}"
            ),
            other => panic!("expected WhatIf error, got {other:?}"),
        }
    }
}

#[test]
fn add_sheet_rejects_leading_or_trailing_apostrophe() {
    let (mut state, _sheet1_id, _sheet2_id) = loaded_state_with_two_sheets();

    for name in ["'Leading", "Trailing'"] {
        let err = state
            .add_sheet(name.to_string(), None, None, None)
            .expect_err("expected invalid sheet name error");
        match err {
            AppStateError::WhatIf(msg) => assert!(
                msg.contains("apostrophe"),
                "expected apostrophe error, got {msg:?}"
            ),
            other => panic!("expected WhatIf error, got {other:?}"),
        }
    }
}

#[test]
fn add_sheet_rejects_names_longer_than_31_chars() {
    let (mut state, _sheet1_id, _sheet2_id) = loaded_state_with_two_sheets();
    let long_name = "a".repeat(32);
    let err = state
        .add_sheet(long_name, None, None, None)
        .expect_err("expected sheet name too long error");
    match err {
        AppStateError::WhatIf(msg) => assert!(
            msg.contains("cannot exceed"),
            "expected length error, got {msg:?}"
        ),
        other => panic!("expected WhatIf error, got {other:?}"),
    }
}

#[test]
fn add_sheet_rejects_names_longer_than_31_utf16_units() {
    let (mut state, _sheet1_id, _sheet2_id) = loaded_state_with_two_sheets();
    // 🙂 is a non-BMP character (2 UTF-16 code units). 16 of them => 32 UTF-16 units.
    let long_name = "🙂".repeat(16);
    let err = state
        .add_sheet(long_name, None, None, None)
        .expect_err("expected sheet name too long error");
    match err {
        AppStateError::WhatIf(msg) => assert!(
            msg.contains("cannot exceed"),
            "expected length error, got {msg:?}"
        ),
        other => panic!("expected WhatIf error, got {other:?}"),
    }
}

#[test]
fn add_sheet_disambiguates_ascii_duplicate() {
    let (mut state, _sheet1_id, _sheet2_id) = loaded_state_with_two_sheets();
    let added = state
        .add_sheet("Sheet1".to_string(), None, None, None)
        .expect("expected add_sheet to disambiguate duplicate");
    assert_eq!(added.name, "Sheet1 2");
}

#[test]
fn add_sheet_truncates_base_name_to_fit_unique_suffix() {
    let long = "a".repeat(31);
    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet(long.clone());
    workbook.add_sheet("Sheet2".to_string());

    let mut state = AppState::new();
    state.load_workbook(workbook);

    let added = state
        .add_sheet(long, None, None, None)
        .expect("expected add_sheet to succeed with a unique suffix");
    assert_eq!(added.name, format!("{} 2", "a".repeat(29)));
    assert_eq!(added.name.len(), 31);
}

#[test]
fn add_sheet_truncates_base_name_to_fit_two_digit_unique_suffix() {
    let long = "a".repeat(31);
    let mut workbook = Workbook::new_empty(None);
    // Seed the workbook with the base name and all single-digit disambiguations so the next suffix
    // will be two digits (" 10"), which is 3 UTF-16 code units.
    workbook.add_sheet(long.clone());
    for n in 2..=9 {
        workbook.add_sheet(format!("{} {n}", "a".repeat(29)));
    }

    let mut state = AppState::new();
    state.load_workbook(workbook);

    let added = state
        .add_sheet(long, None, None, None)
        .expect("expected add_sheet to succeed with a unique suffix");
    assert_eq!(added.name, format!("{} 10", "a".repeat(28)));
    assert_eq!(added.name.encode_utf16().count(), 31);
}

#[test]
fn add_sheet_truncates_unicode_base_name_to_fit_two_digit_unique_suffix() {
    // Base is exactly 31 UTF-16 code units: 15 non-BMP chars (30 units) + "a" (1 unit).
    let base = format!("{}a", "🙂".repeat(15));
    assert_eq!(base.encode_utf16().count(), 31);

    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet(base.clone());

    // Seed single-digit disambiguations (" 2" .. " 9"). These are 2 UTF-16 units, so the base
    // portion can use 29 units. Because the base starts with 15 emoji, truncation happens *before*
    // we reach the trailing "a", so the disambiguated names are just 14 emojis + suffix.
    for n in 2..=9 {
        workbook.add_sheet(format!("{} {n}", "🙂".repeat(14)));
    }

    let mut state = AppState::new();
    state.load_workbook(workbook);

    let added = state
        .add_sheet(base, None, None, None)
        .expect("expected add_sheet to succeed with a unique suffix");

    // Suffix " 10" uses 3 UTF-16 code units, leaving 28 for the base; that's exactly 14 emojis.
    assert_eq!(added.name, format!("{} 10", "🙂".repeat(14)));
    assert_eq!(added.name.encode_utf16().count(), 31);
}

#[test]
fn add_sheet_truncates_base_name_by_utf16_units_to_fit_unique_suffix() {
    // 🙂 counts as 2 UTF-16 code units in Excel; build an exactly-31-unit name.
    let long = format!("{}a", "🙂".repeat(15));
    assert_eq!(long.encode_utf16().count(), 31);

    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet(long.clone());
    workbook.add_sheet("Sheet2".to_string());

    let mut state = AppState::new();
    state.load_workbook(workbook);

    let added = state
        .add_sheet(long, None, None, None)
        .expect("expected add_sheet to succeed with a unique suffix");

    // Suffix " 2" uses 2 UTF-16 code units, leaving 29 for the base; we can only fit 14 emojis
    // (28 units) before exceeding 29.
    assert_eq!(added.name, format!("{} 2", "🙂".repeat(14)));
    assert!(added.name.encode_utf16().count() <= 31);
}

#[test]
fn add_sheet_disambiguates_unicode_case_insensitive_duplicate() {
    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet("é".to_string());
    workbook.add_sheet("Sheet2".to_string());

    let mut state = AppState::new();
    state.load_workbook(workbook);

    // Excel compares sheet names case-insensitively across Unicode; adding "É" should be treated
    // as a duplicate of "é" and disambiguated with a suffix.
    let added = state
        .add_sheet("É".to_string(), None, None, None)
        .expect("expected add_sheet to disambiguate unicode duplicate");
    assert_eq!(added.name, "É 2");
}

#[test]
fn add_sheet_disambiguates_unicode_nfkc_equivalent_duplicate() {
    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet("fi".to_string());
    workbook.add_sheet("Sheet2".to_string());

    let mut state = AppState::new();
    state.load_workbook(workbook);

    // The "fi" ligature (U+FB01) NFKC-normalizes to "fi". Adding it should be treated as a
    // duplicate and disambiguated with a suffix.
    let added = state
        .add_sheet("\u{FB01}".to_string(), None, None, None)
        .expect("expected add_sheet to disambiguate NFKC-equivalent duplicate");
    assert_eq!(added.name, "\u{FB01} 2");
}

#[test]
fn add_sheet_disambiguates_unicode_case_folding_expansion_duplicate() {
    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet("straße".to_string());
    workbook.add_sheet("Sheet2".to_string());

    let mut state = AppState::new();
    state.load_workbook(workbook);

    // German ß uppercases to "SS". Excel compares sheet names case-insensitively across Unicode,
    // so adding "STRASSE" should be treated as a duplicate of "straße".
    let added = state
        .add_sheet("STRASSE".to_string(), None, None, None)
        .expect("expected add_sheet to disambiguate unicode duplicate");
    assert_eq!(added.name, "STRASSE 2");
}

#[test]
fn add_sheet_disambiguates_unicode_nfkc_duplicate_when_suffix_collides() {
    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet("fi".to_string());
    workbook.add_sheet("fi 2".to_string());

    let mut state = AppState::new();
    state.load_workbook(workbook);

    // U+FB01 NFKC-normalizes to "fi". The first disambiguation would be "\u{FB01} 2", which
    // NFKC-normalizes to "fi 2" and therefore collides; we should skip to suffix 3.
    let added = state
        .add_sheet("\u{FB01}".to_string(), None, None, None)
        .expect("expected add_sheet to disambiguate NFKC-equivalent duplicate");
    assert_eq!(added.name, "\u{FB01} 3");
}

#[test]
fn add_sheet_disambiguates_unicode_case_folding_expansion_duplicate_when_suffix_collides() {
    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet("straße".to_string());
    workbook.add_sheet("STRASSE 2".to_string());

    let mut state = AppState::new();
    state.load_workbook(workbook);

    // German ß uppercases to "SS". The first disambiguation ("STRASSE 2") collides with an
    // existing sheet; we should skip to suffix 3.
    let added = state
        .add_sheet("STRASSE".to_string(), None, None, None)
        .expect("expected add_sheet to disambiguate unicode duplicate");
    assert_eq!(added.name, "STRASSE 3");
}

#[test]
fn create_sheet_rejects_invalid_character() {
    let (mut state, _sheet1_id, _sheet2_id) = loaded_state_with_two_sheets();
    let err = state
        .create_sheet("Bad/Name".to_string())
        .expect_err("expected invalid sheet name error");
    match err {
        AppStateError::WhatIf(msg) => assert!(
            msg.contains("invalid character"),
            "expected invalid character error, got {msg:?}"
        ),
        other => panic!("expected WhatIf error, got {other:?}"),
    }
}

#[test]
fn create_sheet_rejects_empty_string() {
    let (mut state, _sheet1_id, _sheet2_id) = loaded_state_with_two_sheets();
    for name in ["   ", "\n\t"] {
        let err = state
            .create_sheet(name.to_string())
            .expect_err("expected empty sheet name error");
        match err {
            AppStateError::WhatIf(msg) => assert!(
                msg.contains("cannot be blank"),
                "expected blank name error, got {msg:?}"
            ),
            other => panic!("expected WhatIf error, got {other:?}"),
        }
    }
}

#[test]
fn create_sheet_rejects_leading_or_trailing_apostrophe() {
    let (mut state, _sheet1_id, _sheet2_id) = loaded_state_with_two_sheets();

    for name in ["'Leading", "Trailing'"] {
        let err = state
            .create_sheet(name.to_string())
            .expect_err("expected invalid sheet name error");
        match err {
            AppStateError::WhatIf(msg) => assert!(
                msg.contains("apostrophe"),
                "expected apostrophe error, got {msg:?}"
            ),
            other => panic!("expected WhatIf error, got {other:?}"),
        }
    }
}

#[test]
fn create_sheet_rejects_duplicate_name_case_insensitive() {
    let (mut state, _sheet1_id, _sheet2_id) = loaded_state_with_two_sheets();
    let err = state
        .create_sheet("sheet1".to_string())
        .expect_err("expected duplicate sheet name error");
    match err {
        AppStateError::WhatIf(msg) => assert!(
            msg.contains("already exists"),
            "expected duplicate error, got {msg:?}"
        ),
        other => panic!("expected WhatIf error, got {other:?}"),
    }
}

#[test]
fn create_sheet_rejects_names_longer_than_31_chars() {
    let (mut state, _sheet1_id, _sheet2_id) = loaded_state_with_two_sheets();
    let long_name = "a".repeat(32);
    let err = state
        .create_sheet(long_name)
        .expect_err("expected sheet name too long error");
    match err {
        AppStateError::WhatIf(msg) => {
            assert!(
                msg.contains("cannot exceed"),
                "expected length error, got {msg:?}"
            );
        }
        other => panic!("expected WhatIf error, got {other:?}"),
    }
}

#[test]
fn create_sheet_rejects_names_longer_than_31_utf16_units() {
    let (mut state, _sheet1_id, _sheet2_id) = loaded_state_with_two_sheets();
    // 🙂 is a non-BMP character (2 UTF-16 code units). 16 of them => 32 UTF-16 units.
    let long_name = "🙂".repeat(16);
    let err = state
        .create_sheet(long_name)
        .expect_err("expected sheet name too long error");
    match err {
        AppStateError::WhatIf(msg) => assert!(
            msg.contains("cannot exceed"),
            "expected length error, got {msg:?}"
        ),
        other => panic!("expected WhatIf error, got {other:?}"),
    }
}

#[test]
fn rename_sheet_rejects_invalid_character() {
    let (mut state, sheet1_id, _sheet2_id) = loaded_state_with_two_sheets();
    let err = state
        .rename_sheet(&sheet1_id, "Bad/Name".to_string())
        .expect_err("expected invalid sheet name error");
    match err {
        AppStateError::WhatIf(msg) => assert!(
            msg.contains("invalid character"),
            "expected invalid character error, got {msg:?}"
        ),
        other => panic!("expected WhatIf error, got {other:?}"),
    }
}

#[test]
fn rename_sheet_rejects_names_longer_than_31_chars() {
    let (mut state, sheet1_id, _sheet2_id) = loaded_state_with_two_sheets();
    let long_name = "a".repeat(32);
    let err = state
        .rename_sheet(&sheet1_id, long_name)
        .expect_err("expected sheet name too long error");
    match err {
        AppStateError::WhatIf(msg) => assert!(
            msg.contains("cannot exceed"),
            "expected length error, got {msg:?}"
        ),
        other => panic!("expected WhatIf error, got {other:?}"),
    }
}

#[test]
fn rename_sheet_rejects_names_longer_than_31_utf16_units() {
    let (mut state, sheet1_id, _sheet2_id) = loaded_state_with_two_sheets();
    // 🙂 is a non-BMP character (2 UTF-16 code units). 16 of them => 32 UTF-16 units.
    let long_name = "🙂".repeat(16);
    let err = state
        .rename_sheet(&sheet1_id, long_name)
        .expect_err("expected sheet name too long error");
    match err {
        AppStateError::WhatIf(msg) => assert!(
            msg.contains("cannot exceed"),
            "expected length error, got {msg:?}"
        ),
        other => panic!("expected WhatIf error, got {other:?}"),
    }
}

#[test]
fn rename_sheet_rejects_leading_or_trailing_apostrophe() {
    let (mut state, sheet1_id, _sheet2_id) = loaded_state_with_two_sheets();

    for name in ["'Leading", "Trailing'"] {
        let err = state
            .rename_sheet(&sheet1_id, name.to_string())
            .expect_err("expected invalid sheet name error");
        match err {
            AppStateError::WhatIf(msg) => assert!(
                msg.contains("apostrophe"),
                "expected apostrophe error, got {msg:?}"
            ),
            other => panic!("expected WhatIf error, got {other:?}"),
        }
    }
}

#[test]
fn rename_sheet_rejects_empty_string() {
    let (mut state, sheet1_id, _sheet2_id) = loaded_state_with_two_sheets();
    for name in ["   ", "\n\t"] {
        let err = state
            .rename_sheet(&sheet1_id, name.to_string())
            .expect_err("expected empty sheet name error");
        match err {
            AppStateError::WhatIf(msg) => assert!(
                msg.contains("cannot be blank"),
                "expected blank name error, got {msg:?}"
            ),
            other => panic!("expected WhatIf error, got {other:?}"),
        }
    }
}

#[test]
fn rename_sheet_rejects_duplicate_name_case_insensitive() {
    let (mut state, _sheet1_id, sheet2_id) = loaded_state_with_two_sheets();
    let err = state
        .rename_sheet(&sheet2_id, "sheet1".to_string())
        .expect_err("expected duplicate sheet name error");
    match err {
        AppStateError::WhatIf(msg) => assert!(
            msg.contains("already exists"),
            "expected duplicate error, got {msg:?}"
        ),
        other => panic!("expected WhatIf error, got {other:?}"),
    }
}

#[test]
fn rename_sheet_allows_case_only_change_on_same_sheet() {
    let (mut state, sheet1_id, _sheet2_id) = loaded_state_with_two_sheets();
    state
        .rename_sheet(&sheet1_id, "sheet1".to_string())
        .expect("expected rename to different case to succeed");
    let workbook = state.get_workbook().expect("workbook loaded");
    let sheet1 = workbook.sheet(&sheet1_id).expect("sheet exists");
    assert_eq!(sheet1.name, "sheet1");
}

#[test]
fn rename_sheet_allows_unicode_case_only_change_on_same_sheet() {
    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet("é".to_string());
    workbook.add_sheet("Sheet2".to_string());
    let sheet1_id = workbook.sheets[0].id.clone();
    let mut state = AppState::new();
    state.load_workbook(workbook);

    state
        .rename_sheet(&sheet1_id, "É".to_string())
        .expect("expected rename to different unicode case to succeed");
    let workbook = state.get_workbook().expect("workbook loaded");
    let sheet1 = workbook.sheet(&sheet1_id).expect("sheet exists");
    assert_eq!(sheet1.name, "É");
}

#[test]
fn rename_sheet_allows_unicode_case_folding_expansion_on_same_sheet() {
    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet("straße".to_string());
    workbook.add_sheet("Sheet2".to_string());
    let sheet1_id = workbook.sheets[0].id.clone();
    let mut state = AppState::new();
    state.load_workbook(workbook);

    // German ß uppercases to "SS". This is a "case-only" rename under Excel-like Unicode
    // case-insensitive matching and should be allowed on the same sheet.
    state
        .rename_sheet(&sheet1_id, "STRASSE".to_string())
        .expect("expected rename to unicode case-folding expansion to succeed");
    let workbook = state.get_workbook().expect("workbook loaded");
    let sheet1 = workbook.sheet(&sheet1_id).expect("sheet exists");
    assert_eq!(sheet1.name, "STRASSE");
}

#[test]
fn rename_sheet_rejects_duplicate_name_with_unicode_case_folding() {
    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet("é".to_string());
    workbook.add_sheet("Sheet2".to_string());
    let sheet2_id = workbook.sheets[1].id.clone();

    let mut state = AppState::new();
    state.load_workbook(workbook);

    // `eq_ignore_ascii_case` would treat these as distinct; Excel compares sheet names
    // case-insensitively across Unicode.
    let err = state
        .rename_sheet(&sheet2_id, "É".to_string())
        .expect_err("expected duplicate sheet name error");

    match err {
        AppStateError::WhatIf(msg) => assert!(
            msg.contains("already exists"),
            "expected duplicate error, got {msg:?}"
        ),
        other => panic!("expected WhatIf error, got {other:?}"),
    }
}

#[test]
fn rename_sheet_rejects_duplicate_name_with_unicode_case_folding_expansion() {
    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet("straße".to_string());
    workbook.add_sheet("Sheet2".to_string());
    let sheet2_id = workbook.sheets[1].id.clone();

    let mut state = AppState::new();
    state.load_workbook(workbook);

    // German ß uppercases to "SS". Excel compares sheet names case-insensitively across Unicode, so
    // this should be treated as a duplicate.
    let err = state
        .rename_sheet(&sheet2_id, "STRASSE".to_string())
        .expect_err("expected duplicate sheet name error");

    match err {
        AppStateError::WhatIf(msg) => assert!(
            msg.contains("already exists"),
            "expected duplicate error, got {msg:?}"
        ),
        other => panic!("expected WhatIf error, got {other:?}"),
    }
}

#[test]
fn rename_sheet_rejects_duplicate_name_with_unicode_nfkc_equivalence() {
    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet("fi".to_string());
    workbook.add_sheet("Sheet2".to_string());
    let sheet2_id = workbook.sheets[1].id.clone();

    let mut state = AppState::new();
    state.load_workbook(workbook);

    // The "fi" ligature (U+FB01) NFKC-normalizes to "fi".
    let err = state
        .rename_sheet(&sheet2_id, "\u{FB01}".to_string())
        .expect_err("expected duplicate sheet name error");

    match err {
        AppStateError::WhatIf(msg) => assert!(
            msg.contains("already exists"),
            "expected duplicate error, got {msg:?}"
        ),
        other => panic!("expected WhatIf error, got {other:?}"),
    }
}

#[test]
fn create_sheet_rejects_duplicate_name_with_unicode_nfkc_equivalence() {
    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet("fi".to_string());
    workbook.add_sheet("Sheet2".to_string());

    let mut state = AppState::new();
    state.load_workbook(workbook);

    // The "fi" ligature (U+FB01) NFKC-normalizes to "fi".
    let err = state
        .create_sheet("\u{FB01}".to_string())
        .expect_err("expected duplicate sheet name error");

    match err {
        AppStateError::WhatIf(msg) => assert!(
            msg.contains("already exists"),
            "expected duplicate error, got {msg:?}"
        ),
        other => panic!("expected WhatIf error, got {other:?}"),
    }
}

#[test]
fn create_sheet_rejects_duplicate_name_with_unicode_case_folding() {
    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet("straße".to_string());
    workbook.add_sheet("Sheet2".to_string());

    let mut state = AppState::new();
    state.load_workbook(workbook);

    // German ß uppercases to "SS". Excel compares sheet names case-insensitively across Unicode, so
    // this should be treated as a duplicate.
    let err = state
        .create_sheet("STRASSE".to_string())
        .expect_err("expected duplicate sheet name error");

    match err {
        AppStateError::WhatIf(msg) => assert!(
            msg.contains("already exists"),
            "expected duplicate error, got {msg:?}"
        ),
        other => panic!("expected WhatIf error, got {other:?}"),
    }
}
