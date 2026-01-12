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
        .add_sheet("Bad/Name".to_string(), None, None)
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
    let err = state
        .add_sheet("   ".to_string(), None, None)
        .expect_err("expected empty sheet name error");
    match err {
        AppStateError::WhatIf(msg) => assert!(
            msg.contains("cannot be blank"),
            "expected blank name error, got {msg:?}"
        ),
        other => panic!("expected WhatIf error, got {other:?}"),
    }
}

#[test]
fn add_sheet_rejects_leading_or_trailing_apostrophe() {
    let (mut state, _sheet1_id, _sheet2_id) = loaded_state_with_two_sheets();

    for name in ["'Leading", "Trailing'"] {
        let err = state
            .add_sheet(name.to_string(), None, None)
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
        .add_sheet(long_name, None, None)
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
fn add_sheet_truncates_base_name_to_fit_unique_suffix() {
    let long = "a".repeat(31);
    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet(long.clone());
    workbook.add_sheet("Sheet2".to_string());

    let mut state = AppState::new();
    state.load_workbook(workbook);

    let added = state
        .add_sheet(long, None, None)
        .expect("expected add_sheet to succeed with a unique suffix");
    assert_eq!(added.name, format!("{} 2", "a".repeat(29)));
    assert_eq!(added.name.len(), 31);
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
    let err = state
        .create_sheet("   ".to_string())
        .expect_err("expected empty sheet name error");
    match err {
        AppStateError::WhatIf(msg) => assert!(
            msg.contains("cannot be blank"),
            "expected blank name error, got {msg:?}"
        ),
        other => panic!("expected WhatIf error, got {other:?}"),
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
    let err = state
        .rename_sheet(&sheet1_id, "   ".to_string())
        .expect_err("expected empty sheet name error");
    match err {
        AppStateError::WhatIf(msg) => assert!(
            msg.contains("cannot be blank"),
            "expected blank name error, got {msg:?}"
        ),
        other => panic!("expected WhatIf error, got {other:?}"),
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
