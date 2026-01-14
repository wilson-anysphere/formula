//! Additional end-to-end tests for **non-ASCII / Unicode passwords**, including non-BMP
//! codepoints (emoji surrogate pairs).
//!
//! Most encrypted workbook coverage lives in `encrypted_ooxml_decrypt.rs` and
//! `encrypted_ooxml_fixtures.rs`. This file exists specifically to ensure we exercise password
//! handling for codepoints outside the BMP, where UTF-16 uses surrogate pairs.
#![cfg(all(feature = "encrypted-workbooks", not(target_arch = "wasm32")))]

use std::io::{Cursor, Write as _};

use formula_io::{open_workbook_model_with_password, open_workbook_with_password, Error, Workbook};
use formula_model::{CellRef, CellValue};
use ms_offcrypto_writer::Ecma376AgileWriter;
use rand::{rngs::StdRng, SeedableRng as _};

fn build_tiny_xlsx() -> Vec<u8> {
    let mut workbook = formula_model::Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1").expect("add sheet");
    let sheet = workbook.sheet_mut(sheet_id).expect("sheet exists");
    sheet.set_value(
        CellRef::from_a1("A1").unwrap(),
        CellValue::String("Hello".to_string()),
    );

    let mut cursor = Cursor::new(Vec::new());
    formula_io::xlsx::write_workbook_to_writer(&workbook, &mut cursor).expect("write xlsx");
    cursor.into_inner()
}

fn encrypt_bytes_with_password(plain: &[u8], password: &str) -> Vec<u8> {
    let mut cursor = Cursor::new(Vec::new());
    let mut rng = StdRng::from_seed([0u8; 32]);
    let mut agile = Ecma376AgileWriter::create(&mut rng, password, &mut cursor)
        .expect("create agile writer");
    agile
        .write_all(plain)
        .expect("write plaintext to agile writer");
    agile.finalize().expect("finalize agile writer");
    cursor.into_inner()
}

#[test]
fn opens_encrypted_ooxml_with_unicode_passwords_including_emoji() {
    let plain_xlsx = build_tiny_xlsx();

    // Include:
    // - non-ASCII BMP codepoints
    // - a non-BMP emoji (surrogate pair in UTF-16)
    // - leading/trailing whitespace to ensure caller input is not trimmed
    //
    // Note: encrypting via `ms-offcrypto-writer` uses a real-world Agile spinCount (100k), which is
    // expensive. Keep this list minimal while still covering both the whitespace and non-whitespace
    // cases (so we exercise both branches of the trimming checks below).
    let passwords = ["pässwörd🔒", " 密码🔒 "];

    for password in passwords {
        let encrypted = encrypt_bytes_with_password(&plain_xlsx, password);
        let tmp = tempfile::tempdir().expect("tempdir");
        let path = tmp.path().join("encrypted.xlsx");
        std::fs::write(&path, encrypted).expect("write encrypted bytes");

        // Missing password should be distinguished from the empty string.
        let err = open_workbook_model_with_password(&path, None).expect_err("expected password required");
        assert!(
            matches!(err, Error::PasswordRequired { .. }),
            "expected Error::PasswordRequired, got {err:?}"
        );

        // Password strings must match exactly; callers must not be normalized or trimmed.
        let trimmed = password.trim();
        if trimmed != password {
            // If the *correct* password has whitespace, the trimmed version must fail.
            let err = open_workbook_model_with_password(&path, Some(trimmed))
                .expect_err("expected invalid password");
            assert!(
                matches!(err, Error::InvalidPassword { .. }),
                "expected Error::InvalidPassword, got {err:?}"
            );
        } else {
            // If the correct password has no leading/trailing whitespace, adding whitespace must fail
            // (guards against `trim()` on caller input).
            let wrong = format!("{password} ");
            let err = open_workbook_model_with_password(&path, Some(&wrong))
                .expect_err("expected invalid password");
            assert!(
                matches!(err, Error::InvalidPassword { .. }),
                "expected Error::InvalidPassword, got {err:?}"
            );
        }

        // Correct password succeeds.
        let wb =
            open_workbook_model_with_password(&path, Some(password)).expect("open decrypted model");
        let sheet = wb.sheet_by_name("Sheet1").expect("Sheet1 missing");
        assert_eq!(
            sheet.value(CellRef::from_a1("A1").unwrap()),
            CellValue::String("Hello".to_string())
        );

        let wb = open_workbook_with_password(&path, Some(password)).expect("open decrypted workbook");
        assert!(matches!(wb, Workbook::Xlsx(_)), "expected Workbook::Xlsx, got {wb:?}");
    }
}
