//! End-to-end decryption test for the Standard (CryptoAPI) OOXML open path.
//!
//! This is gated behind `encrypted-workbooks` because password-based decryption is optional.
#![cfg(all(feature = "encrypted-workbooks", not(target_arch = "wasm32")))]

use std::path::PathBuf;

use formula_io::{open_workbook_model, open_workbook_with_options, OpenOptions, Workbook};
use formula_model::{CellRef, CellValue};

fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encrypted/ooxml")
        .join(rel)
}

fn assert_expected_contents(workbook: &formula_model::Workbook) {
    assert_eq!(workbook.sheets.len(), 1, "expected exactly one sheet");
    assert_eq!(workbook.sheets[0].name, "Sheet1");

    let sheet = workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(
        sheet.value(CellRef::from_a1("A1").unwrap()),
        CellValue::Number(1.0)
    );
    assert_eq!(
        sheet.value(CellRef::from_a1("B1").unwrap()),
        CellValue::String("Hello".to_string())
    );
}

#[test]
fn decrypts_standard_fixture_via_streaming_reader() {
    let plaintext_path = fixture_path("plaintext.xlsx");
    let standard_path = fixture_path("standard.xlsx");

    let plaintext = open_workbook_model(&plaintext_path).expect("open plaintext.xlsx");
    assert_expected_contents(&plaintext);

    let decrypted = open_workbook_with_options(
        &standard_path,
        OpenOptions {
            password: Some("password".to_string()),
        },
    )
    .expect("open standard.xlsx with password");

    let Workbook::Xlsx(package) = decrypted else {
        panic!("expected Workbook::Xlsx for decrypted Standard workbook");
    };

    // Materialize the decrypted ZIP bytes and parse them into a model workbook so we can validate
    // cell contents. (The streaming decrypt path itself is exercised by unit tests inside
    // `encrypted_ooxml.rs`.)
    let decrypted_bytes = package
        .write_to_bytes()
        .expect("serialize decrypted workbook package to bytes");
    let decrypted_model = formula_xlsx::read_workbook_from_reader(std::io::Cursor::new(decrypted_bytes))
        .expect("parse decrypted bytes to model workbook");
    assert_expected_contents(&decrypted_model);

    // Sanity: compare some key cell values with the known-good plaintext fixture.
    let plain_sheet = plaintext.sheet_by_name("Sheet1").expect("Sheet1 missing");
    let dec_sheet = decrypted_model.sheet_by_name("Sheet1").expect("Sheet1 missing");
    for addr in ["A1", "B1"] {
        let cref = CellRef::from_a1(addr).unwrap();
        assert_eq!(dec_sheet.value(cref), plain_sheet.value(cref), "addr={addr}");
    }
}
