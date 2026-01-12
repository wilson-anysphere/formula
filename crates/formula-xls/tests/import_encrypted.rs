use std::io::Write;

mod common;

use common::xls_fixture_builder;

#[test]
fn errors_on_encrypted_xls_filepass() {
    let bytes = xls_fixture_builder::build_encrypted_filepass_fixture_xls();
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(&bytes).expect("write xls bytes");

    let err = formula_xls::import_xls_path(tmp.path()).expect_err("expected encrypted workbook");
    assert!(matches!(err, formula_xls::ImportError::EncryptedWorkbook));

    let msg = err.to_string().to_lowercase();
    assert!(
        msg.contains("encrypted"),
        "expected error message to mention encryption; got: {msg}"
    );
    assert!(
        msg.contains("password"),
        "expected error message to mention password protection; got: {msg}"
    );
}
