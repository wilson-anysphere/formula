use std::path::Path;

use formula_io::{detect_workbook_format, Error};

fn assert_encrypted_ooxml_bytes_detected(bytes: &[u8], stem: &str) {
    let tmp = tempfile::tempdir().expect("tempdir");

    // Test both correct and incorrect extensions to ensure content sniffing detects encryption
    // before attempting to open as legacy BIFF.
    for ext in ["xlsx", "xls", "xlsb"] {
        let path = tmp.path().join(format!("{stem}.{ext}"));
        std::fs::write(&path, bytes).expect("write encrypted fixture");

        let err = detect_workbook_format(&path).expect_err("expected encrypted workbook to error");
        assert!(
            matches!(err, Error::EncryptedWorkbook { .. }),
            "expected Error::EncryptedWorkbook, got {err:?}"
        );

        let msg = err.to_string().to_lowercase();
        assert!(
            msg.contains("encrypted") || msg.contains("password"),
            "expected error message to mention encryption/password protection, got: {msg}"
        );
    }
}

#[test]
fn detects_encrypted_ooxml_agile_fixture_if_present() {
    let fixture_path = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/encrypted/ooxml/agile.xlsx"
    ));

    if !fixture_path.exists() {
        // Fixture is optional (may be added by another task/agent); skip if not present.
        return;
    }

    let bytes = std::fs::read(fixture_path).expect("read agile encrypted fixture");
    assert_encrypted_ooxml_bytes_detected(&bytes, "agile");
}

#[test]
fn detects_encrypted_ooxml_standard_fixture_if_present() {
    let fixture_path = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/encrypted/ooxml/standard.xlsx"
    ));

    if !fixture_path.exists() {
        // Fixture is optional (may be added by another task/agent); skip if not present.
        return;
    }

    let bytes = std::fs::read(fixture_path).expect("read standard encrypted fixture");
    assert_encrypted_ooxml_bytes_detected(&bytes, "standard");
}

