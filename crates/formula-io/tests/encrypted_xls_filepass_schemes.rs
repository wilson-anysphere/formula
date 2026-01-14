use std::path::PathBuf;

use formula_io::{detect_workbook_encryption, LegacyXlsFilePassScheme, WorkbookEncryption};

fn encrypted_xls_fixture_path(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../formula-xls/tests/fixtures/encrypted")
        .join(name)
}

#[test]
fn detects_xor_filepass_scheme() {
    let path = encrypted_xls_fixture_path("biff8_xor_pw_open.xls");
    let info = detect_workbook_encryption(&path).expect("detect encryption");
    assert_eq!(
        info,
        WorkbookEncryption::LegacyXlsFilePass {
            scheme: Some(LegacyXlsFilePassScheme::Xor),
        },
        "expected XOR FILEPASS, got {info:?}"
    );
}

#[test]
fn detects_rc4_standard_filepass_scheme() {
    let path = encrypted_xls_fixture_path("biff8_rc4_standard_pw_open.xls");
    let info = detect_workbook_encryption(&path).expect("detect encryption");
    assert_eq!(
        info,
        WorkbookEncryption::LegacyXlsFilePass {
            scheme: Some(LegacyXlsFilePassScheme::Rc4),
        },
        "expected RC4 (standard) FILEPASS, got {info:?}"
    );
}

