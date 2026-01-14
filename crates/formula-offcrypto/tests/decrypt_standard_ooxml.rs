// Fixtures in `tests/fixtures/` are copied from the MIT-licensed `nolze/msoffcrypto-tool` repo:
// https://github.com/nolze/msoffcrypto-tool
//
// The upstream project is MIT licensed; see their repository for the full license text.

use std::io::{Cursor, Write};
use std::path::PathBuf;

use formula_offcrypto::decrypt_standard_ooxml_from_bytes;
use formula_offcrypto::{EncryptionType, OffcryptoError};

fn fixture(path: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests").join("fixtures").join(path)
}

#[test]
fn decrypts_standard_fixture_docx() {
    let encrypted = std::fs::read(fixture("inputs/ecma376standard_password.docx"))
        .expect("read encrypted fixture");
    let expected = std::fs::read(fixture("outputs/ecma376standard_password_plain.docx"))
        .expect("read expected decrypted fixture");

    let decrypted =
        decrypt_standard_ooxml_from_bytes(encrypted, "Password1234_").expect("decrypt fixture");
    assert!(decrypted.starts_with(b"PK"));
    assert_eq!(decrypted, expected);
}

#[test]
fn wrong_password_returns_error() {
    let encrypted = std::fs::read(fixture("inputs/ecma376standard_password.docx"))
        .expect("read encrypted fixture");

    let err = decrypt_standard_ooxml_from_bytes(encrypted, "not-the-password")
        .expect_err("expected wrong password to error");
    assert!(matches!(err, formula_offcrypto::OffcryptoError::InvalidPassword));
}

#[test]
fn rejects_agile_fixture() {
    // `example_password.xlsx` is an Agile-encrypted OOXML package (EncryptionInfo v4.4).
    let encrypted =
        std::fs::read(fixture("inputs/example_password.xlsx")).expect("read encrypted fixture");

    let err = decrypt_standard_ooxml_from_bytes(encrypted, "any-password")
        .expect_err("expected Agile encryption to be rejected");
    assert!(
        matches!(
            err,
            OffcryptoError::UnsupportedEncryption {
                encryption_type: EncryptionType::Agile
            }
        ),
        "expected UnsupportedEncryption(Agile), got {err:?}"
    );
}

#[test]
fn missing_encryptioninfo_stream_returns_error() {
    // Ensure we return a structured error and never panic on missing required OLE streams.
    let cursor = Cursor::new(Vec::new());
    let ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    let bytes = ole.into_inner().into_inner();

    let err = decrypt_standard_ooxml_from_bytes(bytes, "pw").unwrap_err();
    assert!(
        matches!(err, OffcryptoError::MissingOleStream { stream } if stream == "EncryptionInfo"),
        "expected MissingOleStream(EncryptionInfo), got {err:?}"
    );
}

#[test]
fn invalid_ole_container_returns_io_error() {
    // Not a valid CFB/OLE file.
    let err = decrypt_standard_ooxml_from_bytes(vec![0u8; 32], "pw").unwrap_err();
    assert!(matches!(err, OffcryptoError::Io(_)), "expected Io error, got {err:?}");
}

#[test]
fn supports_encryptioninfo_with_leading_slash_stream_name() {
    // Some producers may store stream names with a leading slash. `read_ole_stream` attempts
    // both `EncryptionInfo` and `/EncryptionInfo`. Ensure the fallback works.
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    // Version header only; the decryptor should short-circuit on Agile (4.4) without trying to
    // parse the full XML or read the EncryptedPackage stream.
    let mut encryption_info = Vec::new();
    encryption_info.extend_from_slice(&4u16.to_le_bytes()); // major
    encryption_info.extend_from_slice(&4u16.to_le_bytes()); // minor
    encryption_info.extend_from_slice(&0u32.to_le_bytes()); // flags

    ole.create_stream("/EncryptionInfo")
        .expect("create /EncryptionInfo")
        .write_all(&encryption_info)
        .expect("write /EncryptionInfo");

    let bytes = ole.into_inner().into_inner();

    let err = decrypt_standard_ooxml_from_bytes(bytes, "pw").unwrap_err();
    assert!(
        matches!(
            err,
            OffcryptoError::UnsupportedEncryption {
                encryption_type: EncryptionType::Agile
            }
        ),
        "expected UnsupportedEncryption(Agile), got {err:?}"
    );
}
