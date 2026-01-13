#![cfg(not(target_arch = "wasm32"))]

// Fixtures in `tests/fixtures/` are copied from the MIT-licensed `nolze/msoffcrypto-tool` repo:
// https://github.com/nolze/msoffcrypto-tool
//
// The upstream project is MIT licensed; see their repository for the full license text.

use std::io::{Cursor, Read, Write};
use std::path::PathBuf;

use formula_offcrypto::decrypt_standard_ooxml_from_bytes;
use formula_offcrypto::{EncryptionType, OffcryptoError};

fn fixture(path: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join(path)
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
    assert!(
        matches!(err, formula_offcrypto::OffcryptoError::InvalidPassword),
        "expected InvalidPassword, got {err:?}"
    );
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
            &err,
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
        matches!(
            &err,
            OffcryptoError::InvalidStructure(msg) if msg.contains("missing `EncryptionInfo` stream")
        ),
        "expected InvalidStructure missing EncryptionInfo, got {err:?}"
    );
}

#[test]
fn invalid_ole_container_returns_error() {
    // Not a valid CFB/OLE file.
    let err = decrypt_standard_ooxml_from_bytes(vec![0u8; 32], "pw").unwrap_err();
    assert!(
        matches!(
            &err,
            OffcryptoError::InvalidStructure(msg) if msg.contains("failed to open OLE compound file")
        ),
        "expected InvalidStructure for invalid OLE container, got {err:?}"
    );
}

#[test]
fn supports_encryptioninfo_with_leading_slash_stream_name() {
    // Some producers may store stream names with a leading slash. Ensure the decrypt entrypoint
    // can still detect Agile encryption and return a structured error (without needing the
    // EncryptedPackage stream).
    let encrypted =
        std::fs::read(fixture("inputs/example_password.xlsx")).expect("read encrypted fixture");
    let mut src = cfb::CompoundFile::open(Cursor::new(encrypted)).expect("open fixture cfb");
    let mut encryption_info = Vec::new();
    src.open_stream("EncryptionInfo")
        .expect("open EncryptionInfo stream")
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo stream");

    // Rewrap the stream contents, but store as `/EncryptionInfo`.
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_stream("/EncryptionInfo")
        .expect("create /EncryptionInfo")
        .write_all(&encryption_info)
        .expect("write /EncryptionInfo");
 
    let err = decrypt_standard_ooxml_from_bytes(ole.into_inner().into_inner(), "pw").unwrap_err();
    assert!(
        matches!(
            &err,
            OffcryptoError::UnsupportedEncryption {
                encryption_type: EncryptionType::Agile
            }
        ),
        "expected UnsupportedEncryption(Agile), got {err:?}"
    );
}

#[test]
fn missing_encryptedpackage_stream_returns_error() {
    // Ensure we surface a structured error (and never panic) when the OLE container is missing
    // `EncryptedPackage`.
    let encrypted =
        std::fs::read(fixture("inputs/ecma376standard_password.docx")).expect("read fixture");
    let mut ole_fixture = cfb::CompoundFile::open(Cursor::new(encrypted)).expect("open fixture cfb");
    let mut encryption_info = Vec::new();
    ole_fixture
        .open_stream("EncryptionInfo")
        .expect("open EncryptionInfo stream")
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo stream");

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_stream("EncryptionInfo")
        .expect("create EncryptionInfo stream")
        .write_all(&encryption_info)
        .expect("write EncryptionInfo stream");

    let err = decrypt_standard_ooxml_from_bytes(ole.into_inner().into_inner(), "Password1234_")
        .unwrap_err();
    assert!(
        matches!(
            &err,
            OffcryptoError::InvalidStructure(msg) if msg.contains("missing `EncryptedPackage` stream")
        ),
        "expected InvalidStructure missing EncryptedPackage, got {err:?}"
    );
}

#[test]
fn missing_encryptedpackage_stream_returns_error_even_with_wrong_password() {
    // Missing `EncryptedPackage` should be treated as a structural error regardless of password.
    // This also ensures we don't do expensive password derivation before confirming required
    // streams exist.
    let encrypted =
        std::fs::read(fixture("inputs/ecma376standard_password.docx")).expect("read fixture");
    let mut ole_fixture = cfb::CompoundFile::open(Cursor::new(encrypted)).expect("open fixture cfb");
    let mut encryption_info = Vec::new();
    ole_fixture
        .open_stream("EncryptionInfo")
        .expect("open EncryptionInfo stream")
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo stream");

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_stream("EncryptionInfo")
        .expect("create EncryptionInfo stream")
        .write_all(&encryption_info)
        .expect("write EncryptionInfo stream");

    let err = decrypt_standard_ooxml_from_bytes(ole.into_inner().into_inner(), "wrong-password")
        .unwrap_err();
    assert!(
        matches!(
            &err,
            OffcryptoError::InvalidStructure(msg) if msg.contains("missing `EncryptedPackage` stream")
        ),
        "expected InvalidStructure missing EncryptedPackage, got {err:?}"
    );
}

#[test]
fn supports_encryptedpackage_with_leading_slash_stream_name() {
    // Ensure we can read an absolute `/EncryptedPackage` stream path.
    let encrypted =
        std::fs::read(fixture("inputs/ecma376standard_password.docx")).expect("read fixture");
    let expected =
        std::fs::read(fixture("outputs/ecma376standard_password_plain.docx")).expect("read expected");

    let mut ole_fixture = cfb::CompoundFile::open(Cursor::new(encrypted)).expect("open fixture cfb");
    let mut encryption_info = Vec::new();
    ole_fixture
        .open_stream("EncryptionInfo")
        .expect("open EncryptionInfo stream")
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo stream");
    let mut encrypted_package = Vec::new();
    ole_fixture
        .open_stream("EncryptedPackage")
        .expect("open EncryptedPackage stream")
        .read_to_end(&mut encrypted_package)
        .expect("read EncryptedPackage stream");

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_stream("EncryptionInfo")
        .expect("create EncryptionInfo stream")
        .write_all(&encryption_info)
        .expect("write EncryptionInfo stream");
    ole.create_stream("/EncryptedPackage")
        .expect("create /EncryptedPackage stream")
        .write_all(&encrypted_package)
        .expect("write /EncryptedPackage stream");

    let decrypted = decrypt_standard_ooxml_from_bytes(ole.into_inner().into_inner(), "Password1234_")
        .expect("decrypt via /EncryptedPackage");
    assert_eq!(decrypted, expected);
}

#[test]
fn supports_case_insensitive_stream_names() {
    // Some producers may vary the casing of stream names. OLE/CFB stream lookup should be
    // best-effort and treat stream paths as case-insensitive.
    let encrypted =
        std::fs::read(fixture("inputs/ecma376standard_password.docx")).expect("read fixture");
    let expected =
        std::fs::read(fixture("outputs/ecma376standard_password_plain.docx")).expect("read expected");

    let mut ole_fixture = cfb::CompoundFile::open(Cursor::new(encrypted)).expect("open fixture cfb");
    let mut encryption_info = Vec::new();
    ole_fixture
        .open_stream("EncryptionInfo")
        .expect("open EncryptionInfo stream")
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo stream");
    let mut encrypted_package = Vec::new();
    ole_fixture
        .open_stream("EncryptedPackage")
        .expect("open EncryptedPackage stream")
        .read_to_end(&mut encrypted_package)
        .expect("read EncryptedPackage stream");

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_stream("encryptioninfo")
        .expect("create encryptioninfo stream")
        .write_all(&encryption_info)
        .expect("write encryptioninfo stream");
    ole.create_stream("encryptedpackage")
        .expect("create encryptedpackage stream")
        .write_all(&encrypted_package)
        .expect("write encryptedpackage stream");

    let decrypted = decrypt_standard_ooxml_from_bytes(ole.into_inner().into_inner(), "Password1234_")
        .expect("decrypt via case-insensitive stream names");
    assert_eq!(decrypted, expected);
}

#[test]
fn supports_standard_decrypt_with_slash_prefixed_stream_names() {
    // End-to-end Standard decrypt, but with both streams stored under absolute paths.
    let encrypted =
        std::fs::read(fixture("inputs/ecma376standard_password.docx")).expect("read fixture");
    let expected =
        std::fs::read(fixture("outputs/ecma376standard_password_plain.docx")).expect("read expected");

    let mut ole_fixture = cfb::CompoundFile::open(Cursor::new(encrypted)).expect("open fixture cfb");
    let mut encryption_info = Vec::new();
    ole_fixture
        .open_stream("EncryptionInfo")
        .expect("open EncryptionInfo stream")
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo stream");
    let mut encrypted_package = Vec::new();
    ole_fixture
        .open_stream("EncryptedPackage")
        .expect("open EncryptedPackage stream")
        .read_to_end(&mut encrypted_package)
        .expect("read EncryptedPackage stream");

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_stream("/EncryptionInfo")
        .expect("create /EncryptionInfo stream")
        .write_all(&encryption_info)
        .expect("write /EncryptionInfo stream");
    ole.create_stream("/EncryptedPackage")
        .expect("create /EncryptedPackage stream")
        .write_all(&encrypted_package)
        .expect("write /EncryptedPackage stream");

    let decrypted =
        decrypt_standard_ooxml_from_bytes(ole.into_inner().into_inner(), "Password1234_")
            .expect("decrypt via /EncryptionInfo + /EncryptedPackage");
    assert_eq!(decrypted, expected);
}
