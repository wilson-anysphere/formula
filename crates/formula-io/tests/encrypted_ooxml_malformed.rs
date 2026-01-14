use std::io::{Cursor, Read, Write};
use std::path::Path;

use formula_io::{
    open_workbook, open_workbook_model, open_workbook_model_with_password, open_workbook_with_options,
    open_workbook_with_password, OpenOptions,
};

fn build_encrypted_ooxml_container(
    encryption_info: &[u8],
    encrypted_package: Option<&[u8]>,
) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    {
        let mut stream = ole
            .create_stream("EncryptionInfo")
            .expect("create EncryptionInfo stream");
        stream
            .write_all(encryption_info)
            .expect("write EncryptionInfo bytes");
    }

    if let Some(bytes) = encrypted_package {
        let mut stream = ole
            .create_stream("EncryptedPackage")
            .expect("create EncryptedPackage stream");
        stream
            .write_all(bytes)
            .expect("write EncryptedPackage bytes");
    }

    ole.into_inner().into_inner()
}

fn encrypted_package_with_size_prefix(decrypted_size: u64, payload: &[u8]) -> Vec<u8> {
    let mut out = Vec::with_capacity(8 + payload.len());
    out.extend_from_slice(&decrypted_size.to_le_bytes());
    out.extend_from_slice(payload);
    out
}

#[test]
fn standard_streaming_open_rejects_invalid_encrypted_package_size_prefix() {
    // `open_workbook_with_options` prefers the Standard AES streaming open path for Standard-encrypted
    // workbooks. If the EncryptedPackage size prefix is inconsistent with the ciphertext length,
    // treat it as an unsupported/malformed encryption container (not a generic `.xlsx` open error).
    let fixture_path = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encrypted/ooxml")
        .join("standard.xlsx");
    let fixture_bytes = std::fs::read(&fixture_path).expect("read standard.xlsx fixture");
    let mut ole = cfb::CompoundFile::open(Cursor::new(fixture_bytes)).expect("open fixture cfb");

    let mut encryption_info = Vec::new();
    ole.open_stream("EncryptionInfo")
        .expect("EncryptionInfo stream")
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo");

    let mut encrypted_package = Vec::new();
    ole.open_stream("EncryptedPackage")
        .expect("EncryptedPackage stream")
        .read_to_end(&mut encrypted_package)
        .expect("read EncryptedPackage");

    assert!(encrypted_package.len() > 8, "fixture EncryptedPackage too small");
    let ciphertext_len = encrypted_package.len() - 8;
    let bogus_plaintext_len = ciphertext_len as u64 + 1;
    encrypted_package[..8].copy_from_slice(&bogus_plaintext_len.to_le_bytes());

    let bytes = build_encrypted_ooxml_container(&encryption_info, Some(&encrypted_package));
    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("corrupt-standard.xlsx");
    std::fs::write(&path, bytes).expect("write corrupt fixture");

    let err = open_workbook_with_options(
        &path,
        OpenOptions {
            password: Some("password".to_string()),
        },
    )
    .expect_err("expected corrupt encrypted workbook to error");
    assert!(
        matches!(err, formula_io::Error::UnsupportedOoxmlEncryption { .. }),
        "expected UnsupportedOoxmlEncryption, got {err:?}"
    );
}

#[test]
fn encrypted_package_size_prefix_only_is_unsupported_ooxml_encryption() {
    // If `EncryptedPackage` only contains the 8-byte size prefix (and no ciphertext payload),
    // treat the workbook as a malformed/unsupported encrypted container rather than as a wrong
    // password.
    let fixture_path = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encrypted/ooxml")
        .join("standard.xlsx");
    let fixture_bytes = std::fs::read(&fixture_path).expect("read standard.xlsx fixture");
    let mut ole = cfb::CompoundFile::open(Cursor::new(fixture_bytes)).expect("open fixture cfb");

    let mut encryption_info = Vec::new();
    ole.open_stream("EncryptionInfo")
        .expect("EncryptionInfo stream")
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo");

    let encrypted_package = 0u64.to_le_bytes();
    let bytes = build_encrypted_ooxml_container(&encryption_info, Some(&encrypted_package));

    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("size-prefix-only.xlsx");
    std::fs::write(&path, bytes).expect("write fixture");

    let err = open_workbook_with_password(&path, Some("password")).expect_err("expected error");
    assert!(
        matches!(err, formula_io::Error::UnsupportedOoxmlEncryption { .. }),
        "expected UnsupportedOoxmlEncryption, got {err:?}"
    );
}

fn assert_err_no_panic<T: std::fmt::Debug>(
    label: &str,
    res: std::thread::Result<Result<T, formula_io::Error>>,
) {
    assert!(
        res.is_ok(),
        "{label} panicked while opening malformed encrypted OOXML container"
    );
    let err = res
        .unwrap()
        .expect_err("expected malformed encrypted workbook to error");
    let msg = err.to_string().to_lowercase();
    assert!(
        msg.contains("encrypt") || msg.contains("password"),
        "expected error message to mention encryption/password, got: {msg}"
    );
}

fn assert_open_errors_no_panic(path: &Path) {
    // Without a password, callers should get a "password required" style error.
    assert_err_no_panic(
        "open_workbook",
        std::panic::catch_unwind(|| open_workbook(path)),
    );
    assert_err_no_panic(
        "open_workbook_model",
        std::panic::catch_unwind(|| open_workbook_model(path)),
    );

    // With a password, the decrypt/open path should still fail gracefully (not panic) on malformed
    // encryption metadata.
    assert_err_no_panic(
        "open_workbook_with_password",
        std::panic::catch_unwind(|| open_workbook_with_password(path, Some("password"))),
    );
    assert_err_no_panic(
        "open_workbook_model_with_password",
        std::panic::catch_unwind(|| open_workbook_model_with_password(path, Some("password"))),
    );
}

#[test]
fn agile_header_with_truncated_xml_does_not_panic() {
    // EncryptionInfo (Agile) format begins with:
    // - major (u16 LE) = 4
    // - minor (u16 LE) = 4
    // - flags (u32 LE) = 0
    // - followed by an XML descriptor (UTF-8), which we intentionally truncate.
    let mut encryption_info = Vec::new();
    encryption_info.extend_from_slice(&4u16.to_le_bytes()); // major
    encryption_info.extend_from_slice(&4u16.to_le_bytes()); // minor
    encryption_info.extend_from_slice(&0u32.to_le_bytes()); // flags
    encryption_info.extend_from_slice(
        br#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption""#,
    );

    let encrypted_package = encrypted_package_with_size_prefix(16, b"deadbeef");
    let bytes = build_encrypted_ooxml_container(&encryption_info, Some(&encrypted_package));

    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("malformed-agile.xlsx");
    std::fs::write(&path, bytes).expect("write fixture");

    assert_open_errors_no_panic(&path);
}

#[test]
fn standard_header_with_truncated_encryption_header_does_not_panic() {
    // EncryptionInfo (Standard/CryptoAPI) begins with:
    // - major (u16 LE) = 3
    // - minor (u16 LE) = 2
    // - flags (u32 LE) indicating CryptoAPI/AES (we set a plausible value)
    // - EncryptionHeaderSize (u32 LE) followed by that many bytes of EncryptionHeader
    // We intentionally declare a header size larger than the remaining data to simulate truncation.
    let mut encryption_info = Vec::new();
    encryption_info.extend_from_slice(&3u16.to_le_bytes()); // major
    encryption_info.extend_from_slice(&2u16.to_le_bytes()); // minor
    encryption_info.extend_from_slice(&0x24u32.to_le_bytes()); // flags (fCryptoAPI | fAES)
    encryption_info.extend_from_slice(&32u32.to_le_bytes()); // EncryptionHeaderSize
    encryption_info.extend_from_slice(&[0u8; 8]); // truncated header payload (should be 32)

    let encrypted_package = encrypted_package_with_size_prefix(16, b"deadbeef");
    let bytes = build_encrypted_ooxml_container(&encryption_info, Some(&encrypted_package));

    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("malformed-standard.xlsx");
    std::fs::write(&path, bytes).expect("write fixture");

    assert_open_errors_no_panic(&path);
}

#[test]
fn encrypted_package_missing_or_empty_does_not_panic() {
    let tmp = tempfile::tempdir().expect("tempdir");

    // Case: Agile EncryptionInfo that looks like a (very) minimal XML descriptor, but the
    // EncryptedPackage stream is present and empty.
    let mut agile_info = Vec::new();
    agile_info.extend_from_slice(&4u16.to_le_bytes()); // major
    agile_info.extend_from_slice(&4u16.to_le_bytes()); // minor
    agile_info.extend_from_slice(&0u32.to_le_bytes()); // flags
    agile_info.extend_from_slice(
        br#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"></encryption>"#,
    );
    let bytes = build_encrypted_ooxml_container(&agile_info, Some(&[]));
    let path = tmp.path().join("missing-package-agile.xlsx");
    std::fs::write(&path, bytes).expect("write fixture");
    assert_open_errors_no_panic(&path);

    // Case: Standard EncryptionInfo with a minimally structured header+verifier, but the
    // EncryptedPackage stream is present and empty.
    let mut standard_info = Vec::new();
    standard_info.extend_from_slice(&3u16.to_le_bytes()); // major
    standard_info.extend_from_slice(&2u16.to_le_bytes()); // minor
    standard_info.extend_from_slice(&0x24u32.to_le_bytes()); // flags (fCryptoAPI | fAES)
    // EncryptionHeaderSize (u32) + EncryptionHeader (32 bytes)
    standard_info.extend_from_slice(&32u32.to_le_bytes());
    // EncryptionHeader: 8 * u32 fields, no CSPName (header size exactly 32)
    standard_info.extend_from_slice(&0u32.to_le_bytes()); // flags
    standard_info.extend_from_slice(&0u32.to_le_bytes()); // sizeExtra
    standard_info.extend_from_slice(&0x0000_660Eu32.to_le_bytes()); // algId (AES-128)
    standard_info.extend_from_slice(&0x0000_8004u32.to_le_bytes()); // algIdHash (SHA-1)
    standard_info.extend_from_slice(&128u32.to_le_bytes()); // keySize (bits)
    standard_info.extend_from_slice(&24u32.to_le_bytes()); // providerType (PROV_RSA_AES)
    standard_info.extend_from_slice(&0u32.to_le_bytes()); // reserved1
    standard_info.extend_from_slice(&0u32.to_le_bytes()); // reserved2
    // EncryptionVerifier
    standard_info.extend_from_slice(&16u32.to_le_bytes()); // saltSize
    standard_info.extend_from_slice(&[0u8; 16]); // salt
    standard_info.extend_from_slice(&[0u8; 16]); // encryptedVerifier
    standard_info.extend_from_slice(&20u32.to_le_bytes()); // verifierHashSize (SHA-1)
    standard_info.extend_from_slice(&[0u8; 20]); // encryptedVerifierHash

    let bytes = build_encrypted_ooxml_container(&standard_info, Some(&[]));
    let path = tmp.path().join("missing-package-standard.xlsx");
    std::fs::write(&path, bytes).expect("write fixture");
    assert_open_errors_no_panic(&path);
}
