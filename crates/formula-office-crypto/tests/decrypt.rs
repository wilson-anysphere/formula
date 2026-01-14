use std::io::{Cursor, Read as _, Seek as _, SeekFrom, Write as _};

use formula_office_crypto::{decrypt_encrypted_package, is_encrypted_ooxml_ole, OfficeCryptoError};

const AGILE_FIXTURE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../fixtures/encrypted/ooxml/agile-large.xlsx"
));
const AGILE_EMPTY_PASSWORD_FIXTURE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../fixtures/encrypted/ooxml/agile-empty-password.xlsx"
));
const AGILE_PLAINTEXT: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../fixtures/encrypted/ooxml/plaintext-large.xlsx"
));
const AGILE_UNICODE_FIXTURE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../fixtures/encrypted/ooxml/agile-unicode.xlsx"
));
const STANDARD_FIXTURE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../fixtures/encrypted/ooxml/standard.xlsx"
));
const STANDARD_4_2_FIXTURE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../fixtures/encrypted/ooxml/standard-4.2.xlsx"
));
const STANDARD_RC4_FIXTURE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../fixtures/encrypted/ooxml/standard-rc4.xlsx"
));
const STANDARD_LARGE_FIXTURE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../fixtures/encrypted/ooxml/standard-large.xlsx"
));
const STANDARD_PLAINTEXT: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../fixtures/encrypted/ooxml/plaintext.xlsx"
));
const STANDARD_BASIC_FIXTURE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../fixtures/encrypted/ooxml/standard-basic.xlsm"
));
const STANDARD_BASIC_PLAINTEXT: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../fixtures/encrypted/ooxml/plaintext-basic.xlsm"
));
const STANDARD_UNICODE_FIXTURE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../fixtures/encrypted/ooxml/standard-unicode.xlsx"
));
const AGILE_UNICODE_EXCEL_FIXTURE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../fixtures/encrypted/ooxml/agile-unicode-excel.xlsx"
));
const AGILE_UNICODE_EXCEL_PLAINTEXT: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../fixtures/encrypted/ooxml/plaintext-excel.xlsx"
));

fn assert_decrypted_zip_contains_workbook(decrypted: &[u8]) {
    assert!(
        decrypted.starts_with(b"PK"),
        "decrypted payload should start with ZIP magic"
    );

    let archive = zip::ZipArchive::new(Cursor::new(decrypted)).expect("open zip archive");
    let mut has_workbook = false;
    for name in archive.file_names() {
        if name.eq_ignore_ascii_case("xl/workbook.xml")
            || name.eq_ignore_ascii_case("xl/workbook.bin")
        {
            has_workbook = true;
            break;
        }
    }
    assert!(
        has_workbook,
        "expected decrypted ZIP to contain xl/workbook.*"
    );
}

#[test]
fn decrypts_agile_encrypted_package() {
    let decrypted = decrypt_encrypted_package(AGILE_FIXTURE, "password").expect("decrypt agile");
    assert_eq!(decrypted.as_slice(), AGILE_PLAINTEXT);
    assert_decrypted_zip_contains_workbook(&decrypted);
}

#[test]
fn decrypts_agile_empty_password_encrypted_package() {
    let decrypted = decrypt_encrypted_package(AGILE_EMPTY_PASSWORD_FIXTURE, "")
        .expect("decrypt agile (empty password)");
    assert_eq!(decrypted.as_slice(), STANDARD_PLAINTEXT);
    assert_decrypted_zip_contains_workbook(&decrypted);
}

#[test]
fn decrypts_agile_unicode_fixture() {
    let decrypted =
        decrypt_encrypted_package(AGILE_UNICODE_FIXTURE, "pässwörd").expect("decrypt agile unicode");
    assert_eq!(decrypted.as_slice(), STANDARD_PLAINTEXT);
    assert_decrypted_zip_contains_workbook(&decrypted);
}

#[test]
fn decrypts_standard_encrypted_package() {
    let decrypted =
        decrypt_encrypted_package(STANDARD_FIXTURE, "password").expect("decrypt standard");
    assert_eq!(decrypted.as_slice(), STANDARD_PLAINTEXT);
    assert_decrypted_zip_contains_workbook(&decrypted);
}

#[test]
fn decrypts_standard_4_2_encrypted_package() {
    // Some producers emit `EncryptionInfo` version 4.2 for Standard/CryptoAPI encryption
    // (still `versionMinor == 2`). Ensure we can decrypt Apache POI-produced files.
    let decrypted =
        decrypt_encrypted_package(STANDARD_4_2_FIXTURE, "password").expect("decrypt standard 4.2");
    assert_eq!(decrypted.as_slice(), STANDARD_PLAINTEXT);
    assert_decrypted_zip_contains_workbook(&decrypted);
}

#[test]
fn decrypts_standard_rc4_encrypted_package() {
    let decrypted =
        decrypt_encrypted_package(STANDARD_RC4_FIXTURE, "password").expect("decrypt standard rc4");
    assert_eq!(decrypted.as_slice(), STANDARD_PLAINTEXT);
    assert_decrypted_zip_contains_workbook(&decrypted);
}

#[test]
fn decrypts_standard_large_encrypted_package() {
    let decrypted =
        decrypt_encrypted_package(STANDARD_LARGE_FIXTURE, "password").expect("decrypt standard");
    assert_eq!(decrypted.as_slice(), AGILE_PLAINTEXT);
    assert_decrypted_zip_contains_workbook(&decrypted);
}

#[test]
fn decrypts_standard_basic_xlsm_fixture() {
    let decrypted =
        decrypt_encrypted_package(STANDARD_BASIC_FIXTURE, "password").expect("decrypt standard");
    assert_eq!(decrypted.as_slice(), STANDARD_BASIC_PLAINTEXT);
    assert_decrypted_zip_contains_workbook(&decrypted);
}

#[test]
fn decrypts_standard_encrypted_package_when_size_header_high_dword_is_reserved() {
    // Some producers treat the 8-byte EncryptedPackage size prefix as (u32 size, u32 reserved).
    // Mutate the high DWORD to a non-zero value and ensure decrypt still succeeds.
    let cursor = Cursor::new(STANDARD_FIXTURE.to_vec());
    let mut ole = cfb::CompoundFile::open(cursor).expect("open cfb");

    {
        let mut stream = ole
            .open_stream("EncryptedPackage")
            .or_else(|_| ole.open_stream("/EncryptedPackage"))
            .expect("open EncryptedPackage");

        let mut header = [0u8; 8];
        stream.read_exact(&mut header).expect("read size prefix");
        header[4..8].copy_from_slice(&1u32.to_le_bytes());
        stream
            .seek(SeekFrom::Start(0))
            .expect("seek EncryptedPackage to start");
        stream.write_all(&header).expect("write size prefix");
    }

    let ole_bytes = ole.into_inner().into_inner();
    let decrypted = decrypt_encrypted_package(&ole_bytes, "password").expect("decrypt standard");
    assert_eq!(decrypted.as_slice(), STANDARD_PLAINTEXT);
    assert_decrypted_zip_contains_workbook(&decrypted);
}

#[test]
fn decrypts_standard_rc4_encrypted_package_when_size_header_high_dword_is_reserved() {
    // Same as `decrypts_standard_encrypted_package_when_size_header_high_dword_is_reserved`, but
    // against the Standard/CryptoAPI RC4 fixture to ensure the compatibility fallback is exercised
    // across cipher variants.
    let cursor = Cursor::new(STANDARD_RC4_FIXTURE.to_vec());
    let mut ole = cfb::CompoundFile::open(cursor).expect("open cfb");

    {
        let mut stream = ole
            .open_stream("EncryptedPackage")
            .or_else(|_| ole.open_stream("/EncryptedPackage"))
            .expect("open EncryptedPackage");

        let mut header = [0u8; 8];
        stream.read_exact(&mut header).expect("read size prefix");
        header[4..8].copy_from_slice(&1u32.to_le_bytes());
        stream
            .seek(SeekFrom::Start(0))
            .expect("seek EncryptedPackage to start");
        stream.write_all(&header).expect("write size prefix");
    }

    let ole_bytes = ole.into_inner().into_inner();
    let decrypted = decrypt_encrypted_package(&ole_bytes, "password").expect("decrypt standard rc4");
    assert_eq!(decrypted.as_slice(), STANDARD_PLAINTEXT);
    assert_decrypted_zip_contains_workbook(&decrypted);
}

#[test]
fn decrypts_standard_rc4_with_case_insensitive_stream_names() {
    // Some OLE writers vary stream name casing; ensure we can still open/decrypt RC4 fixtures via
    // the public `decrypt_encrypted_package()` convenience API.
    let cursor = Cursor::new(STANDARD_RC4_FIXTURE.to_vec());
    let mut ole_in = cfb::CompoundFile::open(cursor).expect("open input cfb");

    let mut encryption_info = Vec::new();
    ole_in
        .open_stream("EncryptionInfo")
        .or_else(|_| ole_in.open_stream("/EncryptionInfo"))
        .expect("open EncryptionInfo")
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo");

    let mut encrypted_package = Vec::new();
    ole_in
        .open_stream("EncryptedPackage")
        .or_else(|_| ole_in.open_stream("/EncryptedPackage"))
        .expect("open EncryptedPackage")
        .read_to_end(&mut encrypted_package)
        .expect("read EncryptedPackage");

    let cursor = Cursor::new(Vec::new());
    let mut ole_out = cfb::CompoundFile::create(cursor).expect("create output cfb");
    ole_out
        .create_stream("encryptioninfo")
        .expect("create encryptioninfo")
        .write_all(&encryption_info)
        .expect("write encryptioninfo");
    ole_out
        .create_stream("encryptedpackage")
        .expect("create encryptedpackage")
        .write_all(&encrypted_package)
        .expect("write encryptedpackage");
    let ole_bytes = ole_out.into_inner().into_inner();

    let decrypted =
        decrypt_encrypted_package(&ole_bytes, "password").expect("decrypt standard rc4");
    assert_eq!(decrypted.as_slice(), STANDARD_PLAINTEXT);
    assert_decrypted_zip_contains_workbook(&decrypted);
}

#[test]
fn decrypts_standard_rc4_with_leading_slash_and_case_variation_stream_names() {
    // Exercise combined quirks:
    // - leading `/` in root stream names
    // - case variation
    //
    // This ensures `open_stream_case_tolerant()` can find the streams via its walk()/case-insensitive
    // fallback even when the naive `/EncryptionInfo` probe doesn't match the casing.
    let cursor = Cursor::new(STANDARD_RC4_FIXTURE.to_vec());
    let mut ole_in = cfb::CompoundFile::open(cursor).expect("open input cfb");

    let mut encryption_info = Vec::new();
    ole_in
        .open_stream("EncryptionInfo")
        .or_else(|_| ole_in.open_stream("/EncryptionInfo"))
        .expect("open EncryptionInfo")
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo");

    let mut encrypted_package = Vec::new();
    ole_in
        .open_stream("EncryptedPackage")
        .or_else(|_| ole_in.open_stream("/EncryptedPackage"))
        .expect("open EncryptedPackage")
        .read_to_end(&mut encrypted_package)
        .expect("read EncryptedPackage");

    let cursor = Cursor::new(Vec::new());
    let mut ole_out = cfb::CompoundFile::create(cursor).expect("create output cfb");
    ole_out
        .create_stream("/encryptioninfo")
        .expect("create /encryptioninfo")
        .write_all(&encryption_info)
        .expect("write /encryptioninfo");
    ole_out
        .create_stream("/encryptedpackage")
        .expect("create /encryptedpackage")
        .write_all(&encrypted_package)
        .expect("write /encryptedpackage");
    let ole_bytes = ole_out.into_inner().into_inner();

    assert!(is_encrypted_ooxml_ole(&ole_bytes));

    let decrypted =
        decrypt_encrypted_package(&ole_bytes, "password").expect("decrypt standard rc4");
    assert_eq!(decrypted.as_slice(), STANDARD_PLAINTEXT);
    assert_decrypted_zip_contains_workbook(&decrypted);
}

#[test]
fn standard_rc4_truncated_encrypted_package_returns_invalid_format() {
    // Ensure we treat truncated EncryptedPackage ciphertext as an InvalidFormat error (after the
    // password verifier succeeds), not as InvalidPassword.
    let cursor = Cursor::new(STANDARD_RC4_FIXTURE.to_vec());
    let mut ole = cfb::CompoundFile::open(cursor).expect("open cfb");

    let mut encryption_info = Vec::new();
    ole.open_stream("EncryptionInfo")
        .or_else(|_| ole.open_stream("/EncryptionInfo"))
        .expect("open EncryptionInfo")
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo");

    let mut encrypted_package = Vec::new();
    ole.open_stream("EncryptedPackage")
        .or_else(|_| ole.open_stream("/EncryptedPackage"))
        .expect("open EncryptedPackage")
        .read_to_end(&mut encrypted_package)
        .expect("read EncryptedPackage");

    // Keep only the 8-byte plaintext size prefix (drop all ciphertext).
    encrypted_package.truncate(8);

    let cursor = Cursor::new(Vec::new());
    let mut ole_out = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole_out
        .create_stream("EncryptionInfo")
        .expect("create EncryptionInfo")
        .write_all(&encryption_info)
        .expect("write EncryptionInfo");
    ole_out
        .create_stream("EncryptedPackage")
        .expect("create EncryptedPackage")
        .write_all(&encrypted_package)
        .expect("write EncryptedPackage");

    let ole_bytes = ole_out.into_inner().into_inner();
    let err = decrypt_encrypted_package(&ole_bytes, "password").expect_err("expected error");
    match err {
        OfficeCryptoError::InvalidFormat(msg) => {
            assert!(
                msg.to_ascii_lowercase().contains("ciphertext"),
                "expected error message to mention ciphertext, got: {msg}"
            );
        }
        other => panic!("expected InvalidFormat, got {other:?}"),
    }
}

#[test]
fn standard_rc4_unsupported_algidhash_returns_unsupported_encryption() {
    // Requirement: Unsupported `algIdHash` values should return UnsupportedEncryption (not
    // InvalidPassword).
    //
    // Mutate the Standard/CryptoAPI RC4 fixture to use an unknown AlgIDHash and ensure the error
    // is surfaced as UnsupportedEncryption.
    let cursor = Cursor::new(STANDARD_RC4_FIXTURE.to_vec());
    let mut ole = cfb::CompoundFile::open(cursor).expect("open cfb");

    {
        let mut stream = ole
            .open_stream("EncryptionInfo")
            .or_else(|_| ole.open_stream("/EncryptionInfo"))
            .expect("open EncryptionInfo");

        // Standard EncryptionInfo layout:
        // - 8-byte version header
        // - 4-byte headerSize
        // - EncryptionHeader (starts at offset 12)
        // Within EncryptionHeader, algIdHash is a DWORD at offset 12.
        let alg_id_hash_offset = 12u64 + 12u64;
        stream
            .seek(SeekFrom::Start(alg_id_hash_offset))
            .expect("seek to algIdHash");
        stream
            .write_all(&0xDEAD_BEEFu32.to_le_bytes())
            .expect("write algIdHash");
    }

    let ole_bytes = ole.into_inner().into_inner();
    let err = decrypt_encrypted_package(&ole_bytes, "password").expect_err("expected failure");
    match err {
        OfficeCryptoError::UnsupportedEncryption(msg) => {
            assert!(
                msg.contains("AlgIDHash"),
                "expected UnsupportedEncryption message to mention AlgIDHash, got: {msg}"
            );
        }
        other => panic!("expected UnsupportedEncryption, got {other:?}"),
    }
}

#[test]
fn wrong_password_returns_invalid_password() {
    let err = decrypt_encrypted_package(AGILE_FIXTURE, "wrong").expect_err("expected error");
    assert!(
        matches!(err, OfficeCryptoError::InvalidPassword),
        "expected InvalidPassword, got {err:?}"
    );

    let err = decrypt_encrypted_package(STANDARD_FIXTURE, "wrong").expect_err("expected error");
    assert!(
        matches!(err, OfficeCryptoError::InvalidPassword),
        "expected InvalidPassword, got {err:?}"
    );

    let err = decrypt_encrypted_package(STANDARD_4_2_FIXTURE, "wrong").expect_err("expected error");
    assert!(
        matches!(err, OfficeCryptoError::InvalidPassword),
        "expected InvalidPassword, got {err:?}"
    );

    let err = decrypt_encrypted_package(STANDARD_RC4_FIXTURE, "wrong").expect_err("expected error");
    assert!(
        matches!(err, OfficeCryptoError::InvalidPassword),
        "expected InvalidPassword, got {err:?}"
    );

    let err = decrypt_encrypted_package(STANDARD_LARGE_FIXTURE, "wrong").expect_err("expected error");
    assert!(
        matches!(err, OfficeCryptoError::InvalidPassword),
        "expected InvalidPassword, got {err:?}"
    );

    let err =
        decrypt_encrypted_package(AGILE_UNICODE_FIXTURE, "wrong").expect_err("expected error");
    assert!(
        matches!(err, OfficeCryptoError::InvalidPassword),
        "expected InvalidPassword, got {err:?}"
    );

    let err =
        decrypt_encrypted_package(STANDARD_BASIC_FIXTURE, "wrong").expect_err("expected error");
    assert!(
        matches!(err, OfficeCryptoError::InvalidPassword),
        "expected InvalidPassword, got {err:?}"
    );

    let err =
        decrypt_encrypted_package(AGILE_UNICODE_EXCEL_FIXTURE, "wrong").expect_err("expected error");
    assert!(
        matches!(err, OfficeCryptoError::InvalidPassword),
        "expected InvalidPassword, got {err:?}"
    );

    let err =
        decrypt_encrypted_package(STANDARD_UNICODE_FIXTURE, "wrong").expect_err("expected error");
    assert!(
        matches!(err, OfficeCryptoError::InvalidPassword),
        "expected InvalidPassword, got {err:?}"
    );
}

#[test]
fn decrypts_agile_unicode_excel_fixture() {
    let decrypted = decrypt_encrypted_package(AGILE_UNICODE_EXCEL_FIXTURE, "pässwörd🔒")
        .expect("decrypt agile");
    assert_eq!(decrypted.as_slice(), AGILE_UNICODE_EXCEL_PLAINTEXT);
    assert_decrypted_zip_contains_workbook(&decrypted);
}

#[test]
fn agile_unicode_excel_password_different_normalization_fails() {
    // NFC password is "pässwörd🔒" (U+00E4, U+00F6). NFD decomposes those into combining marks, but
    // leaves the non-BMP emoji alone.
    let nfd = "pa\u{0308}sswo\u{0308}rd🔒";
    assert_ne!(nfd, "pässwörd🔒");

    let err = decrypt_encrypted_package(AGILE_UNICODE_EXCEL_FIXTURE, nfd)
        .expect_err("expected different normalization to fail");
    assert!(
        matches!(err, OfficeCryptoError::InvalidPassword),
        "expected InvalidPassword, got {err:?}"
    );
}

#[test]
fn decrypts_standard_unicode_fixture() {
    let decrypted = decrypt_encrypted_package(STANDARD_UNICODE_FIXTURE, "pässwörd🔒")
        .expect("decrypt standard");
    assert_eq!(decrypted.as_slice(), STANDARD_PLAINTEXT);
    assert_decrypted_zip_contains_workbook(&decrypted);
}

#[test]
fn standard_unicode_password_different_normalization_fails() {
    // NFC password is "pässwörd🔒" (U+00E4, U+00F6). NFD decomposes those into combining marks, but
    // leaves the non-BMP emoji alone.
    let nfd = "pa\u{0308}sswo\u{0308}rd🔒";
    assert_ne!(nfd, "pässwörd🔒");

    let err = decrypt_encrypted_package(STANDARD_UNICODE_FIXTURE, nfd)
        .expect_err("expected different normalization to fail");
    assert!(
        matches!(err, OfficeCryptoError::InvalidPassword),
        "expected InvalidPassword, got {err:?}"
    );
}
