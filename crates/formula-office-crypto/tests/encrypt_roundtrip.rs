use formula_office_crypto::{
    decrypt_encrypted_package_ole, encrypt_package_to_ole, EncryptOptions, EncryptionScheme,
    HashAlgorithm, OfficeCryptoError,
};
use std::io::{Cursor, Read, Write};
use std::sync::OnceLock;

const AGILE_PASSWORD: &str = "correct horse battery staple";
const FAST_TEST_SPIN_COUNT: u32 = 10_000;

fn basic_xlsx_fixture_bytes() -> &'static [u8] {
    static BYTES: OnceLock<Vec<u8>> = OnceLock::new();
    BYTES
        .get_or_init(|| {
            let path = std::path::Path::new(concat!(
                env!("CARGO_MANIFEST_DIR"),
                "/../../fixtures/xlsx/basic/basic.xlsx"
            ));
            std::fs::read(path).expect("read basic.xlsx fixture")
        })
        .as_slice()
}

fn agile_ole_fixture() -> &'static [u8] {
    static OLE: OnceLock<Vec<u8>> = OnceLock::new();
    OLE.get_or_init(|| {
        encrypt_package_to_ole(
            basic_xlsx_fixture_bytes(),
            AGILE_PASSWORD,
            EncryptOptions {
                spin_count: FAST_TEST_SPIN_COUNT,
                ..Default::default()
            },
        )
        .expect("encrypt")
    })
    .as_slice()
}

fn read_encrypted_ooxml_streams(ole: &[u8]) -> (Vec<u8>, Vec<u8>) {
    let cursor = Cursor::new(ole);
    let mut ole_in = cfb::CompoundFile::open(cursor).expect("open cfb");

    let mut encryption_info = Vec::new();
    ole_in
        .open_stream("EncryptionInfo")
        .expect("open EncryptionInfo")
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo");

    let mut encrypted_package = Vec::new();
    ole_in
        .open_stream("EncryptedPackage")
        .expect("open EncryptedPackage")
        .read_to_end(&mut encrypted_package)
        .expect("read EncryptedPackage");

    (encryption_info, encrypted_package)
}

fn build_encrypted_ooxml_ole(encryption_info: &[u8], encrypted_package: &[u8]) -> Vec<u8> {
    let cursor_out = Cursor::new(Vec::new());
    let mut ole_out = cfb::CompoundFile::create(cursor_out).expect("create cfb");
    ole_out
        .create_stream("EncryptionInfo")
        .expect("create EncryptionInfo")
        .write_all(encryption_info)
        .expect("write EncryptionInfo");
    ole_out
        .create_stream("EncryptedPackage")
        .expect("create EncryptedPackage")
        .write_all(encrypted_package)
        .expect("write EncryptedPackage");

    ole_out.into_inner().into_inner()
}

fn assert_agile_integrity_check_failed(ole: &[u8]) {
    let err = decrypt_encrypted_package_ole(ole, AGILE_PASSWORD).expect_err("expected failure");
    assert!(
        matches!(err, OfficeCryptoError::IntegrityCheckFailed),
        "expected IntegrityCheckFailed, got {err:?}"
    );
}

fn agile_encryption_info_xml_range(encryption_info: &[u8]) -> std::ops::Range<usize> {
    assert!(encryption_info.len() >= 8, "EncryptionInfo stream too short");
    let version_major = u16::from_le_bytes([encryption_info[0], encryption_info[1]]);
    let version_minor = u16::from_le_bytes([encryption_info[2], encryption_info[3]]);
    assert_eq!(
        (version_major, version_minor),
        (4, 4),
        "expected Agile EncryptionInfo version 4.4"
    );

    // Replicate the crate's tolerant XML offset detection (some producers include an explicit XML
    // length field after the 8-byte header).
    let (xml_offset, xml_len) = if encryption_info.len() >= 12 {
        let candidate = u32::from_le_bytes([
            encryption_info[8],
            encryption_info[9],
            encryption_info[10],
            encryption_info[11],
        ]) as usize;
        let available = encryption_info.len().saturating_sub(12);
        if candidate <= available {
            (12usize, candidate)
        } else {
            (8usize, encryption_info.len().saturating_sub(8))
        }
    } else {
        (8usize, encryption_info.len().saturating_sub(8))
    };

    let xml_end = xml_offset + xml_len;
    assert!(encryption_info.len() >= xml_end, "EncryptionInfo XML out of range");

    xml_offset..xml_end
}

fn flip_base64_attr_first_char(xml_bytes: &mut [u8], needle: &[u8]) {
    let pos = xml_bytes
        .windows(needle.len())
        .position(|w| w == needle)
        .expect("attribute not found");
    let value_start = pos + needle.len();
    let value_end = value_start
        + xml_bytes[value_start..]
            .iter()
            .position(|&b| b == b'"')
            .expect("unterminated attribute value");
    assert!(value_end > value_start, "attribute value unexpectedly empty");

    xml_bytes[value_start] = if xml_bytes[value_start] != b'A' {
        b'A'
    } else {
        b'B'
    };
}

#[test]
fn agile_encrypt_decrypt_round_trip() {
    let zip = basic_xlsx_fixture_bytes();
    let ole = agile_ole_fixture();
    let decrypted = decrypt_encrypted_package_ole(ole, AGILE_PASSWORD).expect("decrypt");
    assert_eq!(decrypted.as_slice(), zip);
}

#[test]
fn wrong_password_fails() {
    let ole = agile_ole_fixture();

    let err =
        decrypt_encrypted_package_ole(ole, "not-the-password").expect_err("expected failure");
    assert!(
        matches!(err, OfficeCryptoError::InvalidPassword),
        "expected InvalidPassword, got {err:?}"
    );
}

#[test]
fn standard_encrypt_decrypt_round_trip() {
    let zip = basic_xlsx_fixture_bytes();
    let password = "swordfish";

    // Standard encryption supports AES-128/192/256. Exercise all key sizes to ensure the
    // CryptoAPI `CryptDeriveKey` expansion path works correctly for key lengths > SHA1 digest
    // length (AES-192/256).
    for key_bits in [128usize, 192, 256] {
        let ole = encrypt_package_to_ole(
            zip,
            password,
            EncryptOptions {
                scheme: EncryptionScheme::Standard,
                key_bits,
                hash_algorithm: HashAlgorithm::Sha1,
                // Standard uses a fixed 50k spin count internally (CryptoAPI), but keep the option
                // explicit so callers don't accidentally rely on Agile defaults.
                spin_count: 50_000,
            },
        )
        .expect("encrypt");
        let decrypted = decrypt_encrypted_package_ole(&ole, password).expect("decrypt");
        assert_eq!(decrypted.as_slice(), zip, "key_bits={key_bits}");

        // Wrong password should fail for all Standard AES key sizes.
        let err =
            decrypt_encrypted_package_ole(&ole, "wrong-password").expect_err("wrong password");
        assert!(
            matches!(err, OfficeCryptoError::InvalidPassword),
            "expected InvalidPassword for key_bits={key_bits}, got {err:?}"
        );
    }
}

#[test]
fn standard_unicode_and_whitespace_password_round_trip() {
    let zip = basic_xlsx_fixture_bytes();
    // Trailing whitespace is significant, and the emoji exercises non-BMP UTF-16 surrogate pairs.
    let password = "pässwörd🔒 ";

    let ole = encrypt_package_to_ole(
        zip,
        password,
        EncryptOptions {
            scheme: EncryptionScheme::Standard,
            key_bits: 128,
            hash_algorithm: HashAlgorithm::Sha1,
            spin_count: 50_000,
        },
    )
    .expect("encrypt");
    let decrypted = decrypt_encrypted_package_ole(&ole, password).expect("decrypt");
    assert_eq!(decrypted.as_slice(), zip);

    // Passwords must not be trimmed: the trimmed password should fail.
    let trimmed = password.trim();
    let err = decrypt_encrypted_package_ole(&ole, trimmed).expect_err("expected trimmed password to fail");
    assert!(
        matches!(err, OfficeCryptoError::InvalidPassword),
        "expected InvalidPassword for trimmed password, got {err:?}"
    );

    // Passwords must not be Unicode-normalized: NFC vs NFD should differ.
    let nfd = "pa\u{0308}sswo\u{0308}rd🔒 ";
    let err = decrypt_encrypted_package_ole(&ole, nfd).expect_err("expected NFD password to fail");
    assert!(
        matches!(err, OfficeCryptoError::InvalidPassword),
        "expected InvalidPassword for NFD password, got {err:?}"
    );
}

#[test]
fn tampered_ciphertext_fails_integrity_check() {
    let ole = agile_ole_fixture();

    let (encryption_info, mut encrypted_package) = read_encrypted_ooxml_streams(ole);
    assert!(
        encrypted_package.len() > 8,
        "EncryptedPackage should contain a size prefix and ciphertext"
    );
    encrypted_package[8] ^= 0x01; // Flip a byte in the ciphertext (not the length prefix).

    let tampered = build_encrypted_ooxml_ole(&encryption_info, &encrypted_package);
    assert_agile_integrity_check_failed(&tampered);
}

#[test]
fn tampered_size_prefix_fails_integrity_check() {
    let ole = agile_ole_fixture();

    let (encryption_info, mut encrypted_package) = read_encrypted_ooxml_streams(ole);
    assert!(
        encrypted_package.len() >= 8,
        "EncryptedPackage should contain an 8-byte size prefix"
    );
    let mut size_bytes = [0u8; 8];
    size_bytes.copy_from_slice(&encrypted_package[..8]);
    let size = u64::from_le_bytes(size_bytes);
    let tampered_size = size.saturating_sub(1);
    encrypted_package[..8].copy_from_slice(&tampered_size.to_le_bytes());

    let tampered = build_encrypted_ooxml_ole(&encryption_info, &encrypted_package);
    assert_agile_integrity_check_failed(&tampered);
}

#[test]
fn tampered_size_prefix_high_dword_fails_integrity_check() {
    let ole = agile_ole_fixture();

    // Some producers treat the 8-byte EncryptedPackage size prefix as `(u32 size, u32 reserved)`.
    // Mutate only the high DWORD without recomputing dataIntegrity; the decryptor should still
    // parse the size using the low DWORD, but the HMAC should fail because the stream bytes changed.
    let (encryption_info, mut encrypted_package) = read_encrypted_ooxml_streams(ole);
    assert!(
        encrypted_package.len() >= 8,
        "EncryptedPackage should contain an 8-byte size prefix"
    );
    // Set the reserved high DWORD to a non-zero value.
    encrypted_package[4..8].copy_from_slice(&1u32.to_le_bytes());

    let tampered = build_encrypted_ooxml_ole(&encryption_info, &encrypted_package);
    assert_agile_integrity_check_failed(&tampered);
}

#[test]
fn tampered_encrypted_hmac_value_fails_integrity_check() {
    let ole = agile_ole_fixture();

    // Extract the streams and mutate `dataIntegrity/encryptedHmacValue` inside EncryptionInfo,
    // leaving the EncryptedPackage stream unchanged. This should be detected as an integrity error.
    let (mut encryption_info, encrypted_package) = read_encrypted_ooxml_streams(ole);
    let xml_range = agile_encryption_info_xml_range(&encryption_info);
    let xml_bytes = &mut encryption_info[xml_range];
    flip_base64_attr_first_char(xml_bytes, b"encryptedHmacValue=\"");

    let tampered = build_encrypted_ooxml_ole(&encryption_info, &encrypted_package);
    assert_agile_integrity_check_failed(&tampered);
}

#[test]
fn appended_trailing_bytes_fail_integrity_check() {
    let ole = agile_ole_fixture();

    // Append trailing bytes to the EncryptedPackage stream. The decrypter may ignore them for
    // plaintext recovery, but MS-OFFCRYPTO dataIntegrity HMAC covers the EncryptedPackage stream
    // bytes, so this must be detected as tampering.
    let (encryption_info, mut encrypted_package) = read_encrypted_ooxml_streams(ole);
    encrypted_package.extend_from_slice(&[0xA5u8; 37]);

    let tampered = build_encrypted_ooxml_ole(&encryption_info, &encrypted_package);
    assert_agile_integrity_check_failed(&tampered);
}

#[test]
fn tampered_encrypted_hmac_key_fails_integrity_check() {
    let ole = agile_ole_fixture();

    // Extract the streams and mutate `dataIntegrity/encryptedHmacKey` inside EncryptionInfo,
    // leaving the EncryptedPackage stream unchanged. This should be detected as an integrity error.
    let (mut encryption_info, encrypted_package) = read_encrypted_ooxml_streams(ole);
    let xml_range = agile_encryption_info_xml_range(&encryption_info);
    let xml_bytes = &mut encryption_info[xml_range];
    flip_base64_attr_first_char(xml_bytes, b"encryptedHmacKey=\"");

    let tampered = build_encrypted_ooxml_ole(&encryption_info, &encrypted_package);
    assert_agile_integrity_check_failed(&tampered);
}
