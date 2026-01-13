use formula_offcrypto::{
    decrypt_encrypted_package, decrypt_encrypted_package_ecb, parse_encrypted_package_header,
    parse_encryption_info, validate_agile_segment_decrypt_inputs,
    validate_standard_encrypted_package_stream, DecryptLimits, DecryptOptions, OffcryptoError,
    StandardEncryptionHeader, StandardEncryptionHeaderFlags, StandardEncryptionInfo,
    StandardEncryptionVerifier,
};
use std::io::{Cursor, Read};
use std::path::PathBuf;

fn minimal_standard_encryption_info_bytes() -> Vec<u8> {
    const CALG_AES_128: u32 = 0x0000_660E;
    const CALG_SHA1: u32 = 0x0000_8004;

    let mut bytes = Vec::new();
    bytes.extend_from_slice(&3u16.to_le_bytes());
    bytes.extend_from_slice(&2u16.to_le_bytes());
    bytes.extend_from_slice(&0u32.to_le_bytes()); // flags

    // EncryptionHeader (fixed 8 DWORDs only; CSPName omitted).
    bytes.extend_from_slice(&32u32.to_le_bytes()); // header_size
    bytes.extend_from_slice(
        &(StandardEncryptionHeaderFlags::F_CRYPTOAPI | StandardEncryptionHeaderFlags::F_AES)
            .to_le_bytes(),
    ); // flags
    bytes.extend_from_slice(&0u32.to_le_bytes()); // sizeExtra
    bytes.extend_from_slice(&CALG_AES_128.to_le_bytes()); // algId
    bytes.extend_from_slice(&CALG_SHA1.to_le_bytes()); // algIdHash
    bytes.extend_from_slice(&128u32.to_le_bytes()); // keySize
    bytes.extend_from_slice(&0u32.to_le_bytes()); // providerType
    bytes.extend_from_slice(&0u32.to_le_bytes()); // reserved1
    bytes.extend_from_slice(&0u32.to_le_bytes()); // reserved2

    // EncryptionVerifier
    bytes.extend_from_slice(&16u32.to_le_bytes()); // saltSize
    bytes.extend_from_slice(&[0u8; 16]); // salt
    bytes.extend_from_slice(&[0u8; 16]); // encryptedVerifier
    bytes.extend_from_slice(&20u32.to_le_bytes()); // verifierHashSize (SHA1)
    bytes.extend_from_slice(&[0u8; 32]); // encryptedVerifierHash (padded to 32)

    bytes
}

fn fixture(path: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join(path)
}

fn extract_stream_bytes(cfb_bytes: &[u8], stream_name: &str) -> Vec<u8> {
    let mut ole = cfb::CompoundFile::open(Cursor::new(cfb_bytes)).expect("open cfb");
    let mut stream = ole.open_stream(stream_name).expect("open stream");
    let mut buf = Vec::new();
    stream.read_to_end(&mut buf).expect("read stream");
    buf
}

#[test]
fn truncated_encryption_info_less_than_8_bytes_errors() {
    // EncryptionInfo stream starts with: u16 major, u16 minor, u32 flags (8 bytes total).
    // Provide fewer than 8 bytes and ensure we get a structured error (never panic).
    let bytes = [0u8; 7];
    let err = parse_encryption_info(&bytes).unwrap_err();
    assert!(matches!(err, OffcryptoError::Truncated { .. }));
}

#[test]
fn truncated_encryption_info_at_all_prefix_lengths_errors() {
    let bytes = minimal_standard_encryption_info_bytes();
    for len in 0..bytes.len() {
        let err = parse_encryption_info(&bytes[..len]).unwrap_err();
        assert!(
            matches!(&err, OffcryptoError::Truncated { .. }),
            "len={len} expected Truncated, got {err:?}"
        );
    }
}

#[test]
fn bogus_standard_header_size_is_rejected() {
    for header_size in [0u32, 1u32, 0xFFFF_FFFFu32] {
        let mut bytes = minimal_standard_encryption_info_bytes();
        // header_size is immediately after the 8-byte version+flags prefix.
        bytes[8..12].copy_from_slice(&header_size.to_le_bytes());
        let err = parse_encryption_info(&bytes).unwrap_err();
        match header_size {
            0 | 1 => assert!(
                matches!(err, OffcryptoError::InvalidEncryptionInfo { .. }),
                "header_size={header_size:#x} expected InvalidEncryptionInfo, got {err:?}"
            ),
            _ => assert!(
                matches!(err, OffcryptoError::SizeLimitExceeded { .. }),
                "header_size={header_size:#x} expected SizeLimitExceeded, got {err:?}"
            ),
        }
    }
}

#[test]
fn standard_verify_key_rejects_unaligned_encrypted_verifier_hash() {
    // `encryptedVerifierHash` must be AES-block-aligned (16 bytes) since it is AES-ECB encrypted.
    // Ensure malformed inputs return a structured error rather than panicking.
    let info = StandardEncryptionInfo {
        header: StandardEncryptionHeader {
            flags: StandardEncryptionHeaderFlags::from_raw(
                StandardEncryptionHeaderFlags::F_CRYPTOAPI | StandardEncryptionHeaderFlags::F_AES,
            ),
            size_extra: 0,
            alg_id: 0x0000_660E,
            alg_id_hash: 0x0000_8004,
            key_size_bits: 128,
            provider_type: 0,
            reserved1: 0,
            reserved2: 0,
            csp_name: String::new(),
        },
        verifier: StandardEncryptionVerifier {
            salt: vec![0u8; 16],
            encrypted_verifier: [0u8; 16],
            verifier_hash_size: 20,
            encrypted_verifier_hash: vec![0u8; 31], // NOT a multiple of 16
        },
    };

    let err = formula_offcrypto::standard_verify_key(&info, &[0u8; 16]).unwrap_err();
    assert_eq!(err, OffcryptoError::InvalidCiphertextLength { len: 31 });
}

#[test]
fn agile_header_ok_but_xml_missing_required_attrs_errors() {
    // Minimal Agile header:
    // - major=4, minor=4, flags=0
    // - XML with <keyData> but missing most required attributes.
    let xml = br#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"><keyData saltSize="16"/></encryption>"#;

    let mut bytes = Vec::new();
    bytes.extend_from_slice(&4u16.to_le_bytes());
    bytes.extend_from_slice(&4u16.to_le_bytes());
    bytes.extend_from_slice(&0u32.to_le_bytes()); // flags
    bytes.extend_from_slice(xml);

    let err = parse_encryption_info(&bytes).unwrap_err();
    assert!(matches!(
        err,
        OffcryptoError::InvalidEncryptionInfo { .. }
    ));
}

#[test]
fn standard_header_encryption_header_size_larger_than_buffer_errors() {
    // Standard (3.2) header, but with a header_size that exceeds the available bytes.
    let mut bytes = Vec::new();
    bytes.extend_from_slice(&3u16.to_le_bytes());
    bytes.extend_from_slice(&2u16.to_le_bytes());
    bytes.extend_from_slice(&0u32.to_le_bytes()); // flags
    bytes.extend_from_slice(&100u32.to_le_bytes()); // header_size (too large for empty remainder)

    let err = parse_encryption_info(&bytes).unwrap_err();
    assert!(matches!(err, OffcryptoError::Truncated { .. }));
}

#[test]
fn encrypted_package_shorter_than_8_bytes_errors() {
    let err = parse_encrypted_package_header(&[0u8; 7]).unwrap_err();
    assert!(matches!(err, OffcryptoError::Truncated { .. }));
}

#[test]
fn encrypted_package_header_falls_back_to_low_dword_when_high_dword_is_reserved() {
    // Some producers treat the 8-byte size prefix as (u32 totalSize, u32 reserved). Ensure we
    // tolerate a non-zero "reserved" high DWORD when it is not plausible for the ciphertext.
    let mut encrypted_package = Vec::new();
    encrypted_package.extend_from_slice(&1234u32.to_le_bytes()); // size (low DWORD)
    encrypted_package.extend_from_slice(&1u32.to_le_bytes()); // reserved (high DWORD)
    encrypted_package.extend_from_slice(&[0u8; 2048]); // ciphertext (enough to cover low DWORD)

    let header = parse_encrypted_package_header(&encrypted_package).expect("parse header");
    assert_eq!(header.original_size, 1234);
}

#[test]
fn standard_encrypted_package_ciphertext_not_multiple_of_16_errors() {
    // Ciphertext length after the 8-byte original-size prefix must be block-aligned.
    let mut encrypted_package = 0u64.to_le_bytes().to_vec();
    encrypted_package.extend_from_slice(&[0u8; 15]); // not a multiple of 16

    let err = validate_standard_encrypted_package_stream(&encrypted_package).unwrap_err();
    assert_eq!(err, OffcryptoError::InvalidCiphertextLength { len: 15 });
}

#[test]
fn decrypt_encrypted_package_standard_rejects_short_encrypted_package_before_password_check() {
    let encryption_info = minimal_standard_encryption_info_bytes();
    let encrypted_package = [0u8; 7];
    let err = decrypt_encrypted_package(
        &encryption_info,
        &encrypted_package,
        "wrong-password",
        DecryptOptions::default(),
    )
    .unwrap_err();
    assert!(
        matches!(err, OffcryptoError::Truncated { context } if context == "EncryptedPackageHeader.original_size"),
        "expected Truncated(original_size), got {err:?}"
    );
}

#[test]
fn decrypt_encrypted_package_standard_rejects_unaligned_ciphertext_before_password_check() {
    let encryption_info = minimal_standard_encryption_info_bytes();
    let mut encrypted_package = 0u64.to_le_bytes().to_vec();
    encrypted_package.extend_from_slice(&[0u8; 15]);
    let err = decrypt_encrypted_package(
        &encryption_info,
        &encrypted_package,
        "wrong-password",
        DecryptOptions::default(),
    )
    .unwrap_err();
    assert_eq!(err, OffcryptoError::InvalidCiphertextLength { len: 15 });
}

#[test]
fn decrypt_encrypted_package_standard_rejects_size_mismatch_before_password_check() {
    // total_size=32 requires at least 32 bytes of ciphertext (ciphertext is also padded to 16-byte
    // blocks). Provide only 16 bytes of ciphertext.
    let encryption_info = minimal_standard_encryption_info_bytes();
    let mut encrypted_package = 32u64.to_le_bytes().to_vec();
    encrypted_package.extend_from_slice(&[0u8; 16]);
    let err = decrypt_encrypted_package(
        &encryption_info,
        &encrypted_package,
        "wrong-password",
        DecryptOptions::default(),
    )
    .unwrap_err();
    assert_eq!(
        err,
        OffcryptoError::EncryptedPackageSizeMismatch {
            total_size: 32,
            ciphertext_len: 16
        }
    );
}

#[test]
fn decrypt_encrypted_package_agile_rejects_short_encrypted_package_before_password_check() {
    let encrypted = std::fs::read(fixture("inputs/example_password.xlsx")).expect("read fixture");
    let encryption_info = extract_stream_bytes(&encrypted, "EncryptionInfo");
    let encrypted_package = [0u8; 7];
    let err = decrypt_encrypted_package(
        &encryption_info,
        &encrypted_package,
        "wrong-password",
        DecryptOptions::default(),
    )
    .unwrap_err();
    assert!(
        matches!(err, OffcryptoError::Truncated { context } if context == "EncryptedPackageHeader.original_size"),
        "expected Truncated(original_size), got {err:?}"
    );
}

#[test]
fn decrypt_encrypted_package_agile_rejects_unaligned_ciphertext_before_password_check() {
    let encrypted = std::fs::read(fixture("inputs/example_password.xlsx")).expect("read fixture");
    let encryption_info = extract_stream_bytes(&encrypted, "EncryptionInfo");
    let mut encrypted_package = 0u64.to_le_bytes().to_vec();
    encrypted_package.extend_from_slice(&[0u8; 15]);
    let err = decrypt_encrypted_package(
        &encryption_info,
        &encrypted_package,
        "wrong-password",
        DecryptOptions::default(),
    )
    .unwrap_err();
    assert_eq!(err, OffcryptoError::InvalidCiphertextLength { len: 15 });
}

#[test]
fn decrypt_encrypted_package_agile_rejects_size_mismatch_before_password_check() {
    // total_size=32 requires at least 32 bytes of ciphertext (ciphertext is also padded to 16-byte
    // blocks). Provide only 16 bytes of ciphertext.
    let encrypted = std::fs::read(fixture("inputs/example_password.xlsx")).expect("read fixture");
    let encryption_info = extract_stream_bytes(&encrypted, "EncryptionInfo");
    let mut encrypted_package = 32u64.to_le_bytes().to_vec();
    encrypted_package.extend_from_slice(&[0u8; 16]);
    let err = decrypt_encrypted_package(
        &encryption_info,
        &encrypted_package,
        "wrong-password",
        DecryptOptions::default(),
    )
    .unwrap_err();
    assert_eq!(
        err,
        OffcryptoError::EncryptedPackageSizeMismatch {
            total_size: 32,
            ciphertext_len: 16
        }
    );
}

#[test]
fn decrypt_encrypted_package_standard_rejects_output_too_large_before_password_check() {
    let encryption_info = minimal_standard_encryption_info_bytes();

    let total_size: u64 = 2 * 1024 * 1024; // 2MiB
    let max: u64 = 1024 * 1024; // 1MiB

    let mut encrypted_package = Vec::new();
    encrypted_package.extend_from_slice(&total_size.to_le_bytes());
    encrypted_package.resize(8 + total_size as usize, 0);

    let options = DecryptOptions {
        verify_integrity: false,
        limits: DecryptLimits {
            max_output_size: Some(max),
            ..Default::default()
        },
    };

    let err = decrypt_encrypted_package(&encryption_info, &encrypted_package, "wrong-password", options)
        .unwrap_err();
    assert_eq!(
        err,
        OffcryptoError::OutputTooLarge {
            total_size,
            max,
        }
    );
}

#[test]
fn decrypt_encrypted_package_agile_rejects_output_too_large_before_password_check() {
    let encrypted = std::fs::read(fixture("inputs/example_password.xlsx")).expect("read fixture");
    let encryption_info = extract_stream_bytes(&encrypted, "EncryptionInfo");

    let total_size: u64 = 2 * 1024 * 1024; // 2MiB
    let max: u64 = 1024 * 1024; // 1MiB

    let mut encrypted_package = Vec::new();
    encrypted_package.extend_from_slice(&total_size.to_le_bytes());
    encrypted_package.resize(8 + total_size as usize, 0);

    let options = DecryptOptions {
        verify_integrity: false,
        limits: DecryptLimits {
            max_output_size: Some(max),
            ..Default::default()
        },
    };

    let err = decrypt_encrypted_package(&encryption_info, &encrypted_package, "wrong-password", options)
        .unwrap_err();
    assert_eq!(
        err,
        OffcryptoError::OutputTooLarge {
            total_size,
            max,
        }
    );
}

#[test]
fn standard_decrypt_encrypted_package_ecb_short_header_errors() {
    let err = decrypt_encrypted_package_ecb(&[0u8; 16], &[0u8; 7]).unwrap_err();
    assert!(
        matches!(err, OffcryptoError::InvalidStructure(ref msg) if msg.contains("must be at least 8 bytes")),
        "expected InvalidStructure(short EncryptedPackage), got {err:?}"
    );
}

#[test]
fn standard_decrypt_encrypted_package_ecb_ciphertext_not_block_aligned_errors() {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(&0u64.to_le_bytes());
    bytes.extend_from_slice(&[0u8; 15]);
    let err = decrypt_encrypted_package_ecb(&[0u8; 16], &bytes).unwrap_err();
    assert!(
        matches!(err, OffcryptoError::InvalidStructure(ref msg) if msg.contains("ciphertext length must be a multiple of 16")),
        "expected InvalidStructure(un-aligned ciphertext), got {err:?}"
    );
}

#[test]
fn standard_decrypt_encrypted_package_ecb_original_size_exceeds_plaintext_errors() {
    // 1 AES block of ciphertext => 16 bytes of plaintext after decrypt.
    let mut bytes = Vec::new();
    bytes.extend_from_slice(&17u64.to_le_bytes()); // original_size > plaintext length
    bytes.extend_from_slice(&[0u8; 16]);
    let err = decrypt_encrypted_package_ecb(&[0u8; 16], &bytes).unwrap_err();
    assert!(
        matches!(err, OffcryptoError::InvalidStructure(ref msg) if msg.contains("exceeds plaintext length")),
        "expected InvalidStructure(original size > plaintext len), got {err:?}"
    );
}

#[test]
fn agile_segment_decrypt_wrong_lengths_errors() {
    // expected_plaintext_len=17 implies at least 32 bytes of ciphertext for block alignment.
    let iv = [0u8; 16];
    let ciphertext = [0u8; 16];

    let err = validate_agile_segment_decrypt_inputs(&iv, &ciphertext, 17).unwrap_err();
    assert!(matches!(err, OffcryptoError::InvalidEncryptionInfo { .. }));
}
