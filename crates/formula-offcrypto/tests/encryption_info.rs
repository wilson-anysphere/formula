use formula_offcrypto::{
    parse_encryption_info, EncryptionInfo, OffcryptoError, StandardEncryptionHeader,
    StandardEncryptionVerifier,
};

const CALG_AES_128: u32 = 0x0000_660E;
const CALG_SHA1: u32 = 0x0000_8004;

fn utf16le_bytes(s: &str, terminated: bool) -> Vec<u8> {
    let mut out = Vec::new();
    for cu in s.encode_utf16() {
        out.extend_from_slice(&cu.to_le_bytes());
    }
    if terminated {
        out.extend_from_slice(&0u16.to_le_bytes());
    }
    out
}

fn build_standard_encryption_info_with_version(
    version_major: u16,
    version_minor: u16,
    csp_name: &[u8],
    alg_id: u32,
    alg_id_hash: u32,
    key_size_bits: u32,
    salt_size: u32,
    verifier_hash_size: u32,
    encrypted_verifier_hash_len: usize,
) -> Vec<u8> {
    let mut bytes = Vec::new();

    // EncryptionVersionInfo
    bytes.extend_from_slice(&version_major.to_le_bytes());
    bytes.extend_from_slice(&version_minor.to_le_bytes());
    bytes.extend_from_slice(&0xAABBCCDDu32.to_le_bytes());

    // EncryptionHeader (8 DWORDs + cspName)
    let mut header = Vec::new();
    header.extend_from_slice(&0x11111111u32.to_le_bytes()); // flags
    header.extend_from_slice(&0x22222222u32.to_le_bytes()); // sizeExtra
    header.extend_from_slice(&alg_id.to_le_bytes()); // algId
    header.extend_from_slice(&alg_id_hash.to_le_bytes()); // algIdHash
    header.extend_from_slice(&key_size_bits.to_le_bytes()); // keySize
    header.extend_from_slice(&0x66666666u32.to_le_bytes()); // providerType
    header.extend_from_slice(&0x77777777u32.to_le_bytes()); // reserved1
    header.extend_from_slice(&0x88888888u32.to_le_bytes()); // reserved2
    header.extend_from_slice(csp_name);
    bytes.extend_from_slice(&(header.len() as u32).to_le_bytes());
    bytes.extend_from_slice(&header);

    // EncryptionVerifier
    bytes.extend_from_slice(&salt_size.to_le_bytes()); // saltSize
    bytes.extend((1u8..).take(salt_size as usize)); // salt bytes
    bytes.extend_from_slice(&[0xAA; 16]); // encryptedVerifier
    bytes.extend_from_slice(&verifier_hash_size.to_le_bytes()); // verifierHashSize
    bytes.extend(std::iter::repeat(0xBBu8).take(encrypted_verifier_hash_len)); // encryptedVerifierHash

    bytes
}

fn build_standard_encryption_info(
    csp_name: &[u8],
    alg_id: u32,
    alg_id_hash: u32,
    key_size_bits: u32,
    salt_size: u32,
    verifier_hash_size: u32,
    encrypted_verifier_hash_len: usize,
) -> Vec<u8> {
    build_standard_encryption_info_with_version(
        3,
        2,
        csp_name,
        alg_id,
        alg_id_hash,
        key_size_bits,
        salt_size,
        verifier_hash_size,
        encrypted_verifier_hash_len,
    )
}

#[test]
fn parse_synthetic_standard_encryption_info() {
    let csp_name = utf16le_bytes("Test CSP", true);
    let bytes = build_standard_encryption_info(&csp_name, CALG_AES_128, CALG_SHA1, 128, 16, 20, 32);

    let info = parse_encryption_info(&bytes).expect("parse");
    let EncryptionInfo::Standard {
        version,
        header,
        verifier,
    } = info
    else {
        panic!("expected standard");
    };

    assert_eq!(version.major, 3);
    assert_eq!(version.minor, 2);
    assert_eq!(version.flags, 0xAABBCCDD);

    assert_eq!(
        header,
        StandardEncryptionHeader {
            flags: 0x11111111,
            size_extra: 0x22222222,
            alg_id: CALG_AES_128,
            alg_id_hash: CALG_SHA1,
            key_size_bits: 128,
            provider_type: 0x66666666,
            reserved1: 0x77777777,
            reserved2: 0x88888888,
            csp_name: "Test CSP".to_string(),
        }
    );

    assert_eq!(
        verifier,
        StandardEncryptionVerifier {
            salt: (1u8..=16).collect(),
            encrypted_verifier: [0xAA; 16],
            verifier_hash_size: 20,
            encrypted_verifier_hash: vec![0xBB; 32],
        }
    );
}

#[test]
fn parse_synthetic_standard_encryption_info_accepts_major_2_minor_2() {
    let bytes = build_standard_encryption_info_with_version(
        2,
        2,
        &utf16le_bytes("Test CSP", true),
        CALG_AES_128,
        CALG_SHA1,
        128,
        16,
        20,
        32,
    );
    let info = parse_encryption_info(&bytes).expect("parse");
    let EncryptionInfo::Standard { version, .. } = info else {
        panic!("expected standard");
    };
    assert_eq!((version.major, version.minor), (2, 2));
}

#[test]
fn parse_synthetic_standard_encryption_info_accepts_major_4_minor_2() {
    let bytes = build_standard_encryption_info_with_version(
        4,
        2,
        &utf16le_bytes("Test CSP", true),
        CALG_AES_128,
        CALG_SHA1,
        128,
        16,
        20,
        32,
    );
    let info = parse_encryption_info(&bytes).expect("parse");
    let EncryptionInfo::Standard { version, .. } = info else {
        panic!("expected standard");
    };
    assert_eq!((version.major, version.minor), (4, 2));
}

#[test]
fn truncation_missing_header_size() {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(&3u16.to_le_bytes());
    bytes.extend_from_slice(&2u16.to_le_bytes());
    bytes.extend_from_slice(&0u32.to_le_bytes());

    let err = parse_encryption_info(&bytes).unwrap_err();
    assert!(matches!(err, OffcryptoError::Truncated { .. }));
}

#[test]
fn truncation_header_shorter_than_fixed_fields() {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(&3u16.to_le_bytes());
    bytes.extend_from_slice(&2u16.to_le_bytes());
    bytes.extend_from_slice(&0u32.to_le_bytes());

    // header_size is 16, but a valid header needs at least 32.
    bytes.extend_from_slice(&16u32.to_le_bytes());
    bytes.extend_from_slice(&[0u8; 16]);

    let err = parse_encryption_info(&bytes).unwrap_err();
    assert!(matches!(
        err,
        OffcryptoError::InvalidEncryptionInfo { .. }
    ));
}

#[test]
fn truncation_missing_verifier_fields() {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(&3u16.to_le_bytes());
    bytes.extend_from_slice(&2u16.to_le_bytes());
    bytes.extend_from_slice(&0u32.to_le_bytes());

    // Header with fixed fields only (no CSPName, ok).
    bytes.extend_from_slice(&32u32.to_le_bytes());
    bytes.extend_from_slice(&0u32.to_le_bytes()); // flags
    bytes.extend_from_slice(&0u32.to_le_bytes()); // sizeExtra
    bytes.extend_from_slice(&CALG_AES_128.to_le_bytes()); // algId
    bytes.extend_from_slice(&CALG_SHA1.to_le_bytes()); // algIdHash
    bytes.extend_from_slice(&128u32.to_le_bytes()); // keySize
    bytes.extend_from_slice(&0u32.to_le_bytes()); // providerType
    bytes.extend_from_slice(&0u32.to_le_bytes()); // reserved1
    bytes.extend_from_slice(&0u32.to_le_bytes()); // reserved2

    // Verifier truncated: only saltSize present, but no salt/verifier fields.
    // Use a valid saltSize (16) so the parser attempts to read the missing bytes.
    bytes.extend_from_slice(&16u32.to_le_bytes());

    let err = parse_encryption_info(&bytes).unwrap_err();
    assert!(matches!(err, OffcryptoError::Truncated { .. }));
}

#[test]
fn csp_name_accepts_terminated_and_non_terminated_utf16le() {
    let bytes_term =
        build_standard_encryption_info(&utf16le_bytes("CSP", true), CALG_AES_128, CALG_SHA1, 128, 16, 20, 32);
    let info = parse_encryption_info(&bytes_term).expect("terminated parse");
    let EncryptionInfo::Standard { header, .. } = info else {
        panic!("expected standard");
    };
    assert_eq!(header.csp_name, "CSP");

    let bytes_no_term =
        build_standard_encryption_info(&utf16le_bytes("CSP", false), CALG_AES_128, CALG_SHA1, 128, 16, 20, 32);
    let info = parse_encryption_info(&bytes_no_term).expect("non-terminated parse");
    let EncryptionInfo::Standard { header, .. } = info else {
        panic!("expected standard");
    };
    assert_eq!(header.csp_name, "CSP");
}

#[test]
fn csp_name_rejects_invalid_utf16() {
    // Unpaired surrogate.
    let bad = 0xD800u16.to_le_bytes();
    let bytes = build_standard_encryption_info(&bad, CALG_AES_128, CALG_SHA1, 128, 16, 20, 32);
    let err = parse_encryption_info(&bytes).unwrap_err();
    assert!(matches!(err, OffcryptoError::InvalidCspNameUtf16));
}

#[test]
fn rejects_key_size_mismatch_for_aes() {
    let bytes = build_standard_encryption_info(
        &utf16le_bytes("CSP", true),
        CALG_AES_128,
        CALG_SHA1,
        256, // wrong for AES-128
        16,
        20,
        32,
    );
    let err = parse_encryption_info(&bytes).unwrap_err();
    assert_eq!(err, OffcryptoError::UnsupportedAlgorithm(CALG_AES_128));
}

#[test]
fn rejects_non_sha1_alg_id_hash() {
    let bytes = build_standard_encryption_info(
        &utf16le_bytes("CSP", true),
        CALG_AES_128,
        0x0000_800C, // CALG_SHA_256
        128,
        16,
        20,
        32,
    );
    let err = parse_encryption_info(&bytes).unwrap_err();
    assert_eq!(err, OffcryptoError::UnsupportedAlgorithm(0x0000_800C));
}

#[test]
fn rejects_verifier_salt_size_mismatch() {
    let bytes = build_standard_encryption_info(
        &utf16le_bytes("CSP", true),
        CALG_AES_128,
        CALG_SHA1,
        128,
        8, // invalid for Standard AES verifier
        20,
        32,
    );
    let err = parse_encryption_info(&bytes).unwrap_err();
    assert_eq!(
        err,
        OffcryptoError::InvalidEncryptionInfo {
            context: "EncryptionVerifier.saltSize must be 16 for Standard encryption"
        }
    );
}

#[test]
fn rejects_truncated_encrypted_verifier_hash() {
    // verifierHashSize says SHA1 (20 bytes) => requires 32 bytes of encrypted hash, but provide 16.
    let bytes = build_standard_encryption_info(
        &utf16le_bytes("CSP", true),
        CALG_AES_128,
        CALG_SHA1,
        128,
        16,
        20,
        16,
    );
    let err = parse_encryption_info(&bytes).unwrap_err();
    assert!(matches!(err, OffcryptoError::Truncated { .. }));
}

#[test]
fn rejects_verifier_hash_size_mismatch() {
    let bytes = build_standard_encryption_info(
        &utf16le_bytes("CSP", true),
        CALG_AES_128,
        CALG_SHA1,
        128,
        16,
        32, // not SHA1
        32,
    );
    let err = parse_encryption_info(&bytes).unwrap_err();
    assert_eq!(
        err,
        OffcryptoError::InvalidEncryptionInfo {
            context:
                "EncryptionVerifier.verifierHashSize must be 20 (SHA1) for Standard encryption"
        }
    );
}

#[test]
fn rejects_unsupported_standard_alg_id() {
    let bytes = build_standard_encryption_info(
        &utf16le_bytes("CSP", true),
        0xDEAD_BEEF,
        CALG_SHA1,
        128,
        16,
        20,
        32,
    );
    let err = parse_encryption_info(&bytes).unwrap_err();
    assert_eq!(err, OffcryptoError::UnsupportedAlgorithm(0xDEAD_BEEF));
}

#[test]
fn truncation_missing_encrypted_verifier_bytes() {
    let mut bytes = build_standard_encryption_info(
        &utf16le_bytes("CSP", true),
        CALG_AES_128,
        CALG_SHA1,
        128,
        16,
        20,
        32,
    );

    // Truncate halfway through the encryptedVerifier field (16 bytes).
    let encrypted_verifier_offset = bytes.len() - (16 + 4 + 32);
    bytes.truncate(encrypted_verifier_offset + 8);

    let err = parse_encryption_info(&bytes).unwrap_err();
    assert!(matches!(err, OffcryptoError::Truncated { .. }));
}
