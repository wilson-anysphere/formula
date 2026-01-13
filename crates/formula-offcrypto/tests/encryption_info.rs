use formula_offcrypto::{
    parse_encryption_info, EncryptionInfo, OffcryptoError, StandardEncryptionHeader,
    StandardEncryptionVerifier,
};

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

fn build_standard_encryption_info(csp_name: &[u8]) -> Vec<u8> {
    let mut bytes = Vec::new();

    // EncryptionVersionInfo (major=3, minor=2)
    bytes.extend_from_slice(&3u16.to_le_bytes());
    bytes.extend_from_slice(&2u16.to_le_bytes());
    bytes.extend_from_slice(&0xAABBCCDDu32.to_le_bytes());

    // EncryptionHeader (8 DWORDs + cspName)
    let mut header = Vec::new();
    header.extend_from_slice(&0x11111111u32.to_le_bytes()); // flags
    header.extend_from_slice(&0x22222222u32.to_le_bytes()); // sizeExtra
    header.extend_from_slice(&0x33333333u32.to_le_bytes()); // algId
    header.extend_from_slice(&0x44444444u32.to_le_bytes()); // algIdHash
    header.extend_from_slice(&0x55555555u32.to_le_bytes()); // keySize
    header.extend_from_slice(&0x66666666u32.to_le_bytes()); // providerType
    header.extend_from_slice(&0x77777777u32.to_le_bytes()); // reserved1
    header.extend_from_slice(&0x88888888u32.to_le_bytes()); // reserved2
    header.extend_from_slice(csp_name);
    bytes.extend_from_slice(&(header.len() as u32).to_le_bytes());
    bytes.extend_from_slice(&header);

    // EncryptionVerifier
    bytes.extend_from_slice(&8u32.to_le_bytes()); // saltSize
    bytes.extend_from_slice(&[1, 2, 3, 4, 5, 6, 7, 8]); // salt
    bytes.extend_from_slice(&[0xAA; 16]); // encryptedVerifier
    bytes.extend_from_slice(&20u32.to_le_bytes()); // verifierHashSize
    bytes.extend_from_slice(&[0xBB; 32]); // encryptedVerifierHash (remaining bytes)

    bytes
}

#[test]
fn parse_synthetic_standard_encryption_info() {
    let csp_name = utf16le_bytes("Test CSP", true);
    let bytes = build_standard_encryption_info(&csp_name);

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
            alg_id: 0x33333333,
            alg_id_hash: 0x44444444,
            key_size_bits: 0x55555555,
            provider_type: 0x66666666,
            reserved1: 0x77777777,
            reserved2: 0x88888888,
            csp_name: "Test CSP".to_string(),
        }
    );

    assert_eq!(
        verifier,
        StandardEncryptionVerifier {
            salt: vec![1, 2, 3, 4, 5, 6, 7, 8],
            encrypted_verifier: [0xAA; 16],
            verifier_hash_size: 20,
            encrypted_verifier_hash: vec![0xBB; 32],
        }
    );
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
    assert!(matches!(err, OffcryptoError::Truncated { .. }));
}

#[test]
fn truncation_missing_verifier_fields() {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(&3u16.to_le_bytes());
    bytes.extend_from_slice(&2u16.to_le_bytes());
    bytes.extend_from_slice(&0u32.to_le_bytes());

    // Header with fixed fields only (no CSPName, ok).
    bytes.extend_from_slice(&32u32.to_le_bytes());
    bytes.extend_from_slice(&[0u8; 32]);

    // Verifier truncated: only saltSize present, but no salt/verifier fields.
    bytes.extend_from_slice(&8u32.to_le_bytes());

    let err = parse_encryption_info(&bytes).unwrap_err();
    assert!(matches!(err, OffcryptoError::Truncated { .. }));
}

#[test]
fn csp_name_accepts_terminated_and_non_terminated_utf16le() {
    let bytes_term = build_standard_encryption_info(&utf16le_bytes("CSP", true));
    let info = parse_encryption_info(&bytes_term).expect("terminated parse");
    let EncryptionInfo::Standard { header, .. } = info else {
        panic!("expected standard");
    };
    assert_eq!(header.csp_name, "CSP");

    let bytes_no_term = build_standard_encryption_info(&utf16le_bytes("CSP", false));
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
    let bytes = build_standard_encryption_info(&bad);
    let err = parse_encryption_info(&bytes).unwrap_err();
    assert!(matches!(err, OffcryptoError::InvalidCspNameUtf16));
}

