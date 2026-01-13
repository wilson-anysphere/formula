use formula_offcrypto::{decrypt_encrypted_package, parse_encrypted_package_header, DecryptOptions, OffcryptoError};

#[test]
fn encrypted_package_header_truncated() {
    let bytes = [0u8; 7];
    let err = parse_encrypted_package_header(&bytes).unwrap_err();
    assert!(
        matches!(err, OffcryptoError::Truncated { .. }),
        "err={err:?}"
    );
}

#[test]
fn encrypted_package_header_size_too_large_is_rejected_by_default_decrypt_limits() {
    let too_large = formula_offcrypto::MAX_ENCRYPTED_PACKAGE_ORIGINAL_SIZE + 1;
    let mut bytes = Vec::new();
    bytes.extend_from_slice(&too_large.to_le_bytes());
    let err = decrypt_encrypted_package(&[], &bytes, "", DecryptOptions::default()).unwrap_err();
    assert!(
        matches!(
            err,
            OffcryptoError::OutputTooLarge { total_size, max }
                if total_size == too_large && max == formula_offcrypto::MAX_ENCRYPTED_PACKAGE_ORIGINAL_SIZE
        ),
        "err={err:?}"
    );
}
