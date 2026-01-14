use formula_office_crypto as office_crypto;

/// Build a Standard (CryptoAPI) MS-OFFCRYPTO encrypted OOXML OLE/CFB wrapper around the provided
/// plaintext ZIP bytes.
///
/// The `formula-xlsb` CLIs (`xlsb_dump`, `rgce_coverage`) decrypt encrypted OLE wrappers using
/// `formula-office-crypto`, so tests should generate fixtures using the same implementation.
pub fn build_standard_encrypted_ooxml_ole_bytes(package_bytes: &[u8], password: &str) -> Vec<u8> {
    office_crypto::encrypt_package_to_ole(
        package_bytes,
        password,
        office_crypto::EncryptOptions {
            scheme: office_crypto::EncryptionScheme::Standard,
            // Standard encryption in `formula-office-crypto` is intentionally Excel-compatible and
            // currently supports SHA1 only (CryptoAPI semantics) with a fixed spin count.
            key_bits: 128,
            hash_algorithm: office_crypto::HashAlgorithm::Sha1,
            spin_count: 50_000,
        },
    )
    .expect("encrypt standard OOXML OLE wrapper")
}

