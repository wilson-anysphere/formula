use formula_office_crypto as office_crypto;

/// Build an Agile (ECMA-376) MS-OFFCRYPTO encrypted OOXML OLE/CFB wrapper around the provided
/// plaintext ZIP bytes.
///
/// We keep the crypto parameters intentionally small so CLI integration tests run quickly.
pub fn build_agile_encrypted_ooxml_ole_bytes(package_bytes: &[u8], password: &str) -> Vec<u8> {
    office_crypto::encrypt_package_to_ole(
        package_bytes,
        password,
        office_crypto::EncryptOptions {
            scheme: office_crypto::EncryptionScheme::Agile,
            key_bits: 128,
            hash_algorithm: office_crypto::HashAlgorithm::Sha1,
            // Keep this small for test speed; `formula-office-crypto` enforces max spin_count
            // during decryption as a DoS guard.
            spin_count: 1_000,
        },
    )
    .expect("encrypt agile OOXML OLE wrapper")
}

