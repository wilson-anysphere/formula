use formula_offcrypto::{agile_secret_key, AgileEncryptionInfo, HashAlgorithm};

#[test]
fn agile_secret_key_matches_msoffcrypto_docstring_vector() {
    // Test vector from:
    // https://github.com/nolze/msoffcrypto-tool/blob/master/msoffcrypto/method/ecma376_agile.py
    let password = "Password1234_";
    let info = AgileEncryptionInfo {
        key_data_salt: Vec::new(),
        key_data_hash_algorithm: HashAlgorithm::Sha512,
        key_data_block_size: 16,
        encrypted_hmac_key: Vec::new(),
        encrypted_hmac_value: Vec::new(),
        spin_count: 100_000,
        password_salt: hex::decode("4c725d45dc610f939412a04da7910466").unwrap(),
        password_hash_algorithm: HashAlgorithm::Sha512,
        password_key_bits: 256,
        encrypted_key_value: hex::decode(
            "a16cd5165a7ab9d271113ed386a78cf49692e8e527b0c5fc0055ed080b7cb94b",
        )
        .unwrap(),
        encrypted_verifier_hash_input: Vec::new(),
        encrypted_verifier_hash_value: Vec::new(),
    };

    let secret_key = agile_secret_key(&info, password).unwrap();
    let expected =
        hex::decode("40206609d9faadf24b076aebf2c435b74292c8b8a7aa81bc679be89711b02ac2").unwrap();
    assert_eq!(secret_key.as_slice(), expected.as_slice());
}
